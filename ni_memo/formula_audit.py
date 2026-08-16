"""Formula-population and cache-state gates, separate from business validation."""
from dataclasses import asdict, dataclass
from pathlib import Path
import re
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl.utils.cell import get_column_letter, range_boundaries


EXCEL_ERRORS = {
    "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!"
}
EXTERNAL_REFERENCE = re.compile(r"\[[^\]]+\]")


def count_formula_nodes(paths):
    """Count worksheet formula nodes from raw OOXML, independent of openpyxl."""
    workbooks = []
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            workbooks.extend(sorted(path.rglob("*.xlsx")))
        elif path.suffix.lower() == ".xlsx":
            workbooks.append(path)
    total = 0
    for workbook in workbooks:
        with ZipFile(workbook) as archive:
            sheets = sorted(name for name in archive.namelist()
                            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"))
            for sheet in sheets:
                root = ElementTree.fromstring(archive.read(sheet))
                total += sum(node.tag.rsplit("}", 1)[-1] == "f" for node in root.iter())
    return total


@dataclass(frozen=True)
class FormulaAudit:
    status: str
    total_formulas: int
    cached_ok: int
    cached_blank: tuple[str, ...]
    cached_errors: tuple[str, ...]
    not_applicable_errors: tuple[str, ...]
    broken_references: tuple[str, ...]
    external_references: tuple[str, ...]
    issue_keys: tuple[str, ...]
    confirmed_intrinsic_issues: tuple[str, ...]
    selected_fact_issues: tuple[str, ...]
    unexplained_omissions: int
    recalculation_status: str

    def to_dict(self):
        data = asdict(self)
        for key in ("cached_blank", "cached_errors", "not_applicable_errors",
                    "broken_references", "external_references", "issue_keys",
                    "confirmed_intrinsic_issues", "selected_fact_issues"):
            data[key] = list(data[key])
        return data


def audit_formula_evidence(items, expected_formula_count=None, recalculation_status="not_run",
                           confirmed_intrinsic_issues=(), selected_fact_sources=(),
                           source_identities=None):
    formulas = sorted((item for item in items if item.formula), key=lambda item: (
        item.source_uri.lower(), item.source_location
    ))
    total = len(formulas)
    if total == 0:
        return FormulaAudit(
            "not_applicable", 0, 0, (), (), (), (), (), (), (), (), 0, "not_applicable",
        )

    blank = tuple(item.source_location for item in formulas if item.cached_value is None)
    index = _evidence_index(items)
    error_items = [item for item in formulas if str(item.cached_value).upper() in EXCEL_ERRORS]
    not_applicable = tuple(item.source_location for item in error_items if _evidenced_not_applicable(item, index))
    errors = tuple(item.source_location for item in error_items if item.source_location not in not_applicable)
    broken = tuple(item.source_location for item in formulas if "#REF!" in str(item.formula).upper())
    external = tuple(item.source_location for item in formulas if EXTERNAL_REFERENCE.search(str(item.formula)))
    cached_ok = sum(item.cached_value is not None and str(item.cached_value).upper() not in EXCEL_ERRORS
                    for item in formulas)
    omissions = max(0, (expected_formula_count if expected_formula_count is not None else total) - total)

    error_locations = set(errors)
    issue_keys = tuple(sorted({
        formula_issue_key(
            _source_identity(item.source_uri, source_identities), item.source_location,
        )
        for item in formulas
        if item.source_location in error_locations or "#REF!" in str(item.formula).upper()
    }))
    supplied_confirmed = set(confirmed_intrinsic_issues)
    confirmed = tuple(sorted(
        set(issue_keys) & supplied_confirmed
        if recalculation_status in {"pass", "warning"} else set()
    ))
    selected = tuple(sorted(set(issue_keys) & set(selected_fact_sources)))
    unconfirmed = set(issue_keys) - set(confirmed)

    if recalculation_status == "fail":
        status = "fail"
    elif omissions or selected or unconfirmed:
        status = "fail"
    elif issue_keys:
        status = "warning"
    elif blank or external or recalculation_status != "pass":
        status = "warning"
    else:
        status = "pass"
    return FormulaAudit(
        status, total, cached_ok, blank, errors, not_applicable, broken, external,
        issue_keys, confirmed, selected, omissions, recalculation_status,
    )


def formula_issue_key(source_uri, source_location):
    source = str(Path(source_uri).resolve()).casefold()
    return f"{source}::{source_location}"


def _source_identity(source_uri, identities):
    if not identities:
        return source_uri
    key = str(Path(source_uri).resolve()).casefold()
    return identities.get(key, source_uri)


def _evidence_index(items):
    index = {}
    for item in items:
        if "!" not in item.source_location:
            continue
        sheet, coordinate = item.source_location.rsplit("!", 1)
        if re.fullmatch(r"[A-Za-z]+\d+", coordinate):
            index[(item.source_uri, sheet, coordinate.upper())] = item
    return index


def _evidenced_not_applicable(item, index):
    if "!" not in item.source_location:
        return False
    sheet, _ = item.source_location.rsplit("!", 1)
    expression = str(item.formula).replace("$", "").replace(" ", "").upper()
    cached_error = str(item.cached_value).upper()
    if "XIRR(" in expression and cached_error in EXCEL_ERRORS:
        match = re.search(r"XIRR\(([A-Z]+\d+:[A-Z]+\d+)[,;]", expression)
        return bool(match and _all_zero(item.source_uri, sheet, match.group(1), index))
    if cached_error == "#DIV/0!":
        match = re.search(r"SUM\(([A-Z]+\d+:[A-Z]+\d+)\)/-?([A-Z]+\d+)", expression)
        if not match or not _all_zero(item.source_uri, sheet, match.group(1), index):
            return False
        denominator = index.get((item.source_uri, sheet, match.group(2)))
        return denominator is not None and _numeric_value(denominator) == 0
    return False


def _all_zero(uri, sheet, cell_range, index):
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    values = []
    for row in range(min_row, max_row + 1):
        for column in range(min_col, max_col + 1):
            coordinate = f"{get_column_letter(column)}{row}"
            item = index.get((uri, sheet, coordinate))
            if item is None:
                return False
            values.append(_numeric_value(item))
    return bool(values) and all(value == 0 for value in values)


def _numeric_value(item):
    value = item.cached_value if item.formula else item.value
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None
