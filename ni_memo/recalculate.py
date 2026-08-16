"""Non-destructive LibreOffice recalculation of copied workbook outputs."""
from dataclasses import asdict, dataclass, replace
from pathlib import Path
import shutil
import subprocess

from ni_memo.formula_audit import audit_formula_evidence, count_formula_nodes
from ni_memo.ingest import ingest_inputs


@dataclass(frozen=True)
class RecalculationReport:
    status: str
    engine: str | None
    outputs: tuple[str, ...]
    sources: tuple[str, ...]
    source_formula_count: int
    recalculated_formula_count: int
    audit: dict | None
    messages: tuple[str, ...]

    def to_dict(self):
        data = asdict(self)
        data["outputs"] = list(self.outputs)
        data["sources"] = list(self.sources)
        data["messages"] = list(self.messages)
        return data


def recalculate_workbooks(inputs, out_dir, soffice_path=None):
    workbooks = _xlsx_inputs(inputs)
    source_count = count_formula_nodes(workbooks)
    sources = tuple(str(path) for path in workbooks)
    if not workbooks:
        return RecalculationReport("not_applicable", None, (), (), 0, 0, None, ())
    engine = _find_soffice(soffice_path)
    if engine is None:
        return RecalculationReport(
            "not_run", None, (), sources, source_count, 0, None,
            ("LibreOffice soffice was not found",),
        )

    root = Path(out_dir)
    root.mkdir(parents=True, exist_ok=True)
    outputs, messages = [], []
    for index, source in enumerate(workbooks, start=1):
        target_dir = root / f"workbook-{index:02d}"
        profile = root / f"profile-{index:02d}"
        target_dir.mkdir(parents=True, exist_ok=True)
        profile.mkdir(parents=True, exist_ok=True)
        target = target_dir / source.name
        # 沙箱环境 (Windows 回收站不可用) 下 unlink/replace 删除旧文件会被 safe-delete 拦截。
        # 重算副本是可再生中间产物: 旧文件直接 rename 避让 (保留为 .prev.xlsx), 不删除任何文件。
        # 这样 LibreOffice --convert-to 会写到干净的目标路径, 且全程零删除。
        stale = target.with_suffix(".prev.xlsx")
        if target.exists() and not stale.exists():
            try:
                target.replace(stale)
            except OSError:
                pass
        command = [
            str(engine), "--headless", f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to", "xlsx", "--outdir", str(target_dir), str(source),
        ]
        completed = subprocess.run(
            command, capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=180, check=False,
        )
        raw_out = (completed.stdout or "") + "\n" + (completed.stderr or "")
        messages.extend(line for line in raw_out.splitlines() if line.strip())
        if completed.returncode != 0 or not target.exists():
            return RecalculationReport(
                "fail", str(engine), tuple(str(path) for path in outputs), sources,
                source_count, 0, None,
                tuple(messages or [f"recalculation failed for {source}"]),
            )
        outputs.append(target)

    recalculated_count = count_formula_nodes(outputs)
    evidence = ingest_inputs(outputs)
    source_identities = {
        str(output.resolve()).casefold(): str(source)
        for source, output in zip(workbooks, outputs)
    }
    audit = audit_formula_evidence(
        evidence.items, expected_formula_count=recalculated_count, recalculation_status="pass",
        source_identities=source_identities,
    )
    status = assess_recalculation(source_count, recalculated_count, audit.status)
    return RecalculationReport(
        status, str(engine), tuple(str(path) for path in outputs), sources, source_count,
        recalculated_count, audit.to_dict(), tuple(messages),
    )


def assess_recalculation(source_formula_count, recalculated_formula_count, audit_status):
    if source_formula_count != recalculated_formula_count or audit_status != "pass":
        return "warning"
    return "pass"


def _xlsx_inputs(inputs):
    result = []
    for raw in inputs:
        path = Path(raw).resolve()
        if path.is_dir():
            result.extend(sorted(item for item in path.rglob("*.xlsx") if item.is_file()))
        elif path.suffix.lower() == ".xlsx" and path.is_file():
            result.append(path)
    return sorted(set(result), key=lambda path: str(path).lower())


def _find_soffice(explicit):
    candidates = [
        Path(explicit) if explicit else None,
        Path(shutil.which("soffice.com")) if shutil.which("soffice.com") else None,
        Path(shutil.which("soffice")) if shutil.which("soffice") else None,
        Path(r"C:\Program Files\LibreOffice\program\soffice.com"),
    ]
    return next((path for path in candidates if path and path.is_file()), None)


def merge_recalculated_values(evidence_items, report):
    """Overlay LibreOffice-recalculated cached values onto formula evidence items.

    This is the key step that makes formula output part of the data chain rather
    than an acceptance-only gate: ingest initially reads the workbook's *cached*
    formula results (which may be stale if the model was saved without a full
    recalc). After LibreOffice re-saves the copy with fresh cached values, we
    re-ingest the copy and overlay those fresh values onto the original evidence
    by (resolved source path, sheet, coordinate). discovery then sees the
    recalculated value as the authoritative formula result.
    """
    if not report.outputs:
        return evidence_items, 0
    recalculated = ingest_inputs(report.outputs)
    source_by_output = {
        str(Path(output).resolve()).casefold(): str(Path(source).resolve()).casefold()
        for source, output in zip(report.sources, report.outputs)
    }
    fresh = {}
    for item in recalculated.items:
        if item.formula and item.cached_value is not None:
            output_key = str(Path(item.source_uri).resolve()).casefold()
            key = (source_by_output.get(output_key, output_key), item.source_location)
            fresh[key] = item.cached_value
    merged = []
    overlaid = 0
    for item in evidence_items:
        if item.formula and item.cached_value is not None:
            source_key = str(Path(item.source_uri).resolve()).casefold()
            new_value = fresh.get((source_key, item.source_location))
            if new_value is not None and str(new_value).upper() not in {
                "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!"
            }:
                item = replace(item, cached_value=new_value)
                overlaid += 1
        merged.append(item)
    return tuple(merged), overlaid
