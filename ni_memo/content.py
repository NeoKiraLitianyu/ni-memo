"""Compile schema and frozen facts into typed memo content blocks."""
from dataclasses import dataclass
import re


BLOCK_TYPES = {
    "narrative": "narrative",
    "key_facts_table": "key_facts",
    "terms_table": "terms_table",
    "team_table": "evidence_table",
    "operations_table": "evidence_table",
    "financial_table": "financial_table",
    "returns_table": "evidence_table",
    "sensitivity_table": "evidence_table",
    "comparables_table": "evidence_table",
    "bullet_list": "bullet_list",
    "product_profile": "bullet_list",
    "risk_cards": "risk",
    "chart": "chart",
    "source_notes": "sources",
    "completion_review": "completion_review",
}

@dataclass(frozen=True)
class MemoBlockContent:
    key: str
    title: str
    block_type: str
    fact_keys: tuple[str, ...]
    workbook_regions: tuple[object, ...]


@dataclass(frozen=True)
class MemoSectionContent:
    key: str
    title: str
    blocks: tuple[MemoBlockContent, ...]


@dataclass(frozen=True)
class MemoContent:
    sections: tuple[MemoSectionContent, ...]
    referenced_fact_keys: tuple[str, ...]
    referenced_workbook_regions: tuple[object, ...]


def compose_standard_content(schema, snapshot, workbook_profiles=()):
    known_keys = tuple(field.key for field in schema.fields)
    referenced = set()
    sections = []
    for section in schema.sections:
        blocks = []
        for block in (item for item in schema.blocks if item.section == section.key):
            if block.block_type not in BLOCK_TYPES:
                raise ValueError(f"unsupported memo block type: {block.block_type}")
            if block.block_type == "source_notes":
                # 附录来源索引: 只列正文已引用的 facts (来源文件/单元格), 不重复渲染 region 表格
                # (region 原始表格已在所属章节正文渲染, 附录再挂 region 会导致同一表格出现两遍)
                fact_keys = tuple(sorted(referenced))
                block_regions = ()
            else:
                fact_keys = _fact_keys(block, snapshot, known_keys)
                referenced.update(fact_keys)
                block_regions = ()
            has_content = bool(
                fact_keys or block_regions
                or block.block_type in {"source_notes", "completion_review"}
            )
            block_type = BLOCK_TYPES[block.block_type] if has_content else "gap"
            blocks.append(MemoBlockContent(block.key, block.title, block_type, fact_keys, block_regions))
        sections.append(MemoSectionContent(section.key, section.title, tuple(blocks)))
    return MemoContent(tuple(sections), tuple(sorted(referenced)), ())


def _fact_keys(block, snapshot, known_keys):
    prefixes = block.field_prefixes
    keys = tuple(
        key for key in sorted(snapshot.selected)
        if any(key == prefix or key.startswith(prefix) for prefix in prefixes)
        and any(key == field or key.startswith(field + ".") for field in known_keys)
    )
    if block.key == "financial.historical":
        return tuple(key for key in keys if _period(key).endswith("A"))
    if block.key == "financial.forecast":
        return tuple(key for key in keys if _period(key).endswith("E"))
    return keys


def _period(key):
    match = re.search(r"\.((?:19|20)\d{2}(?:[aAeE]|_\d{1,2}[mM]))$", key)
    return match.group(1).upper() if match else ""
