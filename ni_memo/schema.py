"""Validated declarative memo product schema."""
from dataclasses import dataclass
import json
from pathlib import Path


REQUIREMENTS = {"required_for_generation", "required_for_ic", "optional"}


@dataclass(frozen=True)
class MemoSection:
    key: str
    title: str
    new_word_section: bool = False
    orientation: str = "portrait"


@dataclass(frozen=True)
class DocumentLayer:
    key: str
    title: str


@dataclass(frozen=True)
class MemoBlock:
    key: str
    title: str
    section: str
    block_type: str
    required: bool
    narrative_key: str | None = None
    field_prefixes: tuple[str, ...] = ()
    counts_toward_completion: bool = True


@dataclass(frozen=True)
class MemoField:
    key: str
    label: str
    section: str
    aliases: tuple[str, ...]
    value_kind: str
    unit_family: str
    requirement: str
    coverage_kind: str = "scalar"
    min_items: int = 1
    project_time_sensitive: bool = False
    counts_toward_fact_coverage: bool = True


@dataclass(frozen=True)
class MemoSchema:
    version: str
    document_layers: tuple[DocumentLayer, ...]
    sections: tuple[MemoSection, ...]
    blocks: tuple[MemoBlock, ...]
    fields: tuple[MemoField, ...]


def load_schema(path=None):
    source = Path(path) if path else Path(__file__).parents[1] / "schema" / "standard_ic_memo.json"
    data = json.loads(source.read_text(encoding="utf-8"))
    document_layers = tuple(DocumentLayer(**item) for item in data["document_layers"])
    sections = tuple(MemoSection(
        key=item["key"], title=item["title"],
        new_word_section=item.get("new_word_section", False),
        orientation=item.get("orientation", "portrait"),
    ) for item in data["sections"])
    blocks = tuple(MemoBlock(
        key=item["key"], title=item["title"], section=item["section"],
        block_type=item["block_type"], required=item["required"],
        narrative_key=item.get("narrative_key"),
        field_prefixes=tuple(item.get("field_prefixes", ())),
        counts_toward_completion=item.get("counts_toward_completion", True),
    ) for item in data["blocks"])
    fields = tuple(MemoField(
        key=item["key"], label=item["label"], section=item["section"],
        aliases=tuple(item["aliases"]), value_kind=item["value_kind"],
        unit_family=item["unit_family"], requirement=item["requirement"],
        coverage_kind=item.get("coverage_kind", "scalar"),
        min_items=item.get("min_items", 1),
        project_time_sensitive=item.get("project_time_sensitive", False),
        counts_toward_fact_coverage=item.get("counts_toward_fact_coverage", True),
    ) for item in data["fields"])
    _validate(data["version"], document_layers, sections, blocks, fields)
    return MemoSchema(data["version"], document_layers, sections, blocks, fields)


def _validate(version, document_layers, sections, blocks, fields):
    if not version or not document_layers or not sections or not blocks or not fields:
        raise ValueError("schema version, document layers, sections, blocks, and fields are required")
    layer_keys = [item.key for item in document_layers]
    section_keys = [item.key for item in sections]
    block_keys = [item.key for item in blocks]
    field_keys = [item.key for item in fields]
    if any(len(keys) != len(set(keys)) for keys in (layer_keys, section_keys, block_keys, field_keys)):
        raise ValueError("schema keys must be unique")
    for block in blocks:
        if block.section not in section_keys:
            raise ValueError(f"unknown section for {block.key}: {block.section}")
        if not block.block_type:
            raise ValueError(f"block type is required for {block.key}")
    for section in sections:
        if section.orientation not in {"portrait", "landscape"}:
            raise ValueError(f"invalid orientation for {section.key}: {section.orientation}")
    for field in fields:
        if field.section not in section_keys:
            raise ValueError(f"unknown section for {field.key}: {field.section}")
        if not field.aliases:
            raise ValueError(f"aliases are required for {field.key}")
        if field.requirement not in REQUIREMENTS:
            raise ValueError(f"invalid requirement for {field.key}: {field.requirement}")
        if field.min_items < 1:
            raise ValueError(f"min_items must be positive for {field.key}")
