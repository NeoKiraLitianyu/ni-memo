"""Source-neutral evidence records. Adapters never assign memo schema keys."""
from dataclasses import asdict, dataclass
from typing import Any


@dataclass(frozen=True)
class EvidenceItem:
    source_uri: str
    source_type: str
    source_location: str
    label: str
    value: Any
    unit: str | None = None
    as_of: str | None = None
    formula: str | None = None
    cached_value: Any = None
    raw_text: str = ""
    context: tuple[str, ...] = ()

    def to_dict(self):
        data = asdict(self)
        data["context"] = list(self.context)
        return data


@dataclass(frozen=True)
class IngestResult:
    items: tuple[EvidenceItem, ...]
    input_errors: tuple[str, ...]
    skipped: tuple[str, ...]
