"""Canonical, source-traceable facts used by every ni-memo stage."""
from dataclasses import asdict, dataclass
from typing import Any


class FactStatus:
    VERIFIED = "verified"
    CORROBORATED = "corroborated"
    SOURCE_ONLY = "source_only"
    RETAINED_ORIGINAL = "retained_original"
    CONFLICT = "conflict"
    MISSING = "missing"

    ALL = {
        VERIFIED,
        CORROBORATED,
        SOURCE_ONLY,
        RETAINED_ORIGINAL,
        CONFLICT,
        MISSING,
    }


@dataclass(frozen=True)
class FactRecord:
    key: str
    value: Any
    unit: str | None
    as_of: str | None
    source_type: str
    source_uri: str
    source_location: str
    extraction_method: str
    confidence: str
    status: str
    evidence: tuple[str, ...] = ()

    def __post_init__(self):
        if not self.key.strip():
            raise ValueError("key is required")
        if self.status not in FactStatus.ALL:
            raise ValueError(f"unknown status: {self.status}")
        if self.status in {FactStatus.VERIFIED, FactStatus.CORROBORATED, FactStatus.SOURCE_ONLY}:
            for name in ("source_type", "source_uri", "source_location", "extraction_method"):
                if not str(getattr(self, name)).strip():
                    raise ValueError(f"{name} is required for sourced facts")
        object.__setattr__(self, "evidence", tuple(str(item) for item in self.evidence))

    @property
    def is_verified(self) -> bool:
        return self.status in {FactStatus.VERIFIED, FactStatus.CORROBORATED}

    @classmethod
    def from_dict(cls, data: dict) -> "FactRecord":
        return cls(**data)

    def to_dict(self) -> dict:
        data = asdict(self)
        data["evidence"] = list(self.evidence)
        return data
