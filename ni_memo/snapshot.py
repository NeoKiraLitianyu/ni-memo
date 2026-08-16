"""Frozen fact snapshot shared by every downstream artifact."""
from dataclasses import dataclass
import json
from pathlib import Path
from uuid import uuid4

from ni_memo.facts import FactRecord


@dataclass(frozen=True)
class FactSnapshot:
    snapshot_id: str
    selected: dict[str, FactRecord]
    candidates: dict[str, tuple[FactRecord, ...]]

    @classmethod
    def create(cls, selected, candidates, snapshot_id=None):
        return cls(snapshot_id or str(uuid4()), dict(selected), {
            key: tuple(value) for key, value in candidates.items()
        })

    def to_dict(self):
        return {
            "snapshot_id": self.snapshot_id,
            "selected": {k: self.selected[k].to_dict() for k in sorted(self.selected)},
            "candidates": {
                k: [f.to_dict() for f in self.candidates[k]] for k in sorted(self.candidates)
            },
        }

    def write(self, path):
        Path(path).write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")

    @classmethod
    def read(cls, path):
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.create(
            {k: FactRecord.from_dict(v) for k, v in data["selected"].items()},
            {k: [FactRecord.from_dict(v) for v in values] for k, values in data["candidates"].items()},
            snapshot_id=data["snapshot_id"],
        )

