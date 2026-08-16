"""Narrative source material for memo blocks.

A data model (XLSX/CSV/JSON/PDF) provides *facts* (numbers, names, coordinates).
It cannot provide the *investment narrative* — why the business matters, what the
highlight is, how a risk is mitigated. A human (or LLM with editorial control)
supplies that narrative as a JSON sidecar file, organized by memo block key.

The engine renders narrative material only when the caller supplies it. Without
it, the memo honestly shows "资料未提供" placeholders instead of fabricating a
story (T8).
"""
from dataclasses import dataclass, field
from pathlib import Path
import json


@dataclass(frozen=True)
class NarrativeItem:
    title: str
    body: str
    source: str = "narrative"
    sources: tuple[str, ...] = ()

    def to_dict(self):
        return {
            "title": self.title,
            "body": self.body,
            "source": self.source,
            "sources": list(self.sources),
        }


@dataclass(frozen=True)
class NarrativeBundle:
    """Narrative material keyed by memo block key (schema `narrative_key`)."""
    items: dict[str, tuple[NarrativeItem, ...]]

    @classmethod
    def load(cls, path):
        source_path = Path(path).resolve()
        data = json.loads(source_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("narrative file must be a JSON object keyed by block key")
        items = {}
        for block_key, raw_items in data.items():
            if isinstance(raw_items, str):
                raw_items = [{"title": "", "body": raw_items}]
            if not isinstance(raw_items, list):
                raise ValueError(f"narrative[{block_key}] must be a string or list")
            parsed = []
            for item in raw_items:
                if isinstance(item, str):
                    item = {"title": "", "body": item}
                if not isinstance(item, dict) or not str(item.get("body", "")).strip():
                    raise ValueError(f"narrative[{block_key}] item must have a non-empty body")
                parsed.append(NarrativeItem(
                    title=str(item.get("title", "")).strip(),
                    body=str(item.get("body", "")).strip(),
                    source=str(item.get("source", "")).strip() or str(source_path),
                    sources=_sources(item.get("sources", ()), block_key),
                ))
            items[block_key] = tuple(parsed)
        return cls(items)

    def get(self, block_key):
        return self.items.get(block_key, ())

    def to_dict(self):
        return {key: [item.to_dict() for item in items]
                for key, items in sorted(self.items.items())}


def _sources(raw, block_key):
    if isinstance(raw, str):
        raw = (raw,)
    if not isinstance(raw, (list, tuple)):
        raise ValueError(f"narrative[{block_key}] sources must be a string or list")
    return tuple(str(value).strip() for value in raw if str(value).strip())
