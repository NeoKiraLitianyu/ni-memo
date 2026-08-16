"""Load agent-collected public evidence into canonical facts."""
import json
from pathlib import Path
from urllib.parse import urlparse

from ni_memo.facts import FactRecord, FactStatus


def load_web_evidence(path):
    items = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(items, list):
        raise ValueError("web evidence must be a list")
    facts = []
    for item in items:
        url = str(item.get("url", ""))
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc or any(
            marker in url.lower() for marker in ("_xxx", "placeholder", ".test/")
        ):
            raise ValueError(f"invalid public URL: {url}")
        retrieved = str(item.get("retrieved_at", ""))
        if not retrieved:
            raise ValueError("retrieved_at is required")
        evidence = item.get("evidence", "")
        facts.append(FactRecord(
            key=item["key"], value=item.get("value"), unit=item.get("unit"), as_of=item.get("as_of"),
            source_type=item.get("source_type", "other"), source_uri=url,
            source_location=item.get("source_location", "page"), extraction_method="public_evidence",
            confidence=item.get("confidence", "medium"), status=FactStatus.SOURCE_ONLY,
            evidence=tuple(filter(None, (str(evidence), f"retrieved_at={retrieved}"))),
        ))
    return facts
