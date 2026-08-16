"""Load strict project mapping files."""
import json
from pathlib import Path


def load_mapping(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    rules = data.get("facts") if isinstance(data, dict) else data
    if not isinstance(rules, list):
        raise ValueError("mapping must contain a facts list")
    keys = [r.get("key") for r in rules]
    if any(not key for key in keys) or len(keys) != len(set(keys)):
        raise ValueError("mapping keys must be present and unique")
    return rules
