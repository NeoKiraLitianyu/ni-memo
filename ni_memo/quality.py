"""Quality and completeness gates over the frozen fact snapshot used for rendering."""
from dataclasses import asdict, dataclass
from datetime import date
from urllib.parse import urlparse

from ni_memo.facts import FactStatus


@dataclass(frozen=True)
class QualityReport:
    snapshot_id: str
    passed: bool
    metrics: dict[str, float]
    pending_items: list[str]
    sidecar_items: list[str]
    invalid_sources: list[str]

    @property
    def grade(self):
        if not self.passed:
            return "FAIL"
        if self.metrics["verified_rate"] == 100.0:
            return "PASS"
        return "PASS_WITH_NOTES"

    def to_dict(self):
        return {
            "snapshot_id": self.snapshot_id,
            "pass": self.passed,
            "metrics": self.metrics,
            "pending_items": self.pending_items,
            "sidecar_items": self.sidecar_items,
            "invalid_sources": self.invalid_sources,
            "grade": self.grade,
        }


def evaluate_snapshot(snapshot, required_keys):
    required = sorted(set(required_keys))
    total = len(required) or 1
    counts = {"verified": 0, "source_only": 0, "retained": 0, "missing": 0}
    pending = []
    invalid = []
    for key in required:
        fact = snapshot.selected.get(key)
        if fact is None or fact.status in {FactStatus.MISSING, FactStatus.CONFLICT}:
            counts["missing"] += 1
            pending.append(key)
            continue
        if fact.status in {FactStatus.VERIFIED, FactStatus.CORROBORATED}:
            counts["verified"] += 1
        elif fact.status == FactStatus.SOURCE_ONLY:
            counts["source_only"] += 1
        elif fact.status == FactStatus.RETAINED_ORIGINAL:
            counts["retained"] += 1
        if fact.source_uri.startswith(("http://", "https://")) and not _valid_public_url(fact.source_uri):
            invalid.append(key)
    metrics = {
        "verified_rate": round(counts["verified"] / total * 100, 1),
        "source_only_rate": round(counts["source_only"] / total * 100, 1),
        "retained_rate": round(counts["retained"] / total * 100, 1),
        "missing_rate": round(counts["missing"] / total * 100, 1),
    }
    pending = sorted(set(pending))
    invalid = sorted(set(invalid))
    return QualityReport(snapshot.snapshot_id, not pending and not invalid, metrics, pending, list(pending), invalid)


@dataclass(frozen=True)
class CompletionItem:
    key: str
    label: str
    requirement: str
    status: str
    evidence_status: str
    selected_keys: tuple[str, ...]
    reasons: tuple[str, ...] = ()
    active_keys: tuple[str, ...] = ()
    later_keys: tuple[str, ...] = ()

    def to_dict(self):
        data = asdict(self)
        for key in ("selected_keys", "reasons", "active_keys", "later_keys"):
            data[key] = list(data[key])
        return data


@dataclass(frozen=True)
class CompletionReport:
    required_completed: int
    required_total: int
    required_rate: float
    all_completed: int
    all_total: int
    all_field_rate: float
    missing_required: tuple[str, ...]
    missing_all: tuple[str, ...]
    items: tuple[CompletionItem, ...]
    slots: tuple[CompletionItem, ...] = ()
    required_slots_completed: int = 0
    required_slots_total: int = 0
    required_slot_rate: float = 100.0
    all_slots_completed: int = 0
    all_slots_total: int = 0
    all_slot_rate: float = 100.0

    def to_dict(self):
        return {
            "required_completed": self.required_completed,
            "required_total": self.required_total,
            "required_rate": self.required_rate,
            "all_completed": self.all_completed,
            "all_total": self.all_total,
            "all_field_rate": self.all_field_rate,
            "missing_required": list(self.missing_required),
            "missing_all": list(self.missing_all),
            "required_slots_completed": self.required_slots_completed,
            "required_slots_total": self.required_slots_total,
            "required_slot_rate": self.required_slot_rate,
            "all_slots_completed": self.all_slots_completed,
            "all_slots_total": self.all_slots_total,
            "all_slot_rate": self.all_slot_rate,
            "state_counts": {
                status: sum(item.status == status for item in self.items)
                for status in ("complete", "partial", "conflict", "later_update", "missing")
            },
            "items": [item.to_dict() for item in self.items],
            "slots": [item.to_dict() for item in self.slots],
        }


def evaluate_completion(schema, snapshot, project_as_of=None, narratives=None, analysis=None):
    """Measure truth coverage by field, semantics, project time and conflicts.

    Narratives are accepted for the slot-completion layer but intentionally do not
    affect fact coverage.  Only ``complete`` contributes to either coverage rate.
    """
    cutoff = _parse_cutoff(project_as_of)
    items = []
    for field in schema.fields:
        candidates = tuple(
            (key, fact) for key, fact in snapshot.selected.items()
            if key == field.key or key.startswith(field.key + ".")
        )
        active, later = _partition_by_project_time(field, candidates, cutoff)
        active_usable = tuple(
            (key, fact) for key, fact in active
            if fact.status not in {FactStatus.MISSING, FactStatus.CONFLICT}
        )
        reasons = []
        if any(fact.status == FactStatus.CONFLICT for _, fact in active):
            status, evidence_status = "conflict", "unresolved"
            reasons.append("unresolved_conflict")
        elif not active and later:
            status = "later_update"
            evidence_status = _field_evidence_status(tuple(fact for _, fact in later))
            reasons.append("after_project_as_of")
        elif not active_usable:
            status, evidence_status = "missing", "none"
            reasons.append("no_usable_project_period_fact")
        else:
            coverage_reasons = _coverage_reasons(field, active_usable, cutoff)
            evidence_status = _field_evidence_status(tuple(fact for _, fact in active_usable))
            if coverage_reasons:
                status = "partial"
                reasons.extend(coverage_reasons)
            else:
                status = "complete"
        active_keys = tuple(key for key, _ in active)
        later_keys = tuple(key for key, _ in later)
        items.append(CompletionItem(
            field.key, field.label, field.requirement, status, evidence_status,
            active_keys + later_keys, tuple(reasons), active_keys, later_keys,
        ))
    countable_keys = {
        field.key for field in schema.fields if field.counts_toward_fact_coverage
    }
    required = tuple(
        item for item in items
        if item.requirement != "optional" and item.key in countable_keys
    )
    all_countable = tuple(item for item in items if item.key in countable_keys)
    required_completed = sum(item.status == "complete" for item in required)
    all_completed = sum(item.status == "complete" for item in all_countable)
    slots = _evaluate_slots(schema, snapshot, tuple(items), narratives, analysis)
    required_slots = tuple(item for item in slots if item.requirement == "required")
    required_slots_completed = sum(item.status == "complete" for item in required_slots)
    all_slots_completed = sum(item.status == "complete" for item in slots)
    return CompletionReport(
        required_completed=required_completed,
        required_total=len(required),
        required_rate=round(required_completed / len(required) * 100, 1) if required else 100.0,
        all_completed=all_completed,
        all_total=len(all_countable),
        all_field_rate=(
            round(all_completed / len(all_countable) * 100, 1) if all_countable else 100.0
        ),
        missing_required=tuple(item.key for item in required if item.status != "complete"),
        missing_all=tuple(item.key for item in all_countable if item.status != "complete"),
        items=tuple(items),
        slots=slots,
        required_slots_completed=required_slots_completed,
        required_slots_total=len(required_slots),
        required_slot_rate=(
            round(required_slots_completed / len(required_slots) * 100, 1)
            if required_slots else 100.0
        ),
        all_slots_completed=all_slots_completed,
        all_slots_total=len(slots),
        all_slot_rate=(round(all_slots_completed / len(slots) * 100, 1) if slots else 100.0),
    )


def _evaluate_slots(schema, snapshot, field_items, narratives, analysis):
    slots = []
    for block in schema.blocks:
        if not block.counts_toward_completion:
            continue
        narrative_key = block.narrative_key or block.key
        narrative_items = narratives.get(narrative_key) if narratives is not None else ()
        analysis_items = analysis.get(block.key) if analysis is not None else ()
        child_items = tuple(
            item for item in field_items
            if any(_field_matches_prefix(item.key, prefix) for prefix in block.field_prefixes)
        )
        if narrative_items or analysis_items:
            status = "complete"
            evidence_status = "editorial_sourced"
            reasons = ("narrative_slot",) if narrative_items else ("analysis_slot",)
        elif block.block_type == "source_notes" and snapshot.selected:
            status, evidence_status, reasons = "complete", "source_index", ("source_index",)
        else:
            status, evidence_status, reasons = _slot_state(block, child_items)
        active_keys = tuple(key for item in child_items for key in item.active_keys)
        later_keys = tuple(key for item in child_items for key in item.later_keys)
        slots.append(CompletionItem(
            block.key, block.title, "required" if block.required else "optional",
            status, evidence_status, active_keys + later_keys, tuple(reasons),
            active_keys, later_keys,
        ))
    return tuple(slots)


def _field_matches_prefix(key, prefix):
    if prefix.endswith("."):
        return key.startswith(prefix)
    return key == prefix or key.startswith(prefix + ".")


def _slot_state(block, child_items):
    if not child_items:
        return "missing", "none", ("no_slot_evidence",)
    required_children = tuple(
        item for item in child_items if item.requirement != "optional"
    ) or child_items
    states = {item.status for item in required_children}
    if "conflict" in states:
        return "conflict", "unresolved", ("child_field_conflict",)
    if states == {"complete"}:
        return "complete", "field_states", ("required_child_fields_complete",)
    if states <= {"missing", "later_update"} and "later_update" in states:
        return "later_update", "field_states", ("only_later_or_missing_child_fields",)
    if states != {"missing"}:
        return "partial", "field_states", ("required_child_fields_incomplete",)
    return "missing", "none", ("required_child_fields_missing",)


def _parse_cutoff(value):
    if value in (None, ""):
        return None
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value))
    except ValueError as exc:
        raise ValueError("project_as_of must be YYYY-MM-DD") from exc


def _fact_date(value):
    if not value:
        return None
    text = str(value).strip()
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        return None


def _partition_by_project_time(field, candidates, cutoff):
    if cutoff is None or not field.project_time_sensitive:
        return candidates, ()
    active, later = [], []
    for key, fact in candidates:
        effective = _fact_date(fact.as_of)
        (later if effective and effective > cutoff else active).append((key, fact))
    return tuple(active), tuple(later)


def _coverage_reasons(field, usable, cutoff):
    reasons = []
    if len(usable) < field.min_items:
        reasons.append("minimum_items_not_met")
    if field.coverage_kind == "valuation_table" and not any(
        _has_quantitative_valuation(fact.value) for _, fact in usable
    ):
        reasons.append("quantitative_valuation_required")
    if cutoff is not None and field.project_time_sensitive and any(
        _fact_date(fact.as_of) is None for _, fact in usable
    ):
        reasons.append("date_unbound")
    return tuple(reasons)


def _has_quantitative_valuation(value):
    if not isinstance(value, (dict, list, tuple)):
        return False
    values = value.values() if isinstance(value, dict) else value
    return any(
        isinstance(item, (int, float)) and not isinstance(item, bool)
        or _has_quantitative_valuation(item)
        for item in values
    )


def _field_evidence_status(facts):
    statuses = {fact.status for fact in facts}
    if statuses <= {FactStatus.VERIFIED, FactStatus.CORROBORATED}:
        return "verified"
    if FactStatus.SOURCE_ONLY in statuses:
        return "source_only"
    if FactStatus.RETAINED_ORIGINAL in statuses:
        return "retained_original"
    return "mixed"


def _valid_public_url(url):
    parsed = urlparse(url)
    lowered = url.lower()
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc) and not any(
        marker in lowered for marker in ("_xxx", "example.com", "placeholder")
    )
