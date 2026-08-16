"""Reconcile candidate facts without erasing conflicts."""
from collections import defaultdict
from dataclasses import replace
from urllib.parse import urlparse

from ni_memo.facts import FactStatus
from ni_memo.snapshot import FactSnapshot

AUTHORITATIVE = {"government", "company_filing", "audited_report"}


def reconcile(records, policy=None, snapshot_id=None):
    policy = policy or {}
    grouped = defaultdict(list)
    for record in records:
        grouped[record.key].append(record)

    selected = {}
    for key, candidates in grouped.items():
        choice = _policy_choice(candidates, policy.get(key, {}))
        if choice:
            if choice.source_type == "investor_input":
                # truth override / caller selection is authoritative intent: honor it.
                # A divergent model candidate (e.g. a different unit or deal path in
                # the workbook) is recorded as a candidate but must not downgrade the
                # caller's explicit selection to an unresolved conflict.
                selected[key] = replace(choice, status=FactStatus.SOURCE_ONLY)
            else:
                evaluated = _same_time_candidates(candidates)
                values = {_value_key(c.value) for c in evaluated if c.status != FactStatus.MISSING}
                selected[key] = replace(
                    choice,
                    status=FactStatus.CONFLICT if len(values) > 1 else choice.status,
                )
            continue
        evaluated = _same_time_candidates(candidates)
        values = {_value_key(c.value) for c in evaluated if c.status != FactStatus.MISSING}
        if not values:
            selected[key] = replace(evaluated[0], status=FactStatus.MISSING)
        elif all(c.status == FactStatus.RETAINED_ORIGINAL for c in evaluated):
            selected[key] = evaluated[0]
        else:
            authoritative = [c for c in evaluated if c.source_type in AUTHORITATIVE]
            if authoritative and len({_value_key(c.value) for c in authoritative}) == 1:
                # 权威源 (company_filing/government/audited) 优先于模型中间变量
                selected[key] = replace(authoritative[0], status=FactStatus.VERIFIED)
            elif len(values) > 1:
                selected[key] = replace(evaluated[0], status=FactStatus.CONFLICT)
            elif any(c.source_type in AUTHORITATIVE for c in evaluated):
                selected[key] = replace(evaluated[0], status=FactStatus.VERIFIED)
            elif len({_domain(c.source_uri) for c in evaluated if _domain(c.source_uri)}) >= 2:
                selected[key] = replace(evaluated[0], status=FactStatus.CORROBORATED)
            else:
                selected[key] = replace(evaluated[0], status=FactStatus.SOURCE_ONLY)
    return FactSnapshot.create(selected, grouped, snapshot_id=snapshot_id)


def _policy_choice(candidates, rule):
    if "prefer_value" in rule:
        for item in candidates:
            if _value_key(item.value) == _value_key(rule["prefer_value"]):
                return item
    if "prefer_source_type" in rule:
        for item in candidates:
            if item.source_type == rule["prefer_source_type"]:
                return item
    return None


def _same_time_candidates(candidates):
    """Do not compare later public records against a dated XLSX project snapshot."""
    baseline = next((c for c in candidates if c.source_type == "xlsx" and c.as_of), None)
    if baseline and any(c.as_of and c.as_of != baseline.as_of for c in candidates):
        return [c for c in candidates if c.as_of == baseline.as_of]
    return candidates


def _domain(uri):
    return urlparse(uri).netloc.lower()


def _value_key(value):
    return (type(value).__name__, str(value).strip())
