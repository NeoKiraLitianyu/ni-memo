"""Command-line entry point for a DOCX delivery plus internal audit evidence."""
import argparse
import json
import os
import re
from pathlib import Path

from ni_memo.analysis import compose_model_analysis
from ni_memo.adapters.xlsx import extract_mapped_facts
from ni_memo.adapters.web import load_web_evidence
from ni_memo.discover import discover
from ni_memo.document_quality import (
    compare_document_fidelity,
    compare_formal_template,
    evaluate_document_composition,
)
from ni_memo.evidence import IngestResult
from ni_memo.facts import FactRecord, FactStatus
from ni_memo.fields import FieldUpdateReport, update_docx_fields
from ni_memo.formula_audit import audit_formula_evidence, count_formula_nodes, formula_issue_key
from ni_memo.ingest import ingest_inputs
from ni_memo.mapping import load_mapping
from ni_memo.narratives import NarrativeBundle
from ni_memo.quality import evaluate_completion, evaluate_snapshot
from ni_memo.recalculate import merge_recalculated_values, recalculate_workbooks
from ni_memo.reconcile import reconcile
from ni_memo.render import (render_default_docx, render_docx, render_reference_docx,
                            render_standard_memo)
from ni_memo.schema import load_schema
from ni_memo.visual import render_and_inspect
from ni_memo.workbook_profile import profile_workbook


def build_bundle(workbook, mapping_path, out_dir, template=None, snapshot_id=None, evidence=None,
                 work_dir=None):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    rules = load_mapping(mapping_path)
    facts = extract_mapped_facts(workbook, rules)
    if evidence:
        facts.extend(load_web_evidence(evidence))
    snapshot = reconcile(facts, snapshot_id=snapshot_id)
    work = _resolve_work_dir(work_dir, snapshot.snapshot_id, out)
    required = sorted({rule["key"] for rule in rules} | {fact.key for fact in facts})
    quality = evaluate_snapshot(snapshot, required)

    snapshot.write(work / "facts.json")
    _write_json(work / "acceptance_report.json", quality.to_dict())
    pending = "\n".join(f"- {key}" for key in quality.pending_items) or "无。"
    (work / "pending.md").write_text(
        f"# 待验证与冲突项\n\nsnapshot_id: {snapshot.snapshot_id}\n\n{pending}\n", encoding="utf-8"
    )
    if template:
        render_report = render_docx(template, out / "memo.docx", snapshot)
    else:
        render_report = render_default_docx(out / "memo.docx", snapshot, Path(workbook).stem, quality.grade)
    summary = {
        "snapshot_id": snapshot.snapshot_id,
        "grade": quality.grade,
        "fact_count": len(snapshot.selected),
        "pending_count": len(quality.pending_items),
        "memo": render_report.output_path if render_report else None,
        "work_dir": str(work),
    }
    _write_json(work / "run_summary.json", summary)
    return summary


def build_from_inputs(inputs, out_dir, schema_path=None, snapshot_id=None, run_visual=True,
                      public_evidence=None, work_dir=None, narrative_path=None,
                      reference_docx=None, formal_template=None, truth_overrides=None, policy=None,
                      project_as_of=None):
    """Build the standard product without requiring a Database workbook or template.

    truth_overrides: optional caller-supplied selections with key, value, source_uri,
        source_location, and evidence. They remain investor_input/SOURCE_ONLY; selection
        does not become independent verification. supplement_missing additionally requires
        a schema field key and a source_uri present in this run.
    policy: optional {key: {prefer_value|prefer_source_type}} passed to reconcile.
    """
    inputs = tuple(inputs)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    schema = load_schema(schema_path)
    resolved_id = snapshot_id or "default"
    work = _resolve_work_dir(work_dir, resolved_id, out)

    # 1) 摄入原始证据 (含公式缓存值)
    ingested = ingest_inputs(inputs)
    # 模型语义区域: 供 content 编译把原始工作簿表格区域渲染进 memo (海古德版式需要)
    workbook_profiles = () if reference_docx else _profile_workbooks(inputs)

    # 2) 公式重算: LibreOffice 重算副本 → 重算缓存值覆盖原始缓存值
    #    (公式是数据链的一环: 采集 → 重算 → 结果进 discover, 而非仅验收门禁)
    recalculation = recalculate_workbooks(inputs, work / "recalculated")
    if recalculation.status not in {"not_applicable", "not_run"}:
        evidence_items, overlaid = merge_recalculated_values(ingested.items, recalculation)
        ingested = IngestResult(tuple(evidence_items), ingested.input_errors, ingested.skipped)
    else:
        overlaid = 0

    # 3) 用合并后的证据做确定性 discover
    discovery = discover(schema, ingested.items)

    records = []
    for binding in discovery.bindings.values():
        records.append(_binding_fact(binding))
    for candidates in discovery.mapping_conflicts.values():
        records.extend(_binding_fact(binding) for binding in candidates)
    public_records = load_web_evidence(public_evidence) if public_evidence else []
    records.extend(public_records)
    discovered_keys = set(discovery.bindings) | set(discovery.mapping_conflicts)
    run_sources = {item.source_uri for item in ingested.items}
    if narrative_path:
        run_sources.add(str(Path(narrative_path).resolve()))
    truth_records = _truth_override_records(
        truth_overrides,
        discovered_keys,
        {field.key for field in schema.fields},
        run_sources,
    )
    records.extend(truth_records)
    reconciliation_policy = dict(policy or {})
    for record in truth_records:
        reconciliation_policy.setdefault(record.key, {"prefer_source_type": "investor_input"})
    snapshot = reconcile(records, policy=reconciliation_policy, snapshot_id=resolved_id)
    required = _required_snapshot_keys(schema, snapshot)
    quality = evaluate_snapshot(snapshot, required)
    narratives = NarrativeBundle.load(narrative_path) if narrative_path else None
    analysis = None if reference_docx else compose_model_analysis(schema, snapshot, workbook_profiles)
    completion = evaluate_completion(
        schema, snapshot, project_as_of=project_as_of,
        narratives=narratives, analysis=analysis,
    )
    confirmed_intrinsic = set((recalculation.audit or {}).get("issue_keys", ()))
    selected_fact_sources = {
        formula_issue_key(fact.source_uri, fact.source_location)
        for fact in snapshot.selected.values()
    }
    formula = audit_formula_evidence(
        ingested.items, expected_formula_count=count_formula_nodes(inputs),
        recalculation_status=recalculation.status,
        confirmed_intrinsic_issues=confirmed_intrinsic,
        selected_fact_sources=selected_fact_sources,
    )

    grade = quality.grade
    if reference_docx:
        # golden-reference (byte-preservation) mode: fidelity is the primary gate.
        # A historical memo's narrative/graphics assets are not derivable from the
        # workbook, so fact-coverage gaps (missing/conflict) are expected and must
        # not force FAIL. Grade derives from verification quality instead:
        # PASS if every present required fact is verified/corroborated, else
        # PASS_WITH_NOTES. Formula gates still bind absolutely below.
        verified_rate = quality.metrics.get("verified_rate", 0.0)
        grade = "PASS" if verified_rate == 100.0 else "PASS_WITH_NOTES"
    elif completion.required_rate < 100.0 or completion.required_slot_rate < 100.0:
        grade = "FAIL"
    if formula.status == "fail":
        grade = "FAIL"
    elif formula.status == "warning" and grade == "PASS":
        grade = "PASS_WITH_NOTES"

    snapshot.write(work / "facts.json")
    mapping = discovery.to_dict()
    mapping["schema_version"] = schema.version
    mapping["snapshot_id"] = snapshot.snapshot_id
    _write_json(work / "mapping.json", mapping)
    _write_json(work / "input_report.json", {
        "snapshot_id": snapshot.snapshot_id,
        "project_as_of": project_as_of,
        "evidence_items": len(ingested.items),
        "formula_recalculated_overlays": overlaid,
        "recalculation_status": recalculation.status,
        "input_errors": list(ingested.input_errors),
        "skipped": list(ingested.skipped),
        "sources": sorted({item.source_uri for item in ingested.items}),
        "public_sources": sorted({item.source_uri for item in public_records}),
        "workbook_regions": [
            f"{region.sheet}!{region.cell_range}"
            for profile in workbook_profiles for region in profile.regions
        ],
    })
    pending_lines = [
        f"- {item.key} [{item.status}]"
        + (f"：{', '.join(item.reasons)}" if item.reasons else "")
        for item in completion.items if item.status != "complete"
    ]
    pending_lines.extend(
        f"- 槽位 {item.key} [{item.status}]"
        + (f"：{', '.join(item.reasons)}" if item.reasons else "")
        for item in completion.slots if item.status != "complete"
    )
    if formula.status in {"warning", "fail"}:
        pending_lines.append(f"- 公式门禁：{formula.status}（recalculation={formula.recalculation_status}）")
    (work / "pending.md").write_text(
        f"# 待验证、冲突与门禁项\n\nsnapshot_id: {snapshot.snapshot_id}\n\n"
        + ("\n".join(pending_lines) if pending_lines else "无。") + "\n",
        encoding="utf-8",
    )
    project_name = _project_name(snapshot, inputs)
    if reference_docx:
        # 正式参考文档保真模式: 无损保留历史 memo (不重写其 XML)
        rendered = render_reference_docx(reference_docx, out / "memo.docx", snapshot.snapshot_id)
        fidelity = compare_document_fidelity(out / "memo.docx", reference_docx)
        _write_json(work / "document_fidelity.json", fidelity.to_dict())
        field_update = FieldUpdateReport("not_applicable", None, ())
        composition = None
        formal_report = None
    else:
        rendered = render_standard_memo(
            out / "memo.docx", schema, snapshot, quality, formula, project_name,
            narratives=narratives, workbook_profiles=workbook_profiles, analysis=analysis,
            completion=completion, project_as_of=project_as_of,
        )
        fidelity = None
        field_update = (
            update_docx_fields(out / "memo.docx", work_dir=work / "field-update")
            if run_visual else FieldUpdateReport("not_run", None, ("visual delivery disabled",))
        )
        composition = evaluate_document_composition(out / "memo.docx")
        formal_report = compare_formal_template(
            out / "memo.docx", formal_template, schema,
        ) if formal_template else None
        if formal_report:
            _write_json(work / "formal_template_report.json", formal_report.to_dict())
    visual = (render_and_inspect(
        out / "memo.docx", work / "visual", tuple(section.title for section in schema.sections),
    ) if run_visual else None)
    if visual and visual.status == "fail":
        grade = "FAIL"
    if field_update.status == "fail":
        grade = "FAIL"
    elif field_update.status == "not_run" and grade == "PASS":
        grade = "PASS_WITH_NOTES"
    if fidelity and fidelity.status == "fail":
        grade = "FAIL"
    if formal_report and formal_report.status == "fail":
        grade = "FAIL"
    acceptance = quality.to_dict()
    acceptance["project_as_of"] = project_as_of
    acceptance["formula"] = formula.to_dict()
    acceptance["recalculation"] = recalculation.to_dict()
    acceptance["recalculation"]["overlaid_evidence"] = overlaid
    acceptance["narrative"] = {
        "supplied": narrative_path is not None,
        "blocks": sorted(narratives.items) if narratives else [],
        "total_items": sum(len(items) for items in narratives.items.values()) if narratives else 0,
    }
    acceptance["analysis"] = analysis.metrics() if analysis else {
        "generated": False,
        "block_count": 0,
        "claim_count": 0,
        "body_characters": 0,
        "long_paragraphs": 0,
        "all_claims_sourced": True,
    }
    acceptance["grade"] = grade
    acceptance["pass"] = grade != "FAIL"
    acceptance["visual"] = visual.to_dict() if visual else {"status": "not_run"}
    acceptance["field_update"] = field_update.to_dict()
    acceptance["completion"] = completion.to_dict()
    acceptance["document_fidelity"] = fidelity.to_dict() if fidelity else {"status": "not_applicable"}
    acceptance["formal_template"] = (
        formal_report.to_dict() if formal_report else {"status": "not_applicable"}
    )
    acceptance["document_composition"] = (
        composition.to_dict() if composition else {"status": "not_applicable"}
    )
    _write_json(work / "acceptance_report.json", acceptance)
    summary = {
        "snapshot_id": snapshot.snapshot_id,
        "schema_version": schema.version,
        "project_as_of": project_as_of,
        "grade": grade,
        "fact_count": len(snapshot.selected),
        "pending_count": len(quality.pending_items),
        "formula_status": formula.status,
        "visual_status": visual.status if visual else "not_run",
        "field_update_status": field_update.status,
        "fidelity_status": fidelity.status if fidelity else "not_applicable",
        "composition_status": composition.status if composition else "not_applicable",
        "formal_template_status": formal_report.status if formal_report else "not_applicable",
        "workbook_region_count": sum(len(profile.regions) for profile in workbook_profiles),
        "memo": rendered.output_path,
        "work_dir": str(work),
    }
    _write_json(work / "run_summary.json", summary)
    return summary


def _truth_override_records(overrides, discovered_keys, schema_keys=(), run_sources=()):
    """Bind caller-selected values without promoting them to verified evidence."""
    if not overrides:
        return []
    schema_keys = set(schema_keys)
    allowed_sources = {_source_key(uri) for uri in run_sources}
    records = []
    for item in overrides:
        key = item["key"]
        discovered = key in discovered_keys or any(
            key == dk or dk.startswith(key + ".") for dk in discovered_keys
        )
        supplement = bool(item.get("supplement_missing"))
        if not discovered and not supplement:
            continue
        if supplement and key not in schema_keys:
            raise ValueError(f"truth override key is not in schema: {key}")
        for name in ("source_uri", "source_location"):
            if not str(item.get(name, "")).strip():
                raise ValueError(f"{name} is required for truth override: {key}")
        raw_evidence = item.get("evidence") or ()
        if isinstance(raw_evidence, str):
            raw_evidence = (raw_evidence,)
        evidence = tuple(str(value).strip() for value in raw_evidence if str(value).strip())
        if not evidence:
            raise ValueError(f"evidence is required for truth override: {key}")
        if _source_key(item["source_uri"]) not in allowed_sources:
            raise ValueError(f"truth override source_uri is not a run input: {item['source_uri']}")
        records.append(FactRecord(
            key=key,
            value=item["value"],
            unit=item.get("unit"),
            as_of=item.get("as_of"),
            source_type="investor_input",
            source_uri=str(item["source_uri"]),
            source_location=str(item["source_location"]),
            extraction_method="truth_override",
            confidence=item.get("confidence", "high"),
            status=FactStatus.SOURCE_ONLY,
            evidence=evidence,
        ))
    return records


def _source_key(uri):
    value = str(uri).strip()
    if re.match(r"^https?://", value, flags=re.IGNORECASE):
        return value.casefold()
    return str(Path(value).resolve()).casefold()


def _binding_fact(binding):
    evidence = binding.evidence
    evidence_lines = [f"raw={evidence.raw_text}"]
    if evidence.formula:
        evidence_lines.extend((f"formula={evidence.formula}", f"cached={evidence.cached_value}"))
    return FactRecord(
        key=binding.key,
        value=evidence.value,
        unit=evidence.unit,
        as_of=evidence.as_of,
        source_type=evidence.source_type,
        source_uri=evidence.source_uri,
        source_location=evidence.source_location,
        extraction_method=binding.rule,
        confidence=binding.confidence,
        status=FactStatus.SOURCE_ONLY,
        evidence=tuple(evidence_lines),
    )


def _required_snapshot_keys(schema, snapshot):
    required = []
    selected_keys = set(snapshot.selected)
    for field in schema.fields:
        if field.requirement == "optional":
            continue
        series = sorted(key for key in selected_keys if key.startswith(field.key + "."))
        qualified_years = {
            match.group(1) for key in series
            if (match := re.search(r"\.(20\d{2})[ae]$", key))
        }
        series = [key for key in series
                  if not ((match := re.search(r"\.(20\d{2})$", key)) and match.group(1) in qualified_years)]
        required.extend(series or [field.key])
    return sorted(set(required))


def _project_name(snapshot, inputs):
    company = snapshot.selected.get("company.name")
    if company and company.status != FactStatus.CONFLICT:
        return str(company.value)
    first = Path(next(iter(inputs)))
    return first.stem if first.is_file() else first.name


def _profile_workbooks(inputs):
    paths = set()
    for raw in inputs:
        path = Path(raw).resolve()
        if path.is_dir():
            paths.update(candidate for candidate in path.rglob("*.xlsx") if candidate.is_file())
        elif path.suffix.lower() == ".xlsx" and path.is_file():
            paths.add(path)
    return tuple(profile_workbook(path) for path in sorted(paths, key=lambda item: str(item).casefold()))


def _write_json(path, data):
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _resolve_work_dir(work_dir, snapshot_id, out_dir):
    if work_dir:
        work = Path(work_dir)
    else:
        # Portable default: prefer NI_MEMO_WORK_ROOT, else a sibling of the
        # delivery directory. No machine-specific absolute paths.
        root = os.environ.get("NI_MEMO_WORK_ROOT")
        if root:
            work = Path(root) / snapshot_id
        else:
            work = Path(out_dir).resolve().parent / f".ni-memo-work-{snapshot_id}"
    resolved_work = work.resolve()
    resolved_out = Path(out_dir).resolve()
    if resolved_work == resolved_out or resolved_out in resolved_work.parents:
        raise ValueError("work directory must be separate from and outside the delivery directory")
    work.mkdir(parents=True, exist_ok=True)
    return work


def main(argv=None):
    parser = argparse.ArgumentParser(description="Build a source-traceable DOCX investment memo")
    parser.add_argument("--inputs", nargs="+")
    parser.add_argument("--workbook")
    parser.add_argument("--mapping")
    parser.add_argument("--out", required=True)
    parser.add_argument("--work-dir")
    parser.add_argument("--template")
    parser.add_argument("--evidence")
    parser.add_argument("--narrative")
    parser.add_argument("--schema")
    parser.add_argument("--snapshot-id")
    parser.add_argument("--no-visual", action="store_true")
    parser.add_argument("--reference-docx")
    parser.add_argument("--formal-template",
                        help="compare a generated memo with the fixed formal-template layout contract")
    parser.add_argument("--project-as-of", help="project evidence cutoff in YYYY-MM-DD")
    parser.add_argument("--truth-overrides", metavar="JSON",
                        help="path to JSON list of source-traceable caller selections; each "
                             "requires key, value, source_uri, source_location, and evidence")
    args = parser.parse_args(argv)
    if args.inputs:
        truth_overrides = None
        if args.truth_overrides:
            truth_overrides = json.loads(Path(args.truth_overrides).read_text(encoding="utf-8"))
        summary = build_from_inputs(
            args.inputs, args.out, schema_path=args.schema, snapshot_id=args.snapshot_id,
            run_visual=not args.no_visual, public_evidence=args.evidence, work_dir=args.work_dir,
            narrative_path=args.narrative, reference_docx=args.reference_docx,
            formal_template=args.formal_template,
            truth_overrides=truth_overrides,
            project_as_of=args.project_as_of,
        )
    elif args.workbook and args.mapping:
        summary = build_bundle(
            args.workbook, args.mapping, args.out, args.template,
            snapshot_id=args.snapshot_id, evidence=args.evidence, work_dir=args.work_dir,
        )
    else:
        parser.error("use --inputs ... for the standard engine, or --workbook and --mapping for compatibility")
    print(json.dumps(summary, ensure_ascii=False))
    return 0 if summary["grade"] != "FAIL" else 1


if __name__ == "__main__":
    raise SystemExit(main())
