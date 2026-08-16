---
name: ni-memo
description: Use when an editable primary-market IC memo DOCX must be generated or audited from XLSX, DOCX, PDF, CSV, or JSON, especially when inputs lack a Database sheet or require traceability, formula review, public-source corroboration, template parity, or reproducible acceptance evidence.
---

# ni-memo: source-traceable investment memo engine

## Product contract

Generate one clean investment-memo product with the Haigude semantic system:

1. Project Summary
2. Transaction Terms
3. Investment Highlights & Risks
4. Industry Overview
5. Company Overview
6. Financials: history, forecast, returns, sensitivity
7. Appendix: comparables, sources, diligence gaps

Cover, updateable Word TOC field, headers, footers, status labels, and page numbers are document layers rather than business chapters. The master is independent of any historical memo. A historical DOCX is evidence and a visual reference; never rewrite its paragraphs in place.

## Non-negotiable rules

- Accept mixed XLSX, DOCX, PDF, CSV, and JSON inputs. Never require a worksheet named `Database`.
- Use only facts traceable to a source file, cell, paragraph, table cell, PDF line, JSON path, or public URL.
- `source_only` means present in supplied material, not independently verified.
- Preserve all conflicting candidates. Do not silently choose a convenient value.
- Separate historical project facts from later public-state updates using `as_of` and retrieval dates.
- Evaluate every field as `complete / partial / conflict / later_update / missing`; only `complete` enters the coverage numerator.
- Never turn a missing qualitative field into generated prose that sounds factual.
- Never call a formula population healthy merely because the parser reported zero exceptions. Count raw OOXML formula nodes, inspect cached blanks/errors, broken/external references, and perform non-destructive LibreOffice recalculation when available.
- Treat zero-cashflow IRR/MOIC errors as N/A only when the referenced cash-flow cells prove the condition.
- A rendered DOCX is not visually accepted until an internal temporary PDF and every page image exist and the contact sheet has been reviewed. These QA files are not deliverables.
- Character count, long-paragraph count, and body share are composition diagnostics, not truth or acceptance gates.
- Every generated memo ends with a fixed completion review: required/all-field progress, required/all-slot progress, grouped truth-state gaps, independent-verification coverage, formula status, and an unchecked author completion confirmation. The engine never marks a memo finished on the author's behalf.
- The formal editorial/layout contract is schema version 3.0.0. It fixes seven chapters, ordered semantic slots, six Word sections, styles, table controls, TOC, and PAGE fields while remaining independent from project facts.
- `PASS` requires all required facts and 100% independently verified/corroborated coverage. `PASS_WITH_NOTES` is usable but not signable. `FAIL` may still contain a useful draft and evidence bundle.

## Standard workflow

1. Inspect all supplied files and record their resolved paths.
2. Run the standard engine with a clean D-drive delivery directory and a separate D-drive work directory.
3. Review `mapping.json`, `facts.json`, `pending.md`, and `acceptance_report.json` in the work directory.
4. Distinguish:
   - fact exists in a source but discovery did not bind it: improve the generic alias/context rule and add a regression test;
   - fact is absent: keep it missing or collect public evidence;
   - sources disagree: preserve a conflict with every candidate and date.
5. If public corroboration is needed, browse authoritative or primary sources, create the evidence bundle below, and rerun with `--evidence`.
6. Inspect `visual/contact-sheet.jpg` in the work directory; open full-resolution pages for any dense or suspicious page.
7. Report the actual quality grade and unresolved gates. Do not upgrade the grade manually.

## Command

From the skill directory:

```powershell
python -m ni_memo.cli `
  --inputs "D:\path\model.xlsx" "D:\path\historical-memo.docx" `
  --out "D:\Codex\Outputs\ni-memo\project-name" `
  --work-dir "D:\path\work\project-name-YYYYMMDD" `
  --narrative "D:\path\narrative.json" `
  --truth-overrides "D:\path\truth_overrides.json" `
  --project-as-of "2021-09-30" `
  --snapshot-id "project-name-YYYYMMDD"
```

When `--work-dir` is omitted, the engine defaults to a sibling of the delivery
directory (`.ni-memo-work-<snapshot-id>`) or the `NI_MEMO_WORK_ROOT` env var.
No machine-specific paths are hard-coded.

To verify generated structure and formatting against the formal template, add:

```powershell
  --formal-template "D:\path\to\formal-template.docx"
```

Any supported combination works:

```powershell
python -m ni_memo.cli --inputs "D:\path\data.csv" "D:\path\deck.pdf" --out "D:\Codex\Outputs\ni-memo\project" --work-dir "D:\WorkBuddy\NeoStar\一级市场机会\_runs\ni-memo\project"
```

Use `--no-visual` only for unit tests or an explicitly requested diagnostic run. It is not a delivery mode.

The legacy explicit-mapping path remains available for reproducibility:

```powershell
python -m ni_memo.cli --workbook model.xlsx --mapping mapping.json --out output --work-dir "D:\WorkBuddy\NeoStar\一级市场机会\_runs\ni-memo\project"
```

## Narrative bundle (投资叙事素材)

A JSON sidecar supplies the professional investment narrative that raw models cannot express (business story, industry depth, risk mitigation). Pass with `--narrative`:

```json
{
  "summary.narrative": [
    {"title": "项目概况", "body": "无锡海古德新技术有限公司成立于2008年，围绕氮化铝陶瓷材料发展16年，是国内唯一规模量产氮化铝静电卡盘的企业……"}
  ],
  "case.highlights": [
    {"title": "静电卡盘研发难度大、客户导入壁垒高", "body": "静电卡盘有极高的工艺壁垒……"}
  ]
}
```

- Keys are schema `block.narrative_key` (e.g. `summary.narrative`, `case.highlights`, `case.risks`, `industry.narrative`, `company.*`, `financial.*`).
- Each item renders as a 加粗标题 + 论证段落 under the block's 节 (一、二、…) heading.
- When a block has narrative items, facts for that block render as tables/regions alongside; narrative never replaces source-traceable facts.
- The sidecar is input, not output: numbers inside narrative should match workbook truth or be explicitly marked 待确认.

## Model-derived analysis (模型驱动投资分析)

Standard generation automatically converts frozen financial, transaction, and return evidence into source-bound IC analysis. Each claim is explicitly labeled as model-derived analysis, carries exact lineage, and may explain trends, turning points, sensitivities, risks, and conditional investment implications. It never infers an absent company identity, product, customer, team, industry, or competitor fact from a financial curve.

The IC structure is compatible with professional investor review, but investor-skill is not a runtime dependency. ni-memo remains self-contained: `analysis.py` owns deterministic reasoning, `content.py` compiles typed facts and slots, and `render.py` presents the result without performing financial inference.

## Truth overrides (真值覆盖)

Model workbooks often contain intermediate values that look like deal terms. Use `--truth-overrides` (or `truth_overrides=` in Python) to select a documented caller value without pretending it is independently verified:

```json
[
  {
    "key": "deal.pre_money",
    "value": 13.6,
    "unit": "亿元",
    "source_uri": "D:\\path\\ic-approved-terms.docx",
    "source_location": "交易方案表/投前估值",
    "evidence": "投前估值13.6亿元"
  }
]
```

- Require `source_uri`, `source_location`, and non-empty `evidence`; `source_uri` must be an actual input in the run.
- Record every override as caller-selected SOURCE_ONLY `investor_input`. Ignore a caller-declared authoritative source type. Selection never self-verifies.
- Preserve a conflict when the selected value disagrees with another same-period candidate; policy choice and verification are separate decisions.
- Apply an override normally only to a discovered key. Apply `supplement_missing: true` only to a defined schema field and only with the same complete lineage requirements.
- Use overrides for documented deal-term selection or source-backed gap binding, not to hide a workbook defect.

## Formula gate (公式门禁)

The engine counts raw OOXML formula nodes, inspects cached blanks/errors/broken references, and runs non-destructive LibreOffice recalculation when available. Demote a source-intrinsic defect to WARNING only when the exact same workbook and cell is defective before and after recalculation and that cell is not the selected source of any memo fact. Recalculation status alone cannot demote a defect. A recalculation engine failure, a location mismatch, an unconfirmed defect, or a selected-fact defect remains FAIL.

## Workbook regions (模型语义区域)

The engine profiles every input XLSX into semantic regions for audit and deterministic model analysis. The raw workbook regions stay in --work-dir and never enter the default delivery DOCX. The memo renders only frozen semantic facts, selected decision tables, charts derived from frozen period series, sourced narratives, and source-bound analysis; this prevents page-count inflation and raw-model dumps.

## Formal-template mode (正式模板)

Pass `--formal-template path\to\formal-template.docx` during normal generation. The engine compares all six Word-section layouts, ordered seven chapters, ordered required semantic slots, Normal/Heading/Table styles, repeatable table headers, non-splitting data rows, the updateable TOC field, and the footer page-number field. Project facts and page count may vary; factual content is never copied across projects. The resulting `formal_template_report.json` is an acceptance gate.

## Formal-reference preservation mode (正式参考文档保真)

Pass `--reference-docx path\to\historical.docx` to preserve an explicitly supplied historical DOCX byte-for-byte (no XML rewrite) and run `compare_document_fidelity` on the copy. The standard fact/formula/visual gates still run independently. This mode is for audit/versioning, not for generating a new memo from a template.

## Public evidence bundle

The CLI does not invent or scrape facts. The agent may collect current public evidence and pass a JSON list with `--evidence`:

```json
[
  {
    "key": "company.registered_capital",
    "value": 9294.64,
    "unit": "万元",
    "as_of": "2024-09-27",
    "url": "https://authoritative.example/company-record",
    "source_type": "government",
    "source_location": "registration record",
    "retrieved_at": "2026-08-15",
    "evidence": "concise supporting excerpt or description"
  }
]
```

Placeholder URLs, missing retrieval dates, and non-HTTP(S) URLs are rejected. Two domains with the same value can corroborate a fact; different values remain a conflict.

## Delivery and internal work

The delivery directory contains exactly one user-facing artifact:

- `memo.docx`: editable fixed seven-chapter memo generated from a clean master

Keep all audit evidence in `--work-dir`; do not link or deliver it unless the user explicitly asks:

- `facts.json`: frozen selected facts and every candidate, with exact lineage
- `mapping.json`: discovery rule and binding evidence
- `input_report.json`: accepted/skipped files and public sources
- `pending.md`: missing facts, conflicts, and formula gates
- `acceptance_report.json`: content, completion, formula, recalculation, formal-template, document-composition, and visual results
- `formal_template_report.json`: semantic layout comparison against the explicitly supplied formal template
- `run_summary.json`: top-level result sharing the same snapshot ID
- `recalculated/`: LibreOffice-generated workbook copies; original files remain unchanged
- `visual/memo.pdf`, `visual/page-*.png`, `visual/contact-sheet.jpg`: temporary rendering evidence

Formula and mapping acceptance must use XLSX/source evidence and the frozen fact snapshot, never the rendered PDF. All internal artifacts must use the same `snapshot_id`.

## Architecture seams

- `schema/standard_ic_memo.json`: versioned chapters, blocks, fields, aliases, requirements
- `ni_memo/ingest.py`: source-neutral sparse extraction
- `ni_memo/discover.py`: deterministic semantic binding; ambiguity stays visible
- `ni_memo/reconcile.py`: source/date-aware selection and conflict preservation
- `ni_memo/formula_audit.py`: independent formula-population and cache gates
- `ni_memo/recalculate.py`: non-destructive LibreOffice verification
- `ni_memo/workbook_profile.py`: sparse semantic region profiling of XLSX models
- `ni_memo/content.py`: compile schema + frozen facts into typed memo content blocks
- `ni_memo/analysis.py`: compose deterministic, source-bound model analysis and diligence implications
- `ni_memo/narratives.py`: narrative sidecar loading (投资叙事素材)
- `ni_memo/chart.py`: chart PNG rendering for chart-typed blocks
- `ni_memo/document_quality.py`: generated-memo composition, formal-template semantics, and byte-preservation fidelity gates
- `ni_memo/render.py`: clean fixed-system DOCX rendering (海古德版式: A4/仿宋 11pt/Table Grid/封面/TOC/页脚页码)
- `ni_memo/visual.py`: PDF, page PNG, contact sheet, and structural visual gates
- `ni_memo/cli.py`: one coherent artifact bundle (narrative + truth-overrides + project cutoff + formal-template acceptance)

Do not add a project-specific binder, hard-coded cell map, or copied historical template to the general engine. A new exception must be expressed as a generic semantic rule with a failing regression test first.

## Regression expectations

- Haigude mixed DOCX+XLSX must recover company identity, dated business narrative, highlights, risks, transaction terms, financial series, IRR/MOIC, and exact lineage. It must remain `FAIL` while registered capital conflicts or genuine formula defects remain.
- Project GL v46 XLSX must generate the complete document structure and the model-backed transaction/financial/return sections. Missing company/product/customer/team/competition facts must remain explicit rather than fabricated.
- Both samples must render every page with no blank pages, missing chapter headings, or unresolved template tokens.

## Engine changelog (2.5.1 — 2026-08-16)

Deterministic binding fixes validated against Haigude golden-reference and Project GL v46:

1. **Formula-gate merge ordering (cli.py)**: formula audit now runs on the *post-recalculation* item set. The previous pre-merge snapshot compared stale cached `#NUM!` evidence against fresh recalculated values and mis-failed workbooks whose XIRR caches were simply outdated (GL: 4 XIRR cells).2. **Truth-override conflict downgrade (reconcile.py)**: a policy hit (`prefer_source_type: investor_input`) keeps `SOURCE_ONLY` instead of being re-marked `CONFLICT` merely because multiple candidate values exist. The investor's explicit choice is authoritative, not a disagreement.
3. **Entity-dimension header paths (ingest.py `_nearest_header`)**: two-level column headers (year row + entity row 无锡/东台/合并) are joined into one path (`2022A·无锡·审计数`) so distinct entities become distinct series keys instead of fabricated conflicts. Header scan restricted to rows ≤ 10 so stray data-zone text never shadows the real header.
4. **Row-breakdown key suffixes (discover.py `_row_dimension`)**: labels naming real sub-buckets (高铁/城轨/合计市场规模, IRR/MOIC sensitivity rows) become `·suffix` keys; generic labels (毛利率/净利润/货币资金) stay undivided.
5. **Section-title context (ingest.py `_section_title`)**: block titles (项目概览/退出假设/本次投资/股权结构) are injected as the first context token, so the `deal.pre_money/post_money` guard can distinguish the 投前/投后 valuation in 项目概览 from the exit valuation under 退出假设 (GL D11=3700 vs D45=8231). Block titles are text rows with no numeric value to the right, below row 4, excluding column-header tokens.
6. **CJK two-character alias matching (discover.py `_match`)**: `contained_context_alias` accepts `len(target) >= 2` (投前/投后/估值 are 2 CJK chars); the old 3-char floor made `deal.pre_money` permanently unbound (GL C11=2700).
7. **Semantic guards (discover.py `_source_compatible`)**: `market.growth` requires the label itself to express market/industry semantics (财务行 YOY/CAGR 排除); `deal.pre_money/post_money` rejects 退出/换股/收购前 valuation rows; `business.capacity` rejects sensitivity-analysis sheets; `company.registered_capital` rejects 剩余/受让/对应/转让 rows; `financial.net_profit/revenue/gross_margin` reject mislabeled rows (所得税/营业成本/EBITDA 利润率).
8. **Golden-reference grade (cli.py)**: with `--reference-docx`, completion < 100% is expected (reference narrative assets are not xlsx-derivable) — grade by `verified_rate` (PASS at 100%, else PASS_WITH_NOTES); formula failures still fail absolutely.
9. **Money-unit conversion (analysis.py `_money`)**: `百万元 → 亿元` used `/10`, inflating every analysis-paragraph amount 10x (GL pre-money rendered as 270亿 instead of 27亿; Haigude investment as 3亿 instead of 0.3亿). Corrected to `/100` (1000 万元 = 1 亿元). Table cells already rendered raw value+unit and were unaffected; only narrative analysis paragraphs were wrong. Found via desktop GL v46 end-to-end test; regression-checked on both GL and Haigude (Haigude 13.6亿 pre / 16.6亿 post / 0.3亿 investment now correct).
10. **Cross-block topic dedup (analysis.py)**: the same fact theme (IRR/MOIC, 投前投后估值, 退出PE, 现金波动, 负债率) was composed into near-identical paragraphs in 4 blocks (summary/case/company/financial), producing 27 model paragraphs (3335 chars) of ~100-char boilerplate repeats. `AnalysisClaim` gained a `topic` field; `_dedupe_topics()` keeps only the highest-priority occurrence per topic using a fixed block-priority table (`summary.narrative 0 > transaction.terms 1 > transaction.ownership 2 > case.highlights 3 > case.risks 4 > financial.returns 5 > financial.sensitivity 6 > financial.forecast 7 > industry.narrative 8 > company.basic 9`). 40 `add()` call sites now carry explicit topics. GL v46 dedup3: 27 → 14 model paragraphs (-61% chars), every fact theme kept exactly once, zero information loss.
11. **`_series` regex suffix compatibility (analysis.py)**: 2.5.1's `_row_dimension` rewrote financial keys to `financial.revenue.2018a·营业收入`, so the old `^base\.(20\d{2})([aAeE])$` anchor no longer matched and revenue_a/revenue_e returned empty — silently killing every revenue-based analysis paragraph. `_series` now matches the bare key first, then falls back to the first per-period suffixed variant (period dedup via set). Revenue/EBITDA/asset series paragraphs restored on GL dedup3.
12. **Portable work-dir default (cli.py `_resolve_work_dir`)**: the default work directory hard-coded `D:\WorkBuddy\NeoStar\一级市场机会\_runs\ni-memo`, which breaks distribution to another machine (the friend's Codex install). Now defaults to a sibling of the delivery directory (`.ni-memo-work-<snapshot-id>`) or the `NI_MEMO_WORK_ROOT` env var. Zero hard-coded machine paths remain in the engine.
13. **Distribution kit (2026-08-16)**: added `requirements.txt` (python-docx/lxml/openpyxl/Pillow/pypdf/matplotlib), `README.md` (install + one-command run + acceptance guide), and `tests/test_ni_memo_guards.py` (17 tests: `_money` 10x guard, `_dedupe_topics`, `_series` suffix fallback, hardcoded-path scan). The hardcoded-path test immediately caught the cli.py bug in item 12.

Verified outcomes:
- Haigude golden: `PASS_WITH_NOTES`, fidelity `pass` (byte-identical: 24528 chars / 52 tables / 54 inline shapes / 51 media / 6 sections / 20 footnotes / 64 page breaks), formula `warning` (source-inherent defects captured).
- GL v46 full: `deal.pre_money=2700` (Return Summary!C11), `deal.post_money=3700` (IS!I46), no transaction conflicts; 10 narrative-backed identity fields bound; remaining missing fields (registered capital, market size/growth/comparables) are genuine source gaps, reported explicitly.
- GL v46 desktop e2e (2026-08-16): bare xlsx-only run (no narrative/overrides) — 224 facts, 8 pending identity/qualitative gaps reported honestly, deal terms bound exactly (pre=2700@C11, post=3700@I46, entry=12@E10, invest=1000@D15, IRR=27.45%@Y51, MOIC=2.27x@Y52), 26-page docx with source-bound analysis. Post-fix amounts verified correct (投资额10亿/投前27亿/投后37亿).
- GL v46 identity-completion e2e (2026-08-16, e2e-jinli-dedup3): with only the project name "河北金力" supplied, narrative.json + truth_overrides.json (11 investor_input identity overrides) filled every qualitative field — pending dropped 8 → 1 (registered capital only, a genuine model gap); fact_count 224 → 235; memo正文 carries 袁海朝实控人/比亚迪股东/河北金力+安徽金力基地/东镐换股等完整身份叙事. Model analysis after dedup: 14 paragraphs covering revenue contour (1.52→2.13亿 history, 6.89→32.28亿 forecast, CAGR 47%), profit turnaround, deal structure, IRR 27.45%/MOIC 2.27x, stake reconciliation (27.03% vs 2.01%), gross margin (-6.42%→43.20%), asset expansion (26.28→74.53亿, 2.84x), cash volatility, leverage (69.15%→10%→27.58%), exit PE 30x, EBITDA 1.92→7.91亿, exit valuation 82.32→127.48亿 — amounts all correct.
