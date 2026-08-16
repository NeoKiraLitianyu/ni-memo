"""Regression tests for ni-memo critical lines of defense.

Four guards that must never silently regress:
1. _money unit conversion (百万元 → 亿元 must use /100, not /10 — P0 10x bug)
2. _dedupe_topics cross-block dedup (27 → 14 paragraphs, zero info loss)
3. _series regex suffix compatibility (2.5.1 row-breakdown keys)
4. No hardcoded absolute paths in the engine (distribution portability)

Run: python -m pytest tests/ -q   (from the ni-memo directory)
"""
import importlib
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from ni_memo.analysis import AnalysisClaim, _dedupe_topics, _money, _series  # noqa: E402
from ni_memo.facts import FactRecord  # noqa: E402

# ---------------------------------------------------------------------------
# 1. _money unit conversion (P0 guard)
# ---------------------------------------------------------------------------


class TestMoneyConversion:
    def test_baiwan_to_yi_uses_100(self):
        """1000 万元 = 1 亿元. 百万元 must divide by 100 (100 百万元 = 1 亿元)."""
        assert _money(1000, "百万元") == "10.00亿元"
        assert _money(100, "百万元") == "1.00亿元"
        assert _money(2700, "百万元") == "27.00亿元"  # GL pre-money
        assert _money(1000, "百万元") == "10.00亿元"  # GL investment

    def test_wan_to_yi(self):
        assert _money(10000, "万元") == "1.00亿元"
        assert _money(3000, "万元人民币") == "0.30亿元"  # Haigude investment

    def test_yi_passthrough(self):
        assert _money(13.6, "亿元") == "13.60亿元"

    def test_bare_unit_passthrough(self):
        assert _money(2.5, "x") == "2.50x"

    def test_regression_10x_bug(self):
        """The infamous /10 bug inflated every amount 10x. Never allow /10 back."""
        # 13.6亿元 must NOT become 136亿元
        assert _money(13.6, "亿元") != "136.00亿元"
        # 1000 百万元 = 10亿, NOT 100亿
        assert _money(1000, "百万元") != "100.00亿元"


# ---------------------------------------------------------------------------
# 2. _dedupe_topics cross-block dedup
# ---------------------------------------------------------------------------


def make_claim(block, topic, title="t", body="b"):
    return AnalysisClaim(
        block_key=block,
        title=title,
        body=body,
        sources=("x#1",),
        topic=topic,
    )


class TestDedupeTopics:
    def test_same_topic_keeps_highest_priority(self):
        """summary.narrative (0) wins over financial.returns (5) for same topic."""
        items = {
            "summary.narrative": (make_claim("summary.narrative", "return_profile"),),
            "financial.returns": (make_claim("financial.returns", "return_profile"),),
        }
        out = _dedupe_topics(items)
        assert len(out["summary.narrative"]) == 1
        assert out["financial.returns"] == ()
        assert out["summary.narrative"][0].block_key == "summary.narrative"

    def test_different_topics_all_kept(self):
        items = {
            "summary.narrative": (make_claim("summary.narrative", "revenue_growth"),),
            "financial.returns": (make_claim("financial.returns", "return_profile"),),
            "case.risks": (make_claim("case.risks", "leverage"),),
        }
        out = _dedupe_topics(items)
        total = sum(len(c) for c in out.values())
        assert total == 3

    def test_untopiced_claims_never_dropped(self):
        """Claims without topic must never be dropped by dedup."""
        items = {
            "summary.narrative": (make_claim("summary.narrative", "", "untopiced"),),
            "financial.returns": (make_claim("financial.returns", "return_profile"),),
        }
        out = _dedupe_topics(items)
        assert len(out["summary.narrative"]) == 1

    def test_priority_ordering_full(self):
        """Full priority table: summary > transaction > case > financial > industry > company."""
        blocks = [
            "summary.narrative", "transaction.terms", "transaction.ownership",
            "case.highlights", "case.risks", "financial.returns",
            "financial.sensitivity", "financial.forecast", "industry.narrative",
            "company.basic",
        ]
        items = {b: (make_claim(b, "same_topic"),) for b in blocks}
        out = _dedupe_topics(items)
        survivors = [b for b, c in out.items() if c]
        assert survivors == ["summary.narrative"]
        assert out["summary.narrative"][0].block_key == "summary.narrative"


# ---------------------------------------------------------------------------
# 3. _series regex suffix compatibility
# ---------------------------------------------------------------------------


class FakeSnapshot:
    """Minimal stand-in exposing .selected as dict[str, FactRecord]."""

    def __init__(self, selected):
        self.selected = selected


def fact(key, value, unit=None):
    return FactRecord(
        key=key,
        value=value,
        unit=unit,
        as_of=None,
        source_type="xlsx",
        source_uri="model.xlsx",
        source_location="Sheet!A1",
        extraction_method="test",
        confidence="high",
        status="source_only",
    )


class TestSeries:
    def test_bare_keys(self):
        snap = FakeSnapshot({
            "financial.revenue.2018a": fact("financial.revenue.2018a", 1.52, "亿元"),
            "financial.revenue.2019e": fact("financial.revenue.2019e", 6.89, "亿元"),
        })
        assert _series(snap, "financial.revenue", "A") == [(2018, 1.52, "亿元", "financial.revenue.2018a")]
        assert _series(snap, "financial.revenue", "E") == [(2019, 6.89, "亿元", "financial.revenue.2019e")]

    def test_suffixed_keys_fallback(self):
        """2.5.1 row-breakdown keys (·营业收入) must match via suffixed fallback."""
        snap = FakeSnapshot({
            "financial.revenue.2018a·营业收入": fact("financial.revenue.2018a·营业收入", 1.52, "亿元"),
            "financial.revenue.2019e·营业收入": fact("financial.revenue.2019e·营业收入", 6.89, "亿元"),
        })
        assert _series(snap, "financial.revenue", "A") == [(2018, 1.52, "亿元", "financial.revenue.2018a·营业收入")]
        assert _series(snap, "financial.revenue", "E") == [(2019, 6.89, "亿元", "financial.revenue.2019e·营业收入")]

    def test_mixed_bare_preferred(self):
        """When both bare and suffixed exist, bare wins; no duplicate periods."""
        snap = FakeSnapshot({
            "financial.revenue.2018a": fact("financial.revenue.2018a", 1.52, "亿元"),
            "financial.revenue.2018a·营业收入": fact("financial.revenue.2018a·营业收入", 9.99, "亿元"),
            "financial.revenue.2019e": fact("financial.revenue.2019e", 6.89, "亿元"),
            "financial.revenue.2019e·营业收入": fact("financial.revenue.2019e·营业收入", 8.88, "亿元"),
        })
        a = _series(snap, "financial.revenue", "A")
        e = _series(snap, "financial.revenue", "E")
        assert len(a) == 1 and a[0][1] == 1.52  # bare wins, no duplicate period
        assert len(e) == 1 and e[0][1] == 6.89

    def test_multi_suffix_same_period_dedup(self):
        """Multiple suffixed variants of one period collapse to the first."""
        snap = FakeSnapshot({
            "financial.revenue.2021a·营业收入": fact("financial.revenue.2021a·营业收入", 1.0, "亿元"),
            "financial.revenue.2021a·其他": fact("financial.revenue.2021a·其他", 2.0, "亿元"),
        })
        a = _series(snap, "financial.revenue", "A")
        assert len(a) == 1

    def test_non_numeric_excluded(self):
        snap = FakeSnapshot({
            "financial.revenue.2018a·营业收入": fact("financial.revenue.2018a·营业收入", "N/A"),
        })
        assert _series(snap, "financial.revenue", "A") == []


# ---------------------------------------------------------------------------
# 4. No hardcoded absolute paths (distribution portability)
# ---------------------------------------------------------------------------


class TestNoHardcodedPaths:
    ENGINE_DIR = Path(__file__).resolve().parent.parent / "ni_memo"

    def test_engine_has_no_absolute_paths(self):
        pattern = re.compile(r'["\'](?:[A-Za-z]:[\\/]|/Users/|/home/|//)')
        violations = []
        for py in sorted(self.ENGINE_DIR.glob("*.py")):
            if py.name == "__init__.py":
                continue
            for i, line in enumerate(py.read_text(encoding="utf-8").splitlines(), 1):
                if pattern.search(line) and "LibreOffice" not in line:
                    # Allow the LibreOffice candidate paths (they are existence probes)
                    violations.append(f"{py.name}:{i}: {line.strip()}")
        assert not violations, f"Hardcoded paths found:\n" + "\n".join(violations)

    def test_requirements_exists(self):
        req = Path(__file__).resolve().parent.parent / "requirements.txt"
        assert req.exists(), "requirements.txt missing — distribution blocker"

    def test_readme_exists(self):
        readme = Path(__file__).resolve().parent.parent / "README.md"
        assert readme.exists(), "README.md missing — distribution blocker"
