"""Tests for diff_deck.py — outline transforms and markup-aware matching."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from build_deck import build  # noqa: E402
from diff_deck import diff_outline  # noqa: E402

TRANSFORMED_MD = """**Auto-Agenda:** track
**References:** on

## Slide 1: Acme FY27 Strategy
**Layout:** title
- Title: Acme FY27 Strategy
- Subtitle: Board readout

## Slide 2: Diagnosis
**Layout:** section-divider
- Subtitle: Where we stand today

## Slide 3: Costs have outgrown revenue for three years
- Cost CAGR 12% vs revenue 4%
- Source: Company filings
- Notes: Set up the problem.

## Slide 4: Plan
**Layout:** section-divider
- Subtitle: What we will do next

## Slide 5: Three levers close the gap by FY27
- Tiering, exit, discount discipline
- Notes: The plan.
"""

MARKUP_MD = """## Slide 1: Acme FY27 Strategy
**Layout:** title
- Title: Acme FY27 Strategy
- Subtitle: Board readout

## Slide 2: Costs have outgrown revenue for three years
- icon:trending-up **Cost CAGR 12%** against revenue
- {accent}Revenue CAGR 4%{/} across three years
- Notes: Set up the problem.
"""


def _build_and_diff(md_text, tmp_path):
    outline = tmp_path / "outline.md"
    outline.write_text(md_text, encoding="utf-8")
    deck = tmp_path / "deck.pptx"
    assert build(outline, deck)
    return diff_outline(outline, deck)


def test_diff_applies_auto_agenda_and_references(tmp_path):
    """Auto-Agenda: track + References: on shift every slide position;
    diff must apply the same transforms build() does — no false warnings."""
    errors, warnings = _build_and_diff(TRANSFORMED_MD, tmp_path)
    assert not errors, errors
    assert not any("Slide count mismatch" in w for w in warnings), warnings
    assert not any("expected text not found" in w for w in warnings), warnings


def test_diff_strips_icon_and_rich_markup(tmp_path):
    """icon: prefixes and **…** / {accent}…{/} markers never appear in the
    rendered deck text — expected phrases must be compared stripped."""
    errors, warnings = _build_and_diff(MARKUP_MD, tmp_path)
    assert not errors, errors
    assert not any("expected text not found" in w for w in warnings), warnings
