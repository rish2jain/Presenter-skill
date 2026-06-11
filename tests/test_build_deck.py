"""Tests for outline parsing, validation, and golden-path build."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from build_deck import parse_outline, validate, _parse_data_point, _series_count  # noqa: E402
import builders  # noqa: E402


@pytest.fixture(autouse=True)
def _default_density():
    builders.set_density("compact")


CTX = {"outline_dir": ROOT / "assets", "assets_dir": ROOT / "assets"}


def test_data_block_tolerates_blank_lines():
    md = """## Slide 1: Chart
**Layout:** two-column-split
**Visual:** chart:bar
**Data:**
- 2024: 42

- 2025: 67
- Heading: Growth
- bullet one
"""
    _, slides = parse_outline(md)
    assert slides[0]["data"] == [("2024", 42.0), ("2025", 67.0)]
    assert slides[0]["heading"] == "Growth"
    assert slides[0]["bullets"] == ["bullet one"]


def test_multi_series_data_parsing():
    md = """## Slide 1: Chart
**Layout:** two-column-split
**Visual:** chart:line
**Series:** Revenue, Costs
**Data:**
- Q1: 4.2, 3.1
- Q2: 5.1, 3.0
"""
    _, slides = parse_outline(md)
    assert _series_count(slides[0]) == 2
    assert slides[0]["data"] == [("Q1", [4.2, 3.1]), ("Q2", [5.1, 3.0])]
    errors, _ = validate(slides, CTX)
    assert not errors


def test_multi_series_mismatch_errors():
    md = """## Slide 1: Chart
**Layout:** two-column-split
**Visual:** chart:line
**Series:** Revenue, Costs
**Data:**
- Q1: 4.2
"""
    _, slides = parse_outline(md)
    errors, _ = validate(slides, CTX)
    assert any("multi-series" in e for e in errors)


def test_parse_data_point_dollar_amounts():
    assert _parse_data_point("2024: $42B") == ("2024", 42.0)
    assert _parse_data_point("Q1: -5.2") == ("Q1", -5.2)


def test_validate_warns_missing_notes():
    md = """## Slide 1: X
**Layout:** bullet-list
- one
"""
    _, slides = parse_outline(md)
    _, warnings = validate(slides, CTX)
    assert any("Notes" in w for w in warnings)


def test_validate_title_heading_fallback_warning():
    md = """## Slide 1: My Title
**Layout:** title
"""
    _, slides = parse_outline(md)
    _, warnings = validate(slides, CTX)
    assert any("My Title" in w for w in warnings)


def test_validate_rejects_chart_on_comparison_side():
    md = """## Slide 1: Compare
**Layout:** comparison
**Visual-Left:** chart:bar
- Left label: A
- Right label: B
"""
    _, slides = parse_outline(md)
    errors, _ = validate(slides, CTX)
    assert any("visual_left" in e for e in errors)


def test_example_outline_check():
    outline = ROOT / "assets" / "example-outline.md"
    meta, slides = parse_outline(outline.read_text(encoding="utf-8"))
    errors, _ = validate(slides, CTX)
    assert not errors
    assert len(slides) == 6


def test_golden_build_and_qa(tmp_path):
    outline = ROOT / "assets" / "example-outline.md"
    out = tmp_path / "deck.pptx"
    from build_deck import build

    assert build(outline, out, check_only=False)
    assert out.is_file()

    from qa_check import check_deck

    issues = check_deck(out)
    assert not issues["error"]
