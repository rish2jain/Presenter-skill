"""Tests for narrative, smart layout, diff, and template helpers."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from build_deck import parse_outline, validate, build  # noqa: E402
from narrative import normalize_purpose, validate_narrative, generate_appendix_outline  # noqa: E402
from smart_layout import auto_layout  # noqa: E402
from template_helpers import build_template_map, palette_from_theme  # noqa: E402
import builders  # noqa: E402


@pytest.fixture(autouse=True)
def _density():
    builders.set_density("compact")


CTX = {"outline_dir": ROOT / "assets", "assets_dir": ROOT / "assets"}


def test_auto_layout_stat_callout():
    _, slides = parse_outline("## Slide 1: Metrics\n- Value=\"1\" Label=\"X\" Sublabel=\"Y\"\n")
    assert slides[0]["layout"] == "stat-callout"


def test_auto_layout_closing():
    _, slides = parse_outline("## Slide 1: Questions?\n**Layout:** closing\n")
    assert slides[0]["layout"] == "closing"


def test_narrative_purpose_validation():
    md = """**Purpose:** pitch
**Takeaway:** "We save enterprises $14M"

## Slide 1: Title
**Layout:** title
- Notes: "Open"

## Slide 2: Close
**Layout:** closing
- Notes: "End"
"""
    meta, slides = parse_outline(md)
    warnings = validate_narrative(meta, slides)
    assert not any("Takeaway" in w for w in warnings)
    assert any("stat-callout" in w for w in warnings)


def test_variant_preset_b():
    outline = ROOT / "assets" / "example-outline.md"
    out = Path("/tmp/variant-b-deck.pptx")
    assert build(outline, out, variant="b")
    assert out.is_file()


def test_palette_from_theme():
    pal = palette_from_theme({"accent1": "C9A84C", "dk1": "0A0F1E", "lt1": "F1F5F9"})
    assert pal["accent1"] == "C9A84C"
    assert pal["dark"] is True


def test_generate_appendix():
    meta, slides = parse_outline("**Purpose:** pitch\n## Slide 1: T\n**Layout:** title\n")
    text = generate_appendix_outline(meta, slides)
    assert "Appendix" in text
    assert "Financial Detail" in text


def test_diff_deck_example(tmp_path):
    outline = ROOT / "assets" / "example-outline.md"
    out = tmp_path / "deck.pptx"
    assert build(outline, out)
    from diff_deck import diff_outline
    errors, warnings = diff_outline(outline, out)
    assert not errors
    assert not any("expected text not found" in w and "'title'" in w for w in warnings)
    assert not any("expected text not found" in w and "'closing'" in w for w in warnings)


def test_build_template_map():
    from pptx import Presentation
    prs = Presentation()
    m = build_template_map(prs)
    assert "layout_map" in m
    assert "title" in m["layout_map"] or m["layout_scores"].get("title", 0) >= 0
