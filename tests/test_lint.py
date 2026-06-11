"""Tests for pptx_lint.py cross-slide consistency checks."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

from pptx_lint import lint_deck  # noqa: E402


def _prs():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    return prs


def _blank(prs):
    layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    return prs.slides.add_slide(layout)


def _tb(slide, text, left, top, w=0.8, h=0.35, size=11, color=None):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    if color:
        from pptx.dml.color import RGBColor
        run.font.color.rgb = RGBColor.from_string(color)
    return tb


def test_jiggle_flagged_for_misaligned_page_numbers():
    prs = _prs()
    for n, left in enumerate((11.9, 11.9, 11.4)):  # third one jiggles
        _tb(_blank(prs), str(n + 1), left, 7.08)
    issues = lint_deck(prs)
    assert any("jiggle" in e for e in issues["error"]), issues


def test_aligned_page_numbers_pass():
    prs = _prs()
    for n in range(3):
        _tb(_blank(prs), str(n + 1), 11.9, 7.08)
    issues = lint_deck(prs)
    assert not any("jiggle" in e for e in issues["error"]), issues


def test_page_sequence_gap_flagged():
    prs = _prs()
    for label in ("1", "2", "4"):
        _tb(_blank(prs), label, 11.9, 7.08)
    issues = lint_deck(prs)
    assert any("sequence" in e for e in issues["error"]), issues


def test_font_explosion_warned():
    prs = _prs()
    slide = _blank(prs)
    for i, fname in enumerate(
            ("Calibri", "Arial", "Georgia", "Verdana", "Impact", "Tahoma")):
        tb = _tb(slide, f"text {i}", 1.0, 1.0 + i * 0.5)
        tb.text_frame.paragraphs[0].runs[0].font.name = fname
    issues = lint_deck(prs)
    assert any("font" in w.lower() for w in issues["warn"]), issues


def test_palette_whitelist_flags_off_palette_color():
    prs = _prs()
    slide = _blank(prs)
    _tb(slide, "on palette", 1.0, 1.0, color="C9A84C")   # midnight accent1
    _tb(slide, "rogue", 1.0, 2.0, color="FF00FF")
    issues = lint_deck(prs, palette_key="midnight-executive")
    assert any("FF00FF" in e for e in issues["error"]), issues
    assert not any("C9A84C" in e for e in issues["error"]), issues
