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


def test_missing_font_warned(monkeypatch):
    import pptx_lint
    monkeypatch.setattr(pptx_lint, "installed_fonts",
                        lambda: {"calibri", "arial"})
    prs = _prs()
    slide = _blank(prs)
    tb = _tb(slide, "hello", 1.0, 1.0)
    tb.text_frame.paragraphs[0].runs[0].font.name = "Gill Sans MT"
    issues = pptx_lint.lint_deck(prs)
    assert any("Gill Sans MT" in w for w in issues["warn"]), issues


def test_font_check_skipped_when_inventory_unavailable(monkeypatch):
    import pptx_lint
    monkeypatch.setattr(pptx_lint, "installed_fonts", lambda: None)
    prs = _prs()
    _tb(_blank(prs), "hello", 1.0, 1.0)
    issues = pptx_lint.lint_deck(prs)
    assert not any("not installed" in w for w in issues["warn"]), issues


def test_installed_fonts_returns_none_or_lowercase_set():
    from pptx_lint import installed_fonts
    fonts = installed_fonts()
    assert fonts is None or (
        isinstance(fonts, set) and all(f == f.lower() for f in fonts))


def test_jiggle_blames_outlier_not_majority():
    """Slide 1 is the outlier; only slide 1 should be blamed."""
    prs = _prs()
    # slide 1: left=11.4 (outlier); slides 2-4: left=11.9 (majority)
    for i, left in enumerate((11.4, 11.9, 11.9, 11.9), start=1):
        _tb(_blank(prs), str(i), left, 7.08)
    issues = lint_deck(prs)
    jiggle_errors = [e for e in issues["error"] if "jiggle" in e]
    assert len(jiggle_errors) == 1, jiggle_errors
    assert "Slide 1" in jiggle_errors[0], jiggle_errors[0]


def test_jiggle_no_false_positive_when_all_aligned():
    """Four perfectly aligned page numbers must produce zero jiggle errors."""
    prs = _prs()
    for i in range(1, 5):
        _tb(_blank(prs), str(i), 11.9, 7.08)
    issues = lint_deck(prs)
    assert not any("jiggle" in e for e in issues["error"]), issues


def test_page_gap_over_unnumbered_slide_passes():
    """A gap in page numbers spanning an unnumbered divider slide must not be flagged.

    Slides 1,2,4 carry numbers; slide 3 is an unnumbered layout.
    Number delta (4-2=2) equals slide-position delta (4-2=2), so no error.
    """
    prs = _prs()
    _tb(_blank(prs), "1", 11.9, 7.08)
    _tb(_blank(prs), "2", 11.9, 7.08)
    _blank(prs)                      # divider: no page number
    _tb(_blank(prs), "4", 11.9, 7.08)
    issues = lint_deck(prs)
    assert not any("sequence" in e for e in issues["error"]), issues


def test_lint_with_custom_palette(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    (pdir / "lintbrand.json").write_text(json.dumps({
        "bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
        "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
        "text": "F4F4F4", "text_muted": "9DB2BF", "dark": True}))
    load_custom_palettes(pdir)
    try:
        prs = _prs()
        _tb(_blank(prs), "branded", 1.0, 1.0, color="FEE715")
        issues = lint_deck(prs, palette_key="lintbrand")
        assert not any("unknown palette" in w for w in issues["warn"]), issues
        assert not any("FEE715" in e for e in issues["error"]), issues
    finally:
        PALETTES.pop("lintbrand", None)
