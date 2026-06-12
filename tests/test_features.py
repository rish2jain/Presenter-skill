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


def test_thumbnail_grid_is_labeled(tmp_path):
    from PIL import Image
    from render_slides import create_thumbnail_grid
    imgs = []
    for i in range(3):
        p = tmp_path / f"slide-{i + 1:02d}.png"
        Image.new("RGB", (400, 225), (40, 40, 60)).save(p)
        imgs.append(p)
    out = tmp_path / "grid.png"
    create_thumbnail_grid(imgs, out, cols=2)
    grid = Image.open(out)
    # badge: each cell's top-left corner carries an opaque label strip
    assert grid.getpixel((4, 4)) == (0, 0, 0), "badge background not black at (4,4)"


def test_every_palette_has_chart_series_token():
    from palettes import PALETTES
    for key, pal in PALETTES.items():
        assert len(pal.get("chart_series", [])) >= 3, key


def test_load_custom_palette(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    (pdir / "acme-brand.json").write_text(json.dumps({
        "bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
        "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
        "text": "F4F4F4", "text_muted": "9DB2BF", "dark": True}))
    loaded = load_custom_palettes(pdir)
    try:
        assert "acme-brand" in loaded and "acme-brand" in PALETTES
        pal = PALETTES["acme-brand"]
        assert pal["font_title"]                      # font defaults filled
        assert len(pal["chart_series"]) >= 3          # token derived
    finally:
        PALETTES.pop("acme-brand", None)              # don't leak into other tests


def test_invalid_custom_palette_rejected(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    (pdir / "broken.json").write_text(json.dumps({"bg": "101820"}))
    loaded = load_custom_palettes(pdir)
    assert "broken" not in loaded and "broken" not in PALETTES


def test_custom_palette_with_non_hex_colors_rejected(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    (pdir / "named.json").write_text(json.dumps({
        "bg": "red", "bg_deep": "0A0F14", "surface": "1E2A33",
        "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
        "text": "F4F4F4", "text_muted": "9DB2BF", "dark": True}))
    assert load_custom_palettes(pdir) == []
    assert "named" not in PALETTES


def test_custom_palette_rejects_short_chart_series(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    base = {
        "bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
        "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
        "text": "F4F4F4", "text_muted": "9DB2BF", "dark": True,
        "chart_series": ["FEE715", "8DA9C4"],
    }
    (pdir / "short-series.json").write_text(json.dumps(base))
    assert load_custom_palettes(pdir) == []
    assert "short-series" not in PALETTES


def test_sticker_and_kicker_parsed():
    md = ('## Slide 1: Revenue doubled on pricing discipline\n'
          '- Sticker: Illustrative\n'
          '- Kicker: "Lock in pricing gains before renewals"\n')
    _, slides = parse_outline(md)
    assert slides[0]["sticker"] == "Illustrative"
    assert slides[0]["kicker"] == "Lock in pricing gains before renewals"


def test_kicker_restating_title_warned():
    md = ('## Slide 1: Revenue doubled across all regions this year\n'
          '**Layout:** bullet-list\n'
          '- Kicker: "Revenue doubled across all regions"\n'
          '- Point 1: "Growth was broad-based"\n'
          '- Notes: "n"\n')
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert any("kicker restates the title" in w for w in warnings), warnings


def test_kicker_advancing_argument_not_warned():
    md = ('## Slide 1: Revenue doubled across all regions this year\n'
          '**Layout:** bullet-list\n'
          '- Kicker: "Approve the follow-on investment before quarter close"\n'
          '- Point 1: "Growth was broad-based"\n'
          '- Notes: "n"\n')
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert not any("kicker restates" in w for w in warnings), warnings


def test_heading_over_15_words_warned():
    heading = " ".join(["word"] * 16)
    md = (f'## Slide 1: {heading}\n'
          '**Layout:** bullet-list\n- Point 1: "x"\n- Notes: "n"\n')
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert any("exceeds 15 words" in w for w in warnings), warnings


def test_exhibit_heading_without_number_warned():
    md = ('## Slide 1: Margins expanded sharply on pricing discipline\n'
          '**Layout:** waterfall\n'
          '**Data:**\n- FY24: 42\n- Pricing: +9\n- FY25: total\n'
          '- Notes: "n"\n')
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert any("no number" in w for w in warnings), warnings


def test_exhibit_heading_with_number_not_warned():
    md = ('## Slide 1: Margins expanded 400bps on pricing discipline\n'
          '**Layout:** waterfall\n'
          '**Data:**\n- FY24: 42\n- Pricing: +9\n- FY25: total\n'
          '- Notes: "n"\n')
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert not any("no number" in w for w in warnings), warnings


def test_heading_with_and_warned():
    md = ('## Slide 1: We doubled revenue and we cut operating costs\n'
          '**Layout:** bullet-list\n- Point 1: "x"\n- Notes: "n"\n')
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert any("joins two messages" in w for w in warnings), warnings


def test_sticker_and_kicker_rendered(tmp_path):
    from pptx import Presentation
    md = ('## Slide 1: Margins expanded 400bps on pricing discipline\n'
          '**Layout:** bullet-list\n'
          '- Sticker: Illustrative\n'
          '- Kicker: "Lock in pricing gains before renewals"\n'
          '- Pricing actions held through renewals\n'
          '- Notes: "n"\n')
    outline = tmp_path / "o.md"
    outline.write_text(md)
    out = tmp_path / "deck.pptx"
    assert build(outline, out)
    prs = Presentation(str(out))
    texts = {sh.text_frame.text for sh in prs.slides[0].shapes
             if sh.has_text_frame}
    assert "ILLUSTRATIVE" in texts, texts
    assert "Lock in pricing gains before renewals" in texts, texts
    # sticker tag sits top-right (no section label -> y ~0.14)
    tag = [sh for sh in prs.slides[0].shapes
           if sh.has_text_frame and sh.text_frame.text == "ILLUSTRATIVE"][0]
    assert abs(tag.top / 914400 - 0.16) < 0.05, tag.top
    # kicker box top ~6.18
    kick = [sh for sh in prs.slides[0].shapes
            if sh.has_text_frame
            and sh.text_frame.text == "Lock in pricing gains before renewals"][0]
    assert 6.1 < kick.top / 914400 < 6.5, kick.top


def test_sticker_drops_below_section_label(tmp_path):
    from pptx import Presentation
    md = ('## Slide 1: Growth\n**Layout:** section-divider\n'
          '## Slide 2: Margins expanded 400bps on pricing discipline\n'
          '**Layout:** bullet-list\n'
          '- Sticker: Preliminary\n'
          '- Pricing actions held through renewals\n'
          '- Notes: "n"\n')
    outline = tmp_path / "o.md"
    outline.write_text(md)
    out = tmp_path / "deck.pptx"
    assert build(outline, out)
    prs = Presentation(str(out))
    tag = [sh for sh in prs.slides[1].shapes
           if sh.has_text_frame and sh.text_frame.text == "PRELIMINARY"][0]
    assert abs(tag.top / 914400 - 0.52) < 0.05, tag.top


def test_dump_numbers_extracts_per_slide(capsys, tmp_path):
    from pptx import Presentation
    from pptx.util import Inches
    from qa_check import dump_numbers
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    tb.text_frame.text = "Revenue grew 12% to $4.2B in FY25"
    p = tmp_path / "n.pptx"
    prs.save(str(p))
    dump_numbers(p)
    out = capsys.readouterr().out
    assert "Slide 1" in out and "12%" in out and "$4.2B" in out and "FY25" in out


def test_num_token_rx_suffix_handling():
    from qa_check import NUM_TOKEN_RX
    assert NUM_TOKEN_RX.findall("spread tightened 200bps") == ["200bps"]
    assert NUM_TOKEN_RX.findall("3 key levers") == ["3"]
