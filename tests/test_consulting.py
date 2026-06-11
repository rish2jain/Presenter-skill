"""Tests for the consulting layout builders (builders_consulting.py)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from pptx import Presentation  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

import builders  # noqa: E402
from build_deck import parse_outline, validate  # noqa: E402
from builders_consulting import (build_waterfall_slide, build_harvey_slide,
                                 build_mekko_slide, build_gantt_slide,
                                 _fmt_num)  # noqa: E402
from palettes import get_palette  # noqa: E402

PAL = get_palette("midnight-executive")
CTX = {"outline_dir": Path("."), "assets_dir": Path("assets")}


def _prs():
    from pptx.util import Inches
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    builders.set_canvas(prs)
    return prs


def _texts(slide):
    return [sh.text_frame.text for sh in slide.shapes
            if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip()]


WATERFALL_MD = """## Slide 1: Three levers bridge run-rate down to target
**Layout:** waterfall
**Data:**
- FY25: 46
- Tiering: -8
- Exit: -6
- Discounts: -4
- FY27: total
"""


def test_waterfall_signed_delta_labels():
    _, slides = parse_outline(WATERFALL_MD)
    slide = build_waterfall_slide(_prs(), slides[0], PAL, CTX)
    texts = _texts(slide)
    # deltas are signed, endpoints are plain (the +8 magnitude bug regression)
    for expected in ("-8", "-6", "-4", "46", "28"):
        assert expected in texts, f"{expected!r} missing from {texts}"
    assert "+8" not in texts and "+46" not in texts


def test_waterfall_total_row_parses_as_sentinel():
    _, slides = parse_outline(WATERFALL_MD)
    assert slides[0]["data"][-1] == ("FY27", "total")


def test_fmt_num_signs():
    assert _fmt_num(46) == "46"
    assert _fmt_num(-8, signed=True) == "-8"
    assert _fmt_num(8, signed=True) == "+8"


HARVEY_MD = """## Slide 1: Vendor B leads on weighted criteria
**Layout:** harvey-scorecard
| Criterion | A | B |
|---|---|---|
| Scale | 2 | 4 |
| Support | 0 | 3 |
"""


def test_harvey_pie_angles_in_ooxml_units():
    """Pie adjustments are degrees*0.6 (OOXML 60000ths) — the x100000 regression."""
    _, slides = parse_outline(HARVEY_MD)
    slide = build_harvey_slide(_prs(), slides[0], PAL, CTX)
    pies = [el for el in slide.shapes._spTree.findall(".//" + qn("a:prstGeom"))
            if el.get("prst") == "pie"]
    assert len(pies) == 2  # scores 2 and 3 (0 = empty ring, 4 = full disc)
    gds = {gd.get("name"): gd.get("fmla") for gd in pies[0].findall(".//" + qn("a:gd"))}
    assert gds["adj1"] == "val -5400000"  # -90 degrees
    # score 2 -> end angle 90 degrees = 5400000
    assert gds["adj2"] == "val 5400000"


def test_appendix_marks_slides_and_relaxes_warnings():
    md = """## Slide 1: Main point stated as a full action title
**Layout:** bullet-list
- one
- Notes: "x"

## Appendix
## Slide 2: Backup
**Layout:** table
| a | b |
|---|---|
| 1 | 2 |
"""
    meta, slides = parse_outline(md)
    assert not slides[0].get("_appendix")
    assert slides[1]["_appendix"] is True
    _, warnings = validate(slides, CTX, meta)
    # appendix slides are exempt from notes + action-title warnings
    assert not any("Slide 2" in w and ("notes" in w or "action title" in w)
                   for w in warnings)


def test_quadrant_keys_only_apply_on_matrix_slides():
    md = """## Slide 1: Quarterly revenue grows every quarter this year
**Layout:** two-column-split
**Visual:** chart:bar
**Data:**
- Q1: 4.2
- Q2: 5.1
"""
    _, slides = parse_outline(md)
    assert slides[0]["data"] == [("Q1", 4.2), ("Q2", 5.1)]
    assert "q1" not in slides[0]

    md2 = """## Slide 1: Matrix
**Layout:** matrix-2x2
- Q1: "Quick wins"
- Item: Name="A" X="0.5" Y="0.5"
"""
    _, slides2 = parse_outline(md2)
    assert slides2[0]["q1"] == "Quick wins"


MEKKO_MD = """## Slide 1: Two platforms capture most of segment value
**Layout:** mekko
**Series:** A, B
**Data:**
- Compute: 18, 22
- Storage: 9, 11
"""


def test_mekko_builds_segments_and_widths():
    _, slides = parse_outline(MEKKO_MD)
    slide = build_mekko_slide(_prs(), slides[0], PAL, CTX)
    texts = _texts(slide)
    assert "40" in texts and "20" in texts  # column totals
    # 4 segments + legend chips + accent bar = rect count sanity
    rects = [el for el in slide.shapes._spTree.findall(".//" + qn("a:prstGeom"))
             if el.get("prst") == "rect"]
    assert len(rects) >= 6


GANTT_MD = """## Slide 1: Three workstreams deliver across four quarters
**Layout:** gantt
**Periods:** Q1, Q2, Q3, Q4
- Bar: Row="Infra" Label="Tiering" Start="1" End="2"
- Bar: Row="Platform" Label="Pilot" Start="2" End="4"
- Milestone: Row="Platform" Label="GA" At="4"
"""


def test_gantt_builds_rows_bars_milestone():
    _, slides = parse_outline(GANTT_MD)
    assert len(slides[0]["bars"]) == 2
    assert slides[0]["periods"] == ["Q1", "Q2", "Q3", "Q4"]
    slide = build_gantt_slide(_prs(), slides[0], PAL, CTX)
    texts = _texts(slide)
    for expected in ("Infra", "Platform", "Tiering", "Pilot", "GA", "Q1", "Q4"):
        assert any(expected in t for t in texts), f"{expected!r} missing"
    diamonds = [el for el in slide.shapes._spTree.findall(".//" + qn("a:prstGeom"))
                if el.get("prst") == "diamond"]
    assert len(diamonds) == 1


def test_gantt_validator_rejects_out_of_range_bars():
    md = GANTT_MD.replace('Start="2" End="4"', 'Start="2" End="9"')
    meta, slides = parse_outline(md)
    errors, _ = validate(slides, CTX, meta)
    assert any("outside 1-4" in e for e in errors)


def test_exhibit_numbering_meta():
    md = """**Exhibits:** on

## Slide 1: Costs rise without intervention this year
**Layout:** bullet-list
- one
- Source: "FinOps"
"""
    meta, slides = parse_outline(md)
    assert meta["exhibits"] == "on"
    assert slides[0]["source"] == "FinOps"


def test_action_title_warning_fires_on_topic_labels():
    md = """## Slide 1: Leadership
**Layout:** bullet-list
- one
- Notes: "x"
"""
    meta, slides = parse_outline(md)
    _, warnings = validate(slides, CTX, meta)
    assert any("action title" in w for w in warnings)


BAR_MEKKO_MD = """## Slide 1: EMEA is the margin outlier despite its size
**Layout:** bar-mekko
- Bar: Label="Americas" Size="55" Value="14"
- Bar: Label="EMEA" Size="30" Value="6"
- Bar: Label="APAC" Size="15" Value="11"
- Notes: Width = revenue share, height = EBITDA margin.
"""


def test_bar_mekko_widths_proportional_to_size():
    from builders_consulting import build_bar_mekko_slide
    _, slides = parse_outline(BAR_MEKKO_MD)
    slide = build_bar_mekko_slide(_prs(), slides[0], PAL, CTX)
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    rects = [s for s in slide.shapes
             if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and s.height.inches > 0.5]
    assert len(rects) == 3
    widths = sorted((r.width for r in rects), reverse=True)
    assert widths[0] > widths[1] > widths[2]  # 55 > 30 > 15


def test_bar_mekko_registered_and_validated():
    import builders
    assert "bar-mekko" in builders.LAYOUT_MAP
    _, slides = parse_outline("## Slide 1: Bad mekko\n**Layout:** bar-mekko\n")
    errors, _ = validate(slides, CTX)
    assert any("bar-mekko" in e for e in errors)


BUBBLE_MD = """## Slide 1: Two bets dominate the portfolio by revenue at stake
**Layout:** matrix-2x2
- X-axis: Relative share
- Y-axis: Market growth
- Item: Name="Stars" X="0.8" Y="0.8" Size="40"
- Item: Name="Dogs" X="0.2" Y="0.2" Size="5"
- Notes: Bubble area = revenue.
"""


def test_matrix_bubble_sizes_scale():
    from builders_consulting import build_matrix_slide
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    _, slides = parse_outline(BUBBLE_MD)
    slide = build_matrix_slide(_prs(), slides[0], PAL, CTX)
    ovals = [s for s in slide.shapes
             if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
             and s.width == s.height and s.width.inches > 0.1]
    assert len(ovals) == 2
    big, small = sorted((o.width.inches for o in ovals), reverse=True)
    assert big > small * 1.8, (big, small)  # sqrt(40/5) ≈ 2.8x diameter


def test_matrix_negative_size_falls_back_to_fixed_dot():
    from builders_consulting import build_matrix_slide
    md = BUBBLE_MD.replace('Size="5"', 'Size="-5"')
    _, slides = parse_outline(md)
    slide = build_matrix_slide(_prs(), slides[0], PAL, CTX)  # must not raise
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    ovals = [s for s in slide.shapes
             if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and s.width == s.height]
    assert len(ovals) == 2


def test_bar_mekko_rejects_negative_values_in_validate():
    md = BAR_MEKKO_MD.replace('Size="30"', 'Size="-30"')
    _, slides = parse_outline(md)
    errors, _ = validate(slides, CTX)
    assert any("Size>0" in e for e in errors), errors


def test_matrix_bubble_clamped_inside_plot():
    from builders_consulting import build_matrix_slide
    md = BUBBLE_MD.replace('X="0.8" Y="0.8"', 'X="1.0" Y="1.0"')
    _, slides = parse_outline(md)
    slide = build_matrix_slide(_prs(), slides[0], PAL, CTX)
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    EMU_IN = 914400
    plot_right = (2.0 + 9.6) * EMU_IN  # L + W from the builder
    for s in slide.shapes:
        if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and s.width == s.height \
                and s.width / EMU_IN > 0.1:
            assert s.left + s.width <= plot_right + 100, "bubble escapes plot"
