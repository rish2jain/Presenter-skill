"""Tests for chart annotations: waterfall difference bracket, CAGR arrow."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402

import builders  # noqa: E402
from build_deck import parse_outline  # noqa: E402
from builders_consulting import build_waterfall_slide  # noqa: E402
from palettes import get_palette  # noqa: E402

PAL = get_palette("midnight-executive")
CTX = {"outline_dir": Path("."), "assets_dir": Path("assets")}


def _prs():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    builders.set_canvas(prs)
    return prs


def _texts(slide):
    return [sh.text_frame.text for sh in slide.shapes
            if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip()]


def _cagr_labels(slide):
    return [t for t in _texts(slide) if t.strip().startswith("CAGR ")]


BRACKET_MD = """## Slide 1: Three levers bridge run-rate down to target
**Layout:** waterfall
**Data:**
- FY25: 46
- Tiering: -8
- Exit: -6
- Discounts: -4
- FY27: total
- Bracket: FY25, FY27
"""


def test_bracket_directive_parses():
    _, slides = parse_outline(BRACKET_MD)
    assert slides[0]["bracket"] == "FY25, FY27"


def test_bracket_renders_auto_pct_label():
    _, slides = parse_outline(BRACKET_MD)
    slide = build_waterfall_slide(_prs(), slides[0], PAL, CTX)
    texts = _texts(slide)
    # 46 -> 28 is -39%
    assert any("-39%" in t for t in texts), texts


def test_bracket_custom_label():
    md = BRACKET_MD.replace('- Bracket: FY25, FY27',
                            '- Bracket: FY25, FY27, "Run-rate reset"')
    _, slides = parse_outline(md)
    slide = build_waterfall_slide(_prs(), slides[0], PAL, CTX)
    assert any("Run-rate reset" in t for t in _texts(slide))


CAGR_MD = """## Slide 1: Revenue compounds at double digits
**Layout:** two-column-split
**Visual:** chart:bar
- CAGR: on
**Data:**
- 2021: 10
- 2022: 13
- 2023: 17
- 2024: 22
- Strong compounding story
"""


def test_cagr_arrow_label_on_chart_slide():
    _, slides = parse_outline(CAGR_MD)
    assert slides[0]["cagr"] == "on"
    prs = _prs()
    slide = builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    # (22/10)^(1/3)-1 = 30.1%
    assert any("CAGR" in t and "30" in t for t in _texts(slide)), _texts(slide)


AXIS_MD = """## Slide 1: EMEA revenue lags at identical scale
**Layout:** two-column-split
**Visual:** chart:bar
- Axis-Max: 50
**Data:**
- 2023: 17
- 2024: 22
- Same scale as the Americas chart
"""


def test_axis_max_pins_value_axis():
    _, slides = parse_outline(AXIS_MD)
    assert slides[0]["axis_max"] == "50"
    prs = _prs()
    slide = builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    chart = next(sh.chart for sh in slide.shapes if getattr(sh, "has_chart", False))
    assert chart.value_axis.maximum_scale == 50.0


CAGR_AXIS_MD = """## Slide 1: Revenue compounds at double digits on a fixed scale
**Layout:** two-column-split
**Visual:** chart:bar
- CAGR: on
- Axis-Max: 50
**Data:**
- 2021: 10
- 2022: 13
- 2023: 17
- 2024: 22
- Same scale as sibling slides
"""


def test_cagr_arrow_respects_axis_max():
    _, slides = parse_outline(CAGR_AXIS_MD)
    prs = _prs()
    slide = builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    chart = next(sh.chart for sh in slide.shapes if getattr(sh, "has_chart", False))
    assert chart.value_axis.maximum_scale == 50.0
    assert any("CAGR" in t for t in _texts(slide))


def test_cagr_invalid_axis_max_warns(capsys):
    md = CAGR_AXIS_MD.replace("- Axis-Max: 50", "- Axis-Max: fifty")
    _, slides = parse_outline(md)
    prs = _prs()
    builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    err = capsys.readouterr().err
    assert "Axis-Max not numeric" in err and "fifty" in err and "CAGR" in err


def test_cagr_skips_with_fewer_than_two_data_points():
    md = """## Slide 1: One year only
**Layout:** two-column-split
**Visual:** chart:bar
- CAGR: on
**Data:**
- 2024: 22
- Not enough history for compounding
"""
    _, slides = parse_outline(md)
    prs = _prs()
    slide = builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    assert not _cagr_labels(slide), _texts(slide)


def test_cagr_skips_with_non_positive_start_or_end():
    md = CAGR_MD.replace("- 2021: 10", "- 2021: 0")
    _, slides = parse_outline(md)
    prs = _prs()
    slide = builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    assert not _cagr_labels(slide), _texts(slide)

    md_end = CAGR_MD.replace("- 2024: 22", "- 2024: -1")
    _, slides_end = parse_outline(md_end)
    slide_end = builders.LAYOUT_MAP["two-column-split"](_prs(), slides_end[0], PAL, CTX)
    assert not _cagr_labels(slide_end), _texts(slide_end)


def test_cagr_label_on_declining_series():
    md = CAGR_MD.replace("- 2021: 10", "- 2021: 22").replace("- 2024: 22",
                                                             "- 2024: 10")
    _, slides = parse_outline(md)
    prs = _prs()
    slide = builders.LAYOUT_MAP["two-column-split"](prs, slides[0], PAL, CTX)
    # (10/22)^(1/3)-1 = -23.1%
    assert any("CAGR" in t and "-23" in t for t in _texts(slide)), _texts(slide)


def test_add_arrowhead_to_connector_adds_tail_end():
    from pptx.enum.shapes import MSO_CONNECTOR
    from pptx.oxml.ns import qn
    from charts import add_arrowhead_to_connector

    prs = _prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(1), Inches(3), Inches(2))
    assert add_arrowhead_to_connector(conn) is True
    ln = conn.line._get_or_add_ln()
    tail = ln.find(qn("a:tailEnd"))
    assert tail is not None
    assert tail.get("type") == "triangle"


def test_add_arrowhead_to_connector_noop_without_private_api(caplog):
    from types import SimpleNamespace
    from charts import add_arrowhead_to_connector

    fake_line = SimpleNamespace()  # no _get_or_add_ln
    conn = SimpleNamespace(line=fake_line)
    assert add_arrowhead_to_connector(conn, warn=True) is False
    assert any("Skipping connector arrowhead" in r.message for r in caplog.records)
