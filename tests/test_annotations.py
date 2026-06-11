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
