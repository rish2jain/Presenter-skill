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


TRACKED_MD = """**Auto-Agenda:** track

## Slide 1: Acme FY27 Strategy
**Layout:** title
- Title: Acme FY27 Strategy
- Subtitle: Board readout

## Slide 2: Diagnosis
**Layout:** section-divider
- Subtitle: Where we are

## Slide 3: Costs have outgrown revenue for three years
- Cost CAGR 12% vs revenue 4%
- Notes: Set up the problem.

## Slide 4: Plan
**Layout:** section-divider
- Subtitle: What we will do

## Slide 5: Three levers close the gap by FY27
- Tiering, exit, discount discipline
- Notes: The plan.
"""


def test_auto_agenda_track_inserts_tracker_slides():
    from build_deck import parse_outline, apply_auto_agenda
    meta, slides = parse_outline(TRACKED_MD)
    out = apply_auto_agenda(meta, slides)
    layouts = [s["layout"] for s in out]
    # title, agenda, divider, tracker, content, divider, tracker, content
    assert layouts == ["title", "agenda", "section-divider", "agenda",
                       "bullet-list", "section-divider", "agenda",
                       "bullet-list"], layouts
    trackers = [s for s in out if s["layout"] == "agenda" and s.get("current")]
    assert [t["current"] for t in trackers] == ["Diagnosis", "Plan"]
    assert out[1]["bullets"] == ["Diagnosis", "Plan"]


def test_auto_agenda_on_inserts_only_overview():
    from build_deck import parse_outline, apply_auto_agenda
    meta, slides = parse_outline(TRACKED_MD.replace("track", "on"))
    out = apply_auto_agenda(meta, slides)
    agendas = [s for s in out if s["layout"] == "agenda"]
    assert len(agendas) == 1 and not agendas[0].get("current")


def test_auto_agenda_off_is_identity():
    from build_deck import parse_outline, apply_auto_agenda
    meta, slides = parse_outline(TRACKED_MD.replace("**Auto-Agenda:** track\n\n", ""))
    assert apply_auto_agenda(meta, slides) == slides


def test_ghost_mode_renders_layout_labels(tmp_path):
    from build_deck import build
    from pptx import Presentation
    md = tmp_path / "o.md"
    md.write_text("""## Slide 1: Three levers bridge run-rate down to target
**Layout:** waterfall
**Data:**
- FY25: 46
- Tiering: -8
- FY27: total
- Notes: n.
""")
    out = tmp_path / "g.pptx"
    assert build(str(md), str(out), ghost=True)
    prs = Presentation(str(out))
    texts = " ".join(sh.text_frame.text for sh in prs.slides[0].shapes
                     if getattr(sh, "has_text_frame", False))
    assert "waterfall" in texts            # layout label shown
    assert "Three levers" in texts          # real action title kept
    assert "46" not in texts                # data NOT rendered


def test_stamp_appears_on_every_slide(tmp_path):
    from build_deck import build
    from pptx import Presentation
    md = tmp_path / "o.md"
    md.write_text("""**Stamp:** DRAFT

## Slide 1: Costs have outgrown revenue for three years
- Cost CAGR 12% vs revenue 4%
- Notes: n.

## Slide 2: Three levers close the gap by FY27
- Tiering, exit, discounts
- Notes: n.
""")
    out = tmp_path / "d.pptx"
    assert build(str(md), str(out))
    prs = Presentation(str(out))
    for slide in prs.slides:
        texts = [sh.text_frame.text for sh in slide.shapes
                 if getattr(sh, "has_text_frame", False)]
        assert "DRAFT" in texts, texts


def test_auto_agenda_excludes_appendix_sections():
    from build_deck import parse_outline, apply_auto_agenda
    md = TRACKED_MD + """
## Appendix

## Slide 6: Backup detail
**Layout:** section-divider
- Subtitle: Backup
"""
    meta, slides = parse_outline(md)
    out = apply_auto_agenda(meta, slides)
    overview = next(s for s in out if s["layout"] == "agenda")
    assert overview["bullets"] == ["Diagnosis", "Plan"], overview["bullets"]


def test_auto_agenda_divider_first_gets_no_overview():
    from build_deck import parse_outline, apply_auto_agenda
    md = TRACKED_MD.split("## Slide 2", 1)[1]
    md = "**Auto-Agenda:** track\n\n## Slide 2" + md
    meta, slides = parse_outline(md)
    out = apply_auto_agenda(meta, slides)
    overviews = [s for s in out if s["layout"] == "agenda" and not s.get("current")]
    assert not overviews, [s.get("heading") for s in out]


def test_heading_attributes_set_layout_and_palette():
    from build_deck import parse_outline
    _, slides = parse_outline(
        "## Slide 1: Margin bridge tells the story {layout=waterfall palette=aurora}\n"
        "**Data:**\n- FY25: 46\n- FY27: total\n")
    assert slides[0]["layout"] == "waterfall"
    assert slides[0]["palette"] == "aurora"
    assert slides[0]["heading"] == "Margin bridge tells the story"


def test_explicit_layout_line_overrides_heading_attr():
    from build_deck import parse_outline
    _, slides = parse_outline(
        "## Slide 1: T {layout=waterfall}\n**Layout:** funnel\n"
        "**Data:**\n- A: 10\n- B: 5\n")
    assert slides[0]["layout"] == "funnel"
