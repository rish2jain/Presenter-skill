"""Tests for edit_deck position-based slide operations."""
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from build_deck import build  # noqa: E402
from edit_deck import _resolve_position, duplicate, remove, list_slides  # noqa: E402


@pytest.fixture
def unpacked_three_slides(tmp_path):
    outline = ROOT / "assets" / "example-outline.md"
    pptx = tmp_path / "deck.pptx"
    assert build(outline, pptx)
    unpacked = tmp_path / "unpacked"
    from edit_deck import unpack

    unpack(pptx, unpacked)
    return unpacked


def test_resolve_position_matches_list_order(unpacked_three_slides):
    src = unpacked_three_slides
    assert _resolve_position(src, 1) == 1
    assert _resolve_position(src, 3) == 3
    assert _resolve_position(src, 99) is None


def test_remove_by_position(unpacked_three_slides):
    src = unpacked_three_slides
    assert remove(src, 1)
    assert _resolve_position(src, 1) == 2  # was slide 2, now first


def test_duplicate_by_position(unpacked_three_slides, tmp_path):
    src = unpacked_three_slides
    assert duplicate(src, 2)
    out = tmp_path / "packed.pptx"
    from edit_deck import pack

    assert pack(src, out)
    with zipfile.ZipFile(out) as zf:
        slide_parts = [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    assert len(slide_parts) == 7  # 6 original + 1 duplicate


def test_inventory_and_replace_roundtrip(tmp_path):
    import json
    from pptx import Presentation
    from pptx.util import Inches
    import edit_deck

    # fixture deck: one slide, one textbox
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "Old headline"
    run.font.bold = True
    src = tmp_path / "deck.pptx"
    prs.save(str(src))

    work = tmp_path / "unpacked"
    edit_deck.unpack(str(src), str(work))
    inv = edit_deck.inventory(str(work))
    assert inv == [{"slide": 1, "run": 0, "text": "Old headline"}]

    edits = tmp_path / "edits.json"
    edits.write_text(json.dumps(
        [{"slide": 1, "run": 0, "text": "New headline"}]))
    edit_deck.replace_runs(str(work), str(edits))
    out = tmp_path / "out.pptx"
    edit_deck.pack(str(work), str(out))

    prs2 = Presentation(str(out))
    runs = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert runs[0].text == "New headline"
    assert runs[0].font.bold  # formatting preserved


def test_replace_out_of_range_run_raises(tmp_path):
    import json
    import pytest
    from pptx import Presentation
    from pptx.util import Inches
    import edit_deck
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text_frame.paragraphs[0].add_run().text = "x"
    src = tmp_path / "deck.pptx"
    prs.save(str(src))
    work = tmp_path / "unpacked"
    edit_deck.unpack(str(src), str(work))
    edits = tmp_path / "edits.json"
    edits.write_text(json.dumps([{"slide": 1, "run": 999, "text": "y"}]))
    with pytest.raises(SystemExit, match="out of range"):
        edit_deck.replace_runs(str(work), str(edits))
