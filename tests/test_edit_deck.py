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


# ── selection parsing (extract / append --slides) ───────────────────────────

def test_parse_selection_range():
    from edit_deck import _parse_selection
    assert _parse_selection("3-7", 10) == [3, 4, 5, 6, 7]


def test_parse_selection_list_and_combo():
    from edit_deck import _parse_selection
    assert _parse_selection("2,5,9", 10) == [2, 5, 9]
    assert _parse_selection("1,3-5", 10) == [1, 3, 4, 5]
    assert _parse_selection("2,2,3-4,3", 10) == [2, 3, 4]  # deduped


def test_parse_selection_errors():
    from edit_deck import _parse_selection
    with pytest.raises(ValueError, match="empty"):
        _parse_selection("", 5)
    with pytest.raises(ValueError, match="out of range"):
        _parse_selection("6", 5)
    with pytest.raises(ValueError, match="out of range"):
        _parse_selection("0", 5)
    with pytest.raises(ValueError):
        _parse_selection("abc", 5)
    with pytest.raises(ValueError, match="reversed"):
        _parse_selection("5-3", 5)


# ── extract (split) ──────────────────────────────────────────────────────────

def _first_texts(pptx_path):
    """First non-empty text of each slide, in deck order."""
    from pptx import Presentation
    out = []
    for slide in Presentation(str(pptx_path)).slides:
        first = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                first = shape.text_frame.text.strip().splitlines()[0]
                break
        out.append(first)
    return out


@pytest.fixture(scope="module")
def example_deck(tmp_path_factory):
    """The 6-slide example deck built once for extract/append tests."""
    outline = ROOT / "assets" / "example-outline.md"
    pptx = tmp_path_factory.mktemp("example") / "deck.pptx"
    assert build(outline, pptx)
    return pptx


def test_extract_range_end_to_end(example_deck, tmp_path):
    from pptx import Presentation
    from edit_deck import extract
    out = tmp_path / "sub.pptx"
    assert extract(example_deck, "2-4", out)
    prs = Presentation(str(out))
    assert len(prs.slides) == 3
    originals = _first_texts(example_deck)
    assert _first_texts(out) == originals[1:4]


def test_extract_comma_list(example_deck, tmp_path):
    from pptx import Presentation
    from edit_deck import extract
    out = tmp_path / "sub.pptx"
    assert extract(example_deck, "1,6", out)
    prs = Presentation(str(out))
    assert len(prs.slides) == 2
    originals = _first_texts(example_deck)
    assert _first_texts(out) == [originals[0], originals[5]]


def test_extract_out_of_range_errors(example_deck, tmp_path):
    from edit_deck import extract
    assert not extract(example_deck, "5-9", tmp_path / "sub.pptx")
    assert not extract(example_deck, "", tmp_path / "sub.pptx")


def test_extract_output_equals_input_errors(example_deck):
    from edit_deck import extract
    assert not extract(example_deck, "1-2", example_deck)


# ── append (merge) ───────────────────────────────────────────────────────────

# Minimal valid 1x1 red PNG (no Pillow needed to create the fixture).
_PNG_1PX = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
    "53de0000000c4944415408d763f8cfc000000301010018dd8db00000000049"
    "454e44ae426082")


def _simple_src_deck(path, with_image=False):
    """Two-slide python-pptx deck: textboxes + optional picture on slide 2."""
    import io
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    for i, label in enumerate(("Alpha source slide", "Beta source slide")):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = label
        run.font.size = Pt(24)
        if with_image and i == 1:
            slide.shapes.add_picture(io.BytesIO(_PNG_1PX), Inches(1),
                                     Inches(3), Inches(2), Inches(2))
    prs.save(str(path))
    return path


def test_append_end_to_end(example_deck, tmp_path):
    import subprocess
    from pptx import Presentation
    from edit_deck import append_decks
    src = _simple_src_deck(tmp_path / "src.pptx", with_image=True)
    merged = tmp_path / "merged.pptx"
    assert append_decks(example_deck, src, None, merged)

    prs = Presentation(str(merged))
    assert len(prs.slides) == 8
    texts = _first_texts(merged)
    assert texts[6] == "Alpha source slide"
    assert texts[7] == "Beta source slide"
    # image copied along
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    assert any(s.shape_type == MSO_SHAPE_TYPE.PICTURE
               for s in prs.slides[7].shapes)

    # no duplicate part names in the package
    with zipfile.ZipFile(merged) as zf:
        names = zf.namelist()
        assert len(names) == len(set(names))
        # src master/theme imported alongside dst's own (faithful import)
        assert any("slideMaster2.xml" in n and "_rels" not in n for n in names)

    # qa_check opens the merged deck without integrity blowups
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "qa_check.py"), str(merged)],
        capture_output=True, text=True)
    assert proc.returncode in (0, 1), proc.stderr
    assert "Traceback" not in proc.stderr
    assert "error(s)" in proc.stdout


def test_append_subset(example_deck, tmp_path):
    from pptx import Presentation
    from edit_deck import append_decks
    src = _simple_src_deck(tmp_path / "src.pptx")
    merged = tmp_path / "merged.pptx"
    assert append_decks(example_deck, src, "2", merged)
    prs = Presentation(str(merged))
    assert len(prs.slides) == 7
    assert _first_texts(merged)[6] == "Beta source slide"


def test_append_chart_slide(example_deck, tmp_path):
    """Chart-bearing source slide: chart part + embedded xlsx travel along."""
    from pptx import Presentation
    from edit_deck import append_decks
    dst = _simple_src_deck(tmp_path / "dst.pptx")
    merged = tmp_path / "merged.pptx"
    assert append_decks(dst, example_deck, "3", merged)
    prs = Presentation(str(merged))
    assert len(prs.slides) == 3
    copied = prs.slides[2]
    assert any(getattr(s, "has_chart", False) for s in copied.shapes)
    chart = next(s for s in copied.shapes if getattr(s, "has_chart", False)).chart
    assert chart.plots  # chart part loads
    with zipfile.ZipFile(merged) as zf:
        names = zf.namelist()
        assert any(n.startswith("ppt/charts/chart") for n in names)
        assert any(n.startswith("ppt/embeddings/") for n in names)


def test_append_dedupes_shared_layout_and_master(example_deck, tmp_path):
    """Two src slides sharing a master → master/theme copied exactly once."""
    from edit_deck import append_decks
    src = _simple_src_deck(tmp_path / "src.pptx")
    merged = tmp_path / "merged.pptx"
    assert append_decks(example_deck, src, "1-2", merged)
    with zipfile.ZipFile(merged) as zf:
        masters = [n for n in zf.namelist()
                   if n.startswith("ppt/slideMasters/slideMaster")
                   and n.endswith(".xml")]
    assert len(masters) == 2  # dst's own + one imported


def test_append_invalid_selection_errors(example_deck, tmp_path):
    from edit_deck import append_decks
    src = _simple_src_deck(tmp_path / "src.pptx")
    assert not append_decks(example_deck, src, "9", tmp_path / "m.pptx")


def test_append_multi_layout_master_pruned(example_deck, tmp_path):
    """Src slides on two different layouts of one master: both layouts
    copied, master's sldLayoutIdLst pruned to exactly those two, with
    fresh presentation-unique ids."""
    from lxml import etree
    from pptx import Presentation
    from edit_deck import append_decks, NS
    prs = Presentation()
    for li in (0, 5):
        slide = prs.slides.add_slide(prs.slide_layouts[li])
        if slide.shapes.title is not None:
            slide.shapes.title.text = f"Layout {li} slide"
    src = tmp_path / "src.pptx"
    prs.save(str(src))
    merged = tmp_path / "merged.pptx"
    assert append_decks(example_deck, src, None, merged)
    Presentation(str(merged))  # opens cleanly
    with zipfile.ZipFile(merged) as zf:
        master = etree.fromstring(zf.read("ppt/slideMasters/slideMaster2.xml"))
        layout_ids = master.findall("p:sldLayoutIdLst/p:sldLayoutId", NS)
        rels = etree.fromstring(
            zf.read("ppt/slideMasters/_rels/slideMaster2.xml.rels"))
        layout_rels = [r for r in rels
                       if r.get("Type", "").endswith("/slideLayout")]
    assert len(layout_ids) == 2
    assert len(layout_rels) == 2
    ids = [int(e.get("id")) for e in layout_ids]
    assert all(i > 2147483647 for i in ids) and len(set(ids)) == 2
    # ids must not collide with dst's existing master/layout ids
    with zipfile.ZipFile(example_deck) as zf:
        dst_master = etree.fromstring(
            zf.read("ppt/slideMasters/slideMaster1.xml"))
        dst_ids = {int(e.get("id")) for e in
                   dst_master.findall("p:sldLayoutIdLst/p:sldLayoutId", NS)}
    assert not set(ids) & dst_ids
