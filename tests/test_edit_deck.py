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
