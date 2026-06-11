"""Tests for edit_deck reorder and clean."""
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from build_deck import build  # noqa: E402
from edit_deck import unpack, pack, reorder, clean_orphans, list_slides  # noqa: E402


@pytest.fixture
def unpacked(tmp_path):
    outline = ROOT / "assets" / "example-outline.md"
    pptx = tmp_path / "deck.pptx"
    assert build(outline, pptx)
    unpacked = tmp_path / "unpacked"
    unpack(pptx, unpacked)
    return unpacked


def test_reorder_slides(unpacked, tmp_path):
    assert reorder(unpacked, "6,1,2,3,4,5")
    out = tmp_path / "reordered.pptx"
    assert pack(unpacked, out)
    with zipfile.ZipFile(out) as zf:
        slides = [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    assert len(slides) == 6


def test_clean_orphans(unpacked):
    media = unpacked / "ppt" / "media"
    media.mkdir(parents=True, exist_ok=True)
    orphan = media / "orphan.png"
    orphan.write_bytes(b"fake")
    assert clean_orphans(unpacked)
    assert not orphan.exists()
