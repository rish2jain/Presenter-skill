"""Tests for visual_regress.py and qa_check.py --integrity."""
import sys
import types
from pathlib import Path

import pytest
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import qa_check  # noqa: E402
import visual_regress  # noqa: E402


# ---------------------------------------------------------------- fixtures

def _png(path, color=(24, 40, 96), rect=None, size=(320, 180)):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", size, color)
    if rect:
        ImageDraw.Draw(img).rectangle(rect, fill=(250, 200, 40))
    img.save(path, "PNG")
    return path


def _dirs(tmp_path, names=("slide-01.png", "slide-02.png")):
    base, cur = tmp_path / "baseline", tmp_path / "current"
    for d in (base, cur):
        for name in names:
            _png(d / name)
    return base, cur


def _deck(tmp_path):
    from pptx import Presentation
    p = tmp_path / "deck.pptx"
    Presentation().save(str(p))
    return p


# ------------------------------------------------------- visual_regress.py

def test_identical_slides_pass(tmp_path, capsys):
    base, cur = _dirs(tmp_path)
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "0 changed" in out


def test_changed_slide_flagged_with_pixel_diff(tmp_path, capsys):
    base, cur = _dirs(tmp_path)
    _png(cur / "slide-02.png", rect=(40, 20, 280, 160))
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 1
    out = capsys.readouterr().out
    assert "CHANGED" in out
    assert "% pixels differ" in out


def test_threshold_override_relaxes_flagging(tmp_path):
    base, cur = _dirs(tmp_path)
    _png(cur / "slide-02.png", rect=(40, 20, 280, 160))
    rc = visual_regress.main([str(base), str(cur), "--threshold", "64"])
    assert rc == 0


def test_update_initializes_missing_baseline(tmp_path, capsys):
    base = tmp_path / "baseline"  # never created
    cur = tmp_path / "current"
    _png(cur / "slide-01.png")
    _png(cur / "slide-02.png")
    rc = visual_regress.main([str(base), str(cur), "--update"])
    assert rc == 0
    assert (base / "slide-01.png").is_file()
    assert (base / "slide-02.png").is_file()
    # subsequent plain run passes
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 0


def test_update_blesses_changes_and_exits_zero(tmp_path):
    base, cur = _dirs(tmp_path)
    _png(cur / "slide-02.png", rect=(40, 20, 280, 160))
    rc = visual_regress.main([str(base), str(cur), "--update"])
    assert rc == 0
    # baseline now matches current
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 0


def test_missing_baseline_without_update_errors(tmp_path, capsys):
    cur = tmp_path / "current"
    _png(cur / "slide-01.png")
    rc = visual_regress.main([str(tmp_path / "baseline"), str(cur)])
    assert rc == 2
    err = capsys.readouterr().err
    assert "--update" in err


def test_empty_current_errors(tmp_path, capsys):
    base, _ = _dirs(tmp_path)
    empty = tmp_path / "empty"
    empty.mkdir()
    rc = visual_regress.main([str(base), str(empty)])
    assert rc == 2


def test_new_slide_is_warning_not_failure(tmp_path, capsys):
    base, cur = _dirs(tmp_path)
    _png(cur / "slide-03.png")
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "new (no baseline)" in out


def test_deleted_slide_is_failure(tmp_path, capsys):
    base, cur = _dirs(tmp_path)
    _png(base / "slide-03.png")
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 1
    out = capsys.readouterr().out
    assert "MISSING" in out


def test_pdftoppm_padding_variants_match(tmp_path):
    """pdftoppm pads page numbers by deck size: slide-1.png vs slide-01.png
    must be treated as the same slide, not as a delete + add."""
    base = tmp_path / "baseline"
    cur = tmp_path / "current"
    _png(base / "slide-1.png")
    _png(cur / "slide-01.png")
    rc = visual_regress.main([str(base), str(cur)])
    assert rc == 0


def test_fallback_hash_same_verdicts(tmp_path, monkeypatch):
    """Forcing ImportError on imagehash → local average-hash path produces
    the same pass/fail verdicts on the same fixtures."""
    monkeypatch.setitem(sys.modules, "imagehash", None)
    base, cur = _dirs(tmp_path)
    assert visual_regress.main([str(base), str(cur)]) == 0
    _png(cur / "slide-02.png", rect=(40, 20, 280, 160))
    assert visual_regress.main([str(base), str(cur)]) == 1


def test_local_ahash_unit():
    a = visual_regress._ahash.__doc__  # exists
    assert a is not None
    assert visual_regress._hamming(0b1010, 0b1010) == 0
    assert visual_regress._hamming(0b1010, 0b0101) == 4


# ------------------------------------------------- qa_check.py --integrity

def test_integrity_unavailable_prints_info_and_exit_unchanged(
        tmp_path, monkeypatch, capsys):
    deck = _deck(tmp_path)
    monkeypatch.setitem(sys.modules, "openxml_audit", None)  # force ImportError
    monkeypatch.setattr(qa_check.shutil, "which", lambda name: None)
    monkeypatch.setattr(sys, "argv", ["qa_check.py", str(deck), "--integrity"])
    with pytest.raises(SystemExit) as exc:
        qa_check.main()
    assert exc.value.code == 0
    out = capsys.readouterr().out
    assert ("[INFO] pip install openxml-audit for OOXML schema validation "
            "(skipped)") in out


def test_no_integrity_flag_no_info_line(tmp_path, monkeypatch, capsys):
    deck = _deck(tmp_path)
    monkeypatch.setattr(sys, "argv", ["qa_check.py", str(deck)])
    with pytest.raises(SystemExit) as exc:
        qa_check.main()
    assert exc.value.code == 0
    assert "openxml-audit" not in capsys.readouterr().out


def test_integrity_module_errors_surface_as_qa_issues(
        tmp_path, monkeypatch, capsys):
    deck = _deck(tmp_path)
    fake = types.ModuleType("openxml_audit")
    fake.validate = lambda path: ["slide1.xml: invalid <p:bogus> element"]
    monkeypatch.setitem(sys.modules, "openxml_audit", fake)
    monkeypatch.setattr(sys, "argv", ["qa_check.py", str(deck), "--integrity"])
    with pytest.raises(SystemExit) as exc:
        qa_check.main()
    assert exc.value.code == 1
    out = capsys.readouterr().out
    assert "[ERROR] integrity: slide1.xml" in out


def test_integrity_module_clean_deck_passes(tmp_path, monkeypatch, capsys):
    deck = _deck(tmp_path)
    fake = types.ModuleType("openxml_audit")
    fake.validate = lambda path: []
    monkeypatch.setitem(sys.modules, "openxml_audit", fake)
    monkeypatch.setattr(sys, "argv", ["qa_check.py", str(deck), "--integrity"])
    with pytest.raises(SystemExit) as exc:
        qa_check.main()
    assert exc.value.code == 0
    assert "integrity:" not in capsys.readouterr().out


def test_integrity_cli_fallback(tmp_path, monkeypatch, capsys):
    deck = _deck(tmp_path)
    monkeypatch.setitem(sys.modules, "openxml_audit", None)
    monkeypatch.setattr(qa_check.shutil, "which",
                        lambda name: "/usr/local/bin/openxml-audit")

    class FakeResult:
        returncode = 1
        stdout = "part /ppt/slides/slide1.xml: dangling relationship"
        stderr = ""

    monkeypatch.setattr(qa_check.subprocess, "run",
                        lambda *a, **k: FakeResult())
    monkeypatch.setattr(sys, "argv", ["qa_check.py", str(deck), "--integrity"])
    with pytest.raises(SystemExit) as exc:
        qa_check.main()
    assert exc.value.code == 1
    out = capsys.readouterr().out
    assert "integrity: part /ppt/slides/slide1.xml" in out


def test_integrity_module_unknown_api_skips_without_failing(
        tmp_path, monkeypatch, capsys):
    deck = _deck(tmp_path)
    fake = types.ModuleType("openxml_audit")  # no validate/audit/check
    monkeypatch.setitem(sys.modules, "openxml_audit", fake)
    monkeypatch.setattr(qa_check.shutil, "which", lambda name: None)
    monkeypatch.setattr(sys, "argv", ["qa_check.py", str(deck), "--integrity"])
    with pytest.raises(SystemExit) as exc:
        qa_check.main()
    assert exc.value.code == 0
