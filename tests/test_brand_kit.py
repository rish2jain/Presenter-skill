"""Tests for brand_kit.py — fully offline (network monkeypatched, fixture
HTML/CSS/JSON embedded below; no real requests ever made)."""
import json
import sys
import urllib.error
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import brand_kit  # noqa: E402
import palettes  # noqa: E402
from palettes import REQUIRED_KEYS, load_custom_palettes  # noqa: E402
from qa_check import contrast_ratio  # noqa: E402


# ── Fixtures (inline — no files, no network) ─────────────────────────────────

HOME_URL = "https://acme.test"
CSS_URL = "https://acme.test/static/site.css"
CDN_CSS_URL = "https://cdn.example.net/lib.css"

FIXTURE_HTML = """<!doctype html>
<html><head>
  <link rel="stylesheet" href="/static/site.css">
  <link href="https://cdn.example.net/lib.css" rel="stylesheet">
  <style>body{background:#101820;color:#f4f4f4}</style>
</head><body><div style="color:#fee715">hi</div></body></html>
"""

FIXTURE_CSS = """
body { background: #101820; }
.nav { background: #101820; }
.card { background: #18222c; }
.btn { background: #fee715; color: #101820; }
.link { color: rgb(141, 169, 196); }
.alt { color: #5c946e; }
.tiny { color: #fff; }
.fade { color: rgba(254, 231, 21, 0.5); }
"""

PNG_BYTES = b"\x89PNG\r\n\x1a\n" + b"fake-png-payload"
LOGO_URL = "https://cdn.brandfetch.test/acme/logo.png"
API_URL = "https://api.brandfetch.io/v2/brands/acme.com"

BRANDFETCH_JSON = {
    "name": "Acme",
    "colors": [
        {"hex": "#C9A84C", "type": "accent"},
        {"hex": "#0A0F1E", "type": "dark"},
        {"hex": "#F1F5F9", "type": "light"},
        {"hex": "#06B6D4", "type": "brand"},
        {"hex": "bogus", "type": "brand"},
    ],
    "fonts": [{"name": "Inter", "type": "title"}],
    "logos": [
        {"type": "icon", "formats": [{"src": "https://x.test/i.png",
                                      "format": "png"}]},
        {"type": "logo", "formats": [
            {"src": "https://x.test/logo.svg", "format": "svg"},
            {"src": LOGO_URL, "format": "png"}]},
    ],
}


class FakeHTTP:
    """Serves a dict of url → str/bytes; raises URLError for anything else."""

    def __init__(self, pages):
        self.pages = pages
        self.requested = []
        self.headers = {}

    def __call__(self, url, headers=None, timeout=10):
        self.requested.append(url)
        self.headers[url] = dict(headers or {})
        if url not in self.pages:
            raise urllib.error.URLError(f"no fixture for {url}")
        body = self.pages[url]
        if isinstance(body, str):
            body = body.encode("utf-8")
        return body, url


def _rgb(h):
    return brand_kit._rgb(h)


# ── Pure logic: extraction / dedupe ──────────────────────────────────────────

def test_extract_hex_rgb_and_three_digit():
    text = ("a{color:#0A0F1E} b{color:#abc} c{color:rgb(10, 20, 30)} "
            "d{color:rgba(10,20,30,.5)} e{border:#0a0f1e}")
    counts = brand_kit.extract_colors(text)
    assert counts[(10, 15, 30)] == 2          # #0A0F1E twice, case-insensitive
    assert counts[(170, 187, 204)] == 1       # #abc expanded to #AABBCC
    assert counts[(10, 20, 30)] == 2          # rgb() + rgba()


def test_extract_skips_eight_digit_hex_and_oob_rgb():
    counts = brand_kit.extract_colors("x{a:#11223344} y{b:rgb(300,0,0)}")
    assert not counts


def test_dedupe_merges_near_colors_summing_frequency():
    merged = brand_kit.dedupe([((10, 10, 10), 5), ((12, 12, 12), 4),
                               ((200, 0, 0), 2)])
    assert merged == [((10, 10, 10), 9), ((200, 0, 0), 2)]


# ── Pure logic: palette assembly ─────────────────────────────────────────────

def test_dark_theme_wins_on_frequency():
    pal = brand_kit.assemble_palette([((10, 15, 30), 8), ((250, 250, 248), 3),
                                      ((200, 40, 40), 5)])
    assert pal["dark"] is True
    assert pal["bg"] == "0A0F1E"


def test_light_theme_wins_on_frequency():
    pal = brand_kit.assemble_palette([((250, 250, 248), 10), ((20, 20, 20), 3),
                                      ((200, 30, 40), 5)])
    assert pal["dark"] is False
    assert pal["bg"] == "FAFAF8"


def test_tie_prefers_dark():
    pal = brand_kit.assemble_palette([((10, 10, 10), 5), ((250, 250, 250), 5)])
    assert pal["dark"] is True


def test_accent_fallback_hue_rotation():
    pal = brand_kit.assemble_palette([((10, 12, 20), 9), ((200, 40, 40), 4)])
    accents = {pal["accent1"], pal["accent2"], pal["accent3"]}
    assert len(accents) == 3  # 1 found, 2 derived by ±30° hue rotation
    bg = _rgb(pal["bg"])
    for a in accents:
        assert contrast_ratio(_rgb(a), bg) >= 3.0


def test_contrast_nudging_clears_thresholds():
    # accent (50,50,120) starts well below 3.0:1 on the near-black bg
    pal = brand_kit.assemble_palette([((8, 10, 16), 10), ((50, 50, 120), 5)])
    bg, surface = _rgb(pal["bg"]), _rgb(pal["surface"])
    text = _rgb(pal["text"])
    assert contrast_ratio(text, bg) >= 4.5
    assert contrast_ratio(text, surface) >= 4.5
    for key in ("accent1", "accent2", "accent3"):
        assert contrast_ratio(_rgb(pal[key]), bg) >= 3.0


def test_assemble_emits_all_required_keys():
    pal = brand_kit.assemble_palette([((10, 15, 30), 5), ((200, 40, 40), 2)])
    assert REQUIRED_KEYS <= set(pal)
    assert isinstance(pal["dark"], bool)


# ── End-to-end: CSS scrape source ────────────────────────────────────────────

def test_scrape_writes_validated_palette(tmp_path, monkeypatch):
    monkeypatch.delenv("BRANDFETCH_API_KEY", raising=False)
    fake = FakeHTTP({HOME_URL: FIXTURE_HTML, CSS_URL: FIXTURE_CSS})
    monkeypatch.setattr(brand_kit, "_http_get", fake)
    assert brand_kit.run("acme.test", "bkscrape", tmp_path) == 0
    out = tmp_path / "palettes" / "bkscrape.json"
    assert out.is_file()
    pal = json.loads(out.read_text())
    assert REQUIRED_KEYS <= set(pal)
    assert pal["dark"] is True
    assert pal["bg"] == "101820"
    assert pal["_source"] == HOME_URL
    # cross-origin stylesheet never requested
    assert CDN_CSS_URL not in fake.requested
    # written file loads cleanly through the real loader
    assert "bkscrape" in load_custom_palettes(tmp_path / "palettes")
    assert "bkscrape" in palettes.PALETTES


def test_scrape_survives_css_fetch_failure(tmp_path, monkeypatch):
    monkeypatch.delenv("BRANDFETCH_API_KEY", raising=False)
    fake = FakeHTTP({HOME_URL: FIXTURE_HTML})  # site.css → URLError
    monkeypatch.setattr(brand_kit, "_http_get", fake)
    assert brand_kit.run("acme.test", "bknocss", tmp_path) == 0
    assert (tmp_path / "palettes" / "bknocss.json").is_file()


# ── End-to-end: Brandfetch source ────────────────────────────────────────────

def test_brandfetch_path(tmp_path, monkeypatch):
    monkeypatch.setenv("BRANDFETCH_API_KEY", "testkey")
    fake = FakeHTTP({API_URL: json.dumps(BRANDFETCH_JSON),
                     LOGO_URL: PNG_BYTES})
    monkeypatch.setattr(brand_kit, "_http_get", fake)
    assert brand_kit.run("acme.com", "bkbrand", tmp_path) == 0
    assert fake.headers[API_URL].get("Authorization") == "Bearer testkey"
    pal = json.loads((tmp_path / "palettes" / "bkbrand.json").read_text())
    assert pal["dark"] is True
    assert pal["bg"] == "0A0F1E"          # type "dark" wins bg
    assert pal["accent1"] == "06B6D4"     # type "brand" outranks "accent"
    assert pal["_source"] == "brandfetch"
    logo = tmp_path / "brand" / "bkbrand-logo.png"
    assert logo.is_file() and logo.read_bytes().startswith(b"\x89PNG")


def test_brandfetch_logo_failure_nonfatal(tmp_path, monkeypatch):
    monkeypatch.setenv("BRANDFETCH_API_KEY", "testkey")
    fake = FakeHTTP({API_URL: json.dumps(BRANDFETCH_JSON)})  # logo URL 404s
    monkeypatch.setattr(brand_kit, "_http_get", fake)
    assert brand_kit.run("acme.com", "bknologo", tmp_path) == 0
    assert (tmp_path / "palettes" / "bknologo.json").is_file()
    assert not (tmp_path / "brand" / "bknologo-logo.png").exists()


def test_brandfetch_error_falls_back_to_scrape(tmp_path, monkeypatch):
    monkeypatch.setenv("BRANDFETCH_API_KEY", "testkey")
    fake = FakeHTTP({HOME_URL: FIXTURE_HTML, CSS_URL: FIXTURE_CSS})
    monkeypatch.setattr(brand_kit, "_http_get", fake)
    assert brand_kit.run("acme.test", "bkfall", tmp_path) == 0
    pal = json.loads((tmp_path / "palettes" / "bkfall.json").read_text())
    assert pal["_source"] == HOME_URL


# ── Failure modes ────────────────────────────────────────────────────────────

def test_both_sources_fail_exit_1_no_files(tmp_path, monkeypatch):
    monkeypatch.setenv("BRANDFETCH_API_KEY", "testkey")
    monkeypatch.setattr(brand_kit, "_http_get", FakeHTTP({}))
    assert brand_kit.run("acme.test", "bkfail", tmp_path) == 1
    assert not (tmp_path / "palettes").exists()
    assert not (tmp_path / "brand").exists()


def test_main_exits_1_when_offline(tmp_path, monkeypatch):
    monkeypatch.delenv("BRANDFETCH_API_KEY", raising=False)
    monkeypatch.setattr(brand_kit, "_http_get", FakeHTTP({}))
    monkeypatch.setattr(sys, "argv", ["brand_kit.py", "nosuch.test",
                                      "--name", "bkmain",
                                      "--assets-dir", str(tmp_path)])
    with pytest.raises(SystemExit) as ei:
        brand_kit.main()
    assert ei.value.code == 1


def test_main_rejects_unsafe_name(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, "argv", ["brand_kit.py", "acme.test",
                                      "--name", "../evil",
                                      "--assets-dir", str(tmp_path)])
    with pytest.raises(SystemExit) as ei:
        brand_kit.main()
    assert ei.value.code == 2


# ── Float-space contrast nudging ─────────────────────────────────────────────

def test_lighten_reaches_255_via_float_state():
    """_nudge must escape the integer stall that occurs when each step is
    rounded immediately.  Verify by targeting an impossible ratio that forces
    all 50 float-space steps: the final colour must be brighter than (246,246,246)
    (i.e. the nudge actually moved) even though integer-rounded lighten() alone
    would stall there."""
    start = (246, 246, 246)
    # _nudge with toward_light=True and a very high target runs all 50 steps;
    # the returned colour must be lighter than the integer-stalled value.
    result = brand_kit._nudge(
        start,
        [(start, 21.0)],   # 21:1 is unreachable; forces all steps to run
        toward_light=True,
        role="test_lighten",
        bg_label="bg",
    )
    # Float accumulation must have pushed at least one channel past 246
    assert max(result) > 246, (
        f"nudge stalled at {result}; float accumulation is not working"
    )


def test_repro_gray_bg_text_meets_contrast():
    """assemble_palette([((108,108,108),10), ((200,60,60),4)]) — text on
    surface must reach ≥4.5:1 after float-space nudging."""
    pal = brand_kit.assemble_palette([((108, 108, 108), 10), ((200, 60, 60), 4)])
    bg = _rgb(pal["bg"])
    surface = _rgb(pal["surface"])
    text = _rgb(pal["text"])
    assert contrast_ratio(text, bg) >= 4.5, (
        f"text vs bg = {contrast_ratio(text, bg):.2f}:1")
    assert contrast_ratio(text, surface) >= 4.5, (
        f"text vs surface = {contrast_ratio(text, surface):.2f}:1")


def test_repro_green_bg_text_meets_contrast():
    """assemble_palette with bg (40,123,40) — text on surface must reach
    ≥4.5:1 after float-space nudging."""
    pal = brand_kit.assemble_palette([((40, 123, 40), 10), ((200, 60, 60), 4)])
    bg = _rgb(pal["bg"])
    surface = _rgb(pal["surface"])
    text = _rgb(pal["text"])
    assert contrast_ratio(text, bg) >= 4.5, (
        f"text vs bg = {contrast_ratio(text, bg):.2f}:1")
    assert contrast_ratio(text, surface) >= 4.5, (
        f"text vs surface = {contrast_ratio(text, surface):.2f}:1")


def test_nudge_warn_on_impossible_target(capsys):
    """When the contrast target genuinely cannot be reached in 50 steps,
    a [WARN] line is emitted to stderr and the function still returns (exit 0
    is preserved — the palette is written)."""
    # Force an impossible scenario: text = white (255,255,255), bg = white
    # (ratio = 1.0:1, target 4.5) — there is nowhere to nudge toward_light.
    # We use a very high target ratio so it definitely can't be met.
    impossible_bg = (255, 255, 255)
    result = brand_kit._nudge(
        (254, 254, 254),
        [(impossible_bg, 21.0)],   # 21:1 is beyond maximum possible WCAG ratio
        toward_light=True,
        role="text",
        bg_label="bg",
    )
    captured = capsys.readouterr()
    assert "[WARN] contrast target missed:" in captured.err
    assert "text" in captured.err
    assert isinstance(result, tuple) and len(result) == 3


# ── Brandfetch malformed-shape crash guard ───────────────────────────────────

def test_brandfetch_malformed_colors_list_falls_back_to_scrape(
        tmp_path, monkeypatch):
    """When Brandfetch returns colors as a list of plain strings (AttributeError
    on .get()), the run() pipeline must fall back to CSS scrape without raising."""
    malformed_json = {
        "name": "BadCo",
        "colors": ["#FF0000", "#00FF00"],   # strings, not dicts → AttributeError
        "fonts": [],
        "logos": [],
    }
    monkeypatch.setenv("BRANDFETCH_API_KEY", "testkey")
    fake = FakeHTTP({
        API_URL: json.dumps(malformed_json),
        HOME_URL: FIXTURE_HTML,
        CSS_URL: FIXTURE_CSS,
    })
    monkeypatch.setattr(brand_kit, "_http_get", fake)
    # Must not raise; falls back to scrape and succeeds
    assert brand_kit.run("acme.test", "bkmalformed", tmp_path) == 0
    pal = json.loads((tmp_path / "palettes" / "bkmalformed.json").read_text())
    # Scrape result, not Brandfetch
    assert pal["_source"] == HOME_URL
