# Presentation Skill Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the enhancements identified in the June 2026 ecosystem research: a deterministic deck-lint pass, font preflight, faster QA rendering, think-cell-style chart annotations, new consulting chart variants, auto agenda/tracker slides, stamps, ghost-deck mode, outline ergonomics (heading attributes, line numbers), custom JSON palettes, edit-mode inventory/replace, and an LLM consistency-check stage.

**Architecture:** All features extend the existing python-pptx pipeline (`scripts/build_deck.py` → `scripts/builders*.py` → QA via `scripts/qa_check.py` + `scripts/render_slides.py`). New cross-slide checks live in a new `scripts/pptx_lint.py`. Chart annotations live in `scripts/charts.py` (native-chart overlays) and `scripts/builders_consulting.py` (shape-drawn waterfall). No new hard dependencies; `unoserver` and `fontconfig` are optional accelerators with fallbacks.

**Tech Stack:** Python 3, python-pptx 1.0.2 (pinned), lxml, Pillow, pytest. Optional: LibreOffice/unoserver, poppler, fontconfig.

**Conventions:** Existing code style (no type annotations in scripts — match the file you touch). Conventional commits, no attribution footer (disabled globally). All work on a feature branch.

**Research provenance (for context, not execution):** lint rules ≈ Macabacus Deck Check / UpSlide Slide Check; annotations ≈ think-cell; labeled grid ≈ anthropics/skills `thumbnail.py`; inventory/replace ≈ tfriedel/claude-office-skills; heading attributes ≈ Quarto/pandoc; JSON palettes ≈ reveal.js token themes.

---

## Phase 0 — Setup

### Task 0: Feature branch

**Files:** none

- [ ] **Step 1: Create branch and verify clean tree**

```bash
cd /Users/rishabh/.claude/skills/presentation-skill
git status && git branch
git checkout -b feature/skill-enhancements
```

Expected: clean tree, new branch `feature/skill-enhancements`.

- [ ] **Step 2: Verify baseline is green**

```bash
python3 scripts/smoke_test.py && python3 -m pytest tests/ -q
```

Expected: smoke test passes, all tests pass. If not, STOP and report — do not build on a red baseline.

---

## Phase 1 — QA & Infrastructure

### Task 1: `pptx_lint.py` — cross-slide consistency lint

Deterministic deck-wide checks that complement `qa_check.py` (which is per-slide): anti-jiggle for recurring elements, page-number sequence, font/color inventory, and palette-whitelist enforcement.

**Files:**
- Create: `scripts/pptx_lint.py`
- Test: `tests/test_lint.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_lint.py`:

```python
"""Tests for pptx_lint.py cross-slide consistency checks."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

from pptx_lint import lint_deck  # noqa: E402


def _prs():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    return prs


def _blank(prs):
    layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    return prs.slides.add_slide(layout)


def _tb(slide, text, left, top, w=0.8, h=0.35, size=11, color=None):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    if color:
        from pptx.dml.color import RGBColor
        run.font.color.rgb = RGBColor.from_string(color)
    return tb


def test_jiggle_flagged_for_misaligned_page_numbers():
    prs = _prs()
    for n, left in enumerate((11.9, 11.9, 11.4)):  # third one jiggles
        _tb(_blank(prs), str(n + 1), left, 7.08)
    issues = lint_deck(prs)
    assert any("jiggle" in e for e in issues["error"]), issues


def test_aligned_page_numbers_pass():
    prs = _prs()
    for n in range(3):
        _tb(_blank(prs), str(n + 1), 11.9, 7.08)
    issues = lint_deck(prs)
    assert not any("jiggle" in e for e in issues["error"]), issues


def test_page_sequence_gap_flagged():
    prs = _prs()
    for label in ("1", "2", "4"):
        _tb(_blank(prs), label, 11.9, 7.08)
    issues = lint_deck(prs)
    assert any("sequence" in e for e in issues["error"]), issues


def test_font_explosion_warned():
    prs = _prs()
    slide = _blank(prs)
    for i, fname in enumerate(
            ("Calibri", "Arial", "Georgia", "Verdana", "Impact", "Tahoma")):
        tb = _tb(slide, f"text {i}", 1.0, 1.0 + i * 0.5)
        tb.text_frame.paragraphs[0].runs[0].font.name = fname
    issues = lint_deck(prs)
    assert any("font" in w.lower() for w in issues["warn"]), issues


def test_palette_whitelist_flags_off_palette_color():
    prs = _prs()
    slide = _blank(prs)
    _tb(slide, "on palette", 1.0, 1.0, color="C9A84C")   # midnight accent1
    _tb(slide, "rogue", 1.0, 2.0, color="FF00FF")
    issues = lint_deck(prs, palette_key="midnight-executive")
    assert any("FF00FF" in e for e in issues["error"]), issues
    assert not any("C9A84C" in e for e in issues["error"]), issues
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_lint.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'pptx_lint'`

- [ ] **Step 3: Implement `scripts/pptx_lint.py`**

```python
#!/usr/bin/env python3
"""
pptx_lint.py — deterministic cross-slide consistency lint for a .pptx.

Complements qa_check.py (per-slide defects) with deck-wide checks in the
Macabacus "Deck Check" / UpSlide "Slide Check" class:

  1. Anti-jiggle: recurring elements (page numbers, footers, kickers) must sit
     at identical coordinates on every slide.
  2. Page-number sequence: consecutive, no gaps or duplicates.
  3. Font inventory: more than MAX_FONTS distinct fonts reads as inconsistent.
  4. Color inventory / palette whitelist: with --palette, any explicit run or
     fill color outside the palette (plus known extras) is an error.

Usage:
    python3 scripts/pptx_lint.py deck.pptx [--palette midnight-executive]
"""
import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from qa_check import iter_shapes, _solid_fill_rgb

EMU_IN = 914400
POS_TOL_IN = 0.03           # jiggle tolerance
FOOTER_ZONE = 0.80          # top > 80% of slide height = footer zone
HEADER_ZONE = 0.12          # top < 12% = header zone (kickers, stamps)
MAX_FONTS = 4
MAX_COLORS = 14
PAGENUM_RX = re.compile(r"^(B·)?\d+$")
# waterfall negative red, pure white/black (cards, shadows) are always allowed
EXTRA_ALLOWED = {"D9655B", "FFFFFF", "000000"}


def _shape_box_in(shape):
    try:
        if None in (shape.left, shape.top):
            return None
        return (shape.left / EMU_IN, shape.top / EMU_IN)
    except (AttributeError, TypeError):
        return None


def _iter_text_shapes(prs):
    for n, slide in enumerate(prs.slides, 1):
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False) \
                    and shape.text_frame.text.strip():
                yield n, slide, shape


def check_jiggle(prs, issues):
    """Recurring header/footer-zone elements must not move between slides."""
    sh_in = prs.slide_height / EMU_IN
    groups = defaultdict(list)  # role -> [(slide_no, left, top)]
    for n, _slide, shape in _iter_text_shapes(prs):
        box = _shape_box_in(shape)
        if box is None:
            continue
        left, top = box
        in_footer = top > FOOTER_ZONE * sh_in
        in_header = top < HEADER_ZONE * sh_in
        if not (in_footer or in_header):
            continue
        text = shape.text_frame.text.strip()
        role = "page-number" if PAGENUM_RX.match(text) else text[:60]
        groups[role].append((n, left, top))
    for role, occ in groups.items():
        if len(occ) < 3:
            continue
        ref_l, ref_t = occ[0][1], occ[0][2]
        for n, left, top in occ[1:]:
            if abs(left - ref_l) > POS_TOL_IN or abs(top - ref_t) > POS_TOL_IN:
                issues["error"].append(
                    f"Slide {n}: '{role}' jiggle — at ({left:.2f},{top:.2f})in "
                    f"but ({ref_l:.2f},{ref_t:.2f})in on slide {occ[0][0]}")


def check_page_sequence(prs, issues):
    """Plain-digit footer-zone page numbers must be strictly consecutive."""
    sh_in = prs.slide_height / EMU_IN
    pages = []  # (slide_no, value)
    for n, _slide, shape in _iter_text_shapes(prs):
        box = _shape_box_in(shape)
        if box is None or box[1] <= FOOTER_ZONE * sh_in:
            continue
        text = shape.text_frame.text.strip()
        if text.isdigit():
            pages.append((n, int(text)))
    for (n1, v1), (n2, v2) in zip(pages, pages[1:]):
        if v2 != v1 + 1:
            issues["error"].append(
                f"Slide {n2}: page-number sequence broken "
                f"({v1} on slide {n1}, then {v2})")


def collect_inventory(prs):
    """(fonts, colors): name/hex -> set of slide numbers using it."""
    fonts, colors = defaultdict(set), defaultdict(set)
    for n, slide, shape in _iter_text_shapes(prs):
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if not run.text.strip():
                    continue
                if run.font.name:
                    fonts[run.font.name].add(n)
                try:
                    rgb = run.font.color.rgb
                except (TypeError, AttributeError):
                    rgb = None
                if rgb is not None:
                    colors[str(rgb)].add(n)
    for n, slide in enumerate(prs.slides, 1):
        for shape in iter_shapes(slide.shapes):
            fill = _solid_fill_rgb(shape)
            if fill is not None:
                colors["%02X%02X%02X" % fill].add(n)
    return fonts, colors


def check_inventory(prs, issues, palette_key=None):
    fonts, colors = collect_inventory(prs)
    if len(fonts) > MAX_FONTS:
        issues["warn"].append(
            f"deck uses {len(fonts)} fonts ({', '.join(sorted(fonts))}) — "
            f"more than {MAX_FONTS} reads as inconsistent")
    if palette_key:
        from palettes import PALETTES
        pal = PALETTES.get(palette_key)
        if pal is None:
            issues["warn"].append(f"unknown palette '{palette_key}' — "
                                  "skipping whitelist check")
            return
        allowed = {v.upper() for v in pal.values()
                   if isinstance(v, str) and re.fullmatch(r"[0-9A-Fa-f]{6}", v)}
        allowed |= {c.upper() for c in pal.get("chart_series", [])}
        allowed |= EXTRA_ALLOWED
        for hexv, slides in sorted(colors.items()):
            if hexv.upper() not in allowed:
                where = ", ".join(str(s) for s in sorted(slides)[:5])
                issues["error"].append(
                    f"off-palette color {hexv} on slide(s) {where} "
                    f"(palette: {palette_key})")
    elif len(colors) > MAX_COLORS:
        issues["warn"].append(
            f"deck uses {len(colors)} distinct colors — more than "
            f"{MAX_COLORS}; pass --palette to enforce a whitelist")


def lint_deck(prs, palette_key=None):
    issues = {"error": [], "warn": []}
    check_jiggle(prs, issues)
    check_page_sequence(prs, issues)
    check_inventory(prs, issues, palette_key)
    return issues


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--palette", default=None,
                        help="Enforce this palette's colors as a whitelist")
    args = parser.parse_args()
    prs = Presentation(args.pptx)
    issues = lint_deck(prs, args.palette)
    for e in issues["error"]:
        print(f"  [ERROR] {e}")
    for w in issues["warn"]:
        print(f"  [WARN]  {w}")
    print(f"\n{Path(args.pptx).name}: {len(issues['error'])} error(s), "
          f"{len(issues['warn'])} warning(s)")
    sys.exit(1 if issues["error"] else 0)


if __name__ == "__main__":
    main()
```

Note: `pal.get("chart_series", [])` is forward-compatible with Task 13; it is an empty list until then.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_lint.py -q`
Expected: 5 passed

- [ ] **Step 5: Run lint against the real smoke deck**

```bash
python3 scripts/build_deck.py assets/example-outline.md --output /tmp/lint-deck.pptx
python3 scripts/pptx_lint.py /tmp/lint-deck.pptx --palette midnight-executive
```

Expected: exit 0. If the skill's own output fails its own lint (e.g. a builder uses a literal hex not in EXTRA_ALLOWED), that is a real finding — fix the builder or add the hex to `EXTRA_ALLOWED` with a comment saying which builder uses it.

- [ ] **Step 6: Commit**

```bash
git add scripts/pptx_lint.py tests/test_lint.py
git commit -m "feat: pptx_lint deck-wide consistency checks (jiggle, sequence, palette whitelist)"
```

### Task 2: Font preflight (silent-substitution guard)

LibreOffice silently substitutes missing fonts, which poisons visual QA. Check deck/palette fonts against installed fonts before rendering.

**Files:**
- Modify: `scripts/pptx_lint.py` (add `check_fonts_installed` + wire into `lint_deck`)
- Test: `tests/test_lint.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_lint.py`:

```python
def test_missing_font_warned(monkeypatch):
    import pptx_lint
    monkeypatch.setattr(pptx_lint, "installed_fonts",
                        lambda: {"calibri", "arial"})
    prs = _prs()
    slide = _blank(prs)
    tb = _tb(slide, "hello", 1.0, 1.0)
    tb.text_frame.paragraphs[0].runs[0].font.name = "Gill Sans MT"
    issues = pptx_lint.lint_deck(prs)
    assert any("Gill Sans MT" in w for w in issues["warn"]), issues


def test_font_check_skipped_when_inventory_unavailable(monkeypatch):
    import pptx_lint
    monkeypatch.setattr(pptx_lint, "installed_fonts", lambda: None)
    prs = _prs()
    _tb(_blank(prs), "hello", 1.0, 1.0)
    issues = pptx_lint.lint_deck(prs)
    assert not any("not installed" in w for w in issues["warn"]), issues
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_lint.py -q`
Expected: 2 new tests FAIL with `AttributeError: ... has no attribute 'installed_fonts'`

- [ ] **Step 3: Implement**

Add to `scripts/pptx_lint.py` (below `collect_inventory`):

```python
def installed_fonts():
    """Lowercased family names visible to the renderer, or None if unknowable.

    fc-list is the most truthful source for what LibreOffice will see;
    matplotlib's font_manager is the cross-platform fallback.
    """
    import shutil
    import subprocess
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(["fc-list", ":", "family"],
                                 capture_output=True, text=True, timeout=20)
            if out.returncode == 0:
                names = set()
                for line in out.stdout.splitlines():
                    for fam in line.split(","):
                        names.add(fam.strip().lower())
                if names:
                    return names
        except (OSError, subprocess.TimeoutExpired):
            pass
    try:
        from matplotlib import font_manager
        return {Path(f).stem.split("-")[0].lower()
                for f in font_manager.findSystemFonts()} or None
    except ImportError:
        return None


def check_fonts_installed(prs, issues):
    installed = installed_fonts()
    if installed is None:
        issues["warn"].append(
            "cannot enumerate installed fonts (no fc-list or matplotlib) — "
            "visual QA may silently substitute fonts")
        return
    fonts, _ = collect_inventory(prs)
    for name, slides in sorted(fonts.items()):
        if name.lower() not in installed:
            where = ", ".join(str(s) for s in sorted(slides)[:5])
            issues["warn"].append(
                f"font '{name}' (slide(s) {where}) not installed — renderer "
                "will substitute; QA thumbnails won't match PowerPoint")
```

In `lint_deck`, add `check_fonts_installed(prs, issues)` after `check_inventory(...)`.

Note for `test_font_check_skipped_when_inventory_unavailable`: when `installed_fonts()` returns `None` the only message added contains "cannot enumerate", not "not installed" — the test asserts exactly that distinction.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_lint.py -q`
Expected: all pass

- [ ] **Step 5: Commit**

```bash
git add scripts/pptx_lint.py tests/test_lint.py
git commit -m "feat: font preflight in pptx_lint (warn on fonts the renderer will substitute)"
```

### Task 3: unoserver fast path in `render_slides.py`

`unoconvert` (from unoserver, if the user runs its daemon) converts without a cold LibreOffice launch per call — 50–75% less CPU in the render→fix→re-render loop.

**Files:**
- Modify: `scripts/render_slides.py:22-60` (extract PDF conversion, add unoconvert path)

- [ ] **Step 1: Refactor PDF conversion into `_pptx_to_pdf` with unoconvert fast path**

In `scripts/render_slides.py`, replace the body of `render_with_libreoffice` up to the `pdf_path = pdf_files[0]` line. New helper above it:

```python
def _pptx_to_pdf(pptx_path, tmp):
    """Convert to PDF in tmp. Prefers a running unoserver (unoconvert) —
    no cold LibreOffice launch per call — falling back to soffice."""
    if shutil.which("unoconvert"):
        pdf = Path(tmp) / (Path(pptx_path).stem + ".pdf")
        result = subprocess.run(
            ["unoconvert", "--convert-to", "pdf", str(pptx_path), str(pdf)],
            capture_output=True, text=True, timeout=180)
        if result.returncode == 0 and pdf.is_file():
            return pdf
        print(f"[WARN] unoconvert failed ({result.stderr.strip()[:120]}) — "
              "falling back to soffice", file=sys.stderr)
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf",
         "--outdir", tmp, str(pptx_path)],
        capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice failed: {result.stderr}")
    pdf_files = list(Path(tmp).glob("*.pdf"))
    if not pdf_files:
        raise RuntimeError("No PDF output from LibreOffice")
    return pdf_files[0]
```

And in `render_with_libreoffice`, the `with tempfile.TemporaryDirectory() as tmp:` block becomes:

```python
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = _pptx_to_pdf(pptx_path, tmp)

        # Convert PDF pages to PNG — per page when filtering, so 1,3,5
        # renders exactly those slides (not the 1-5 range)
        page_runs = ([(p, p) for p in sorted(slide_filter)]
                     if slide_filter else [(None, None)])
        ...  # unchanged from here
```

Also update the `main()` renderer-availability condition so unoconvert alone is enough:

```python
    if not args.fallback and shutil.which("pdftoppm") and (
            shutil.which("soffice") or shutil.which("unoconvert")):
```

- [ ] **Step 2: Verify renders still work end-to-end**

```bash
python3 scripts/build_deck.py assets/example-outline.md --output /tmp/render-deck.pptx
python3 scripts/render_slides.py /tmp/render-deck.pptx --out /tmp/qa-thumbs/
```

Expected: `Rendered N slides (LibreOffice)` (the label is unchanged; the fast path is internal) and PNGs in `/tmp/qa-thumbs/`. If neither soffice nor unoconvert is installed, expected: Pillow fallback warning — that still counts as pass for this step.

- [ ] **Step 3: Document in SKILL.md dependencies**

In `SKILL.md`, in the Dependencies code block, after the LibreOffice line add:

```bash
# Optional: pip install unoserver, then run `unoserver` in the background —
# render_slides.py auto-uses it (much faster repeated QA renders)
```

- [ ] **Step 4: Commit**

```bash
git add scripts/render_slides.py SKILL.md
git commit -m "perf: unoserver fast path for QA rendering with soffice fallback"
```

### Task 4: Labeled thumbnail grid

Number each cell in the QA grid so the fresh-eyes reviewer can map visuals back to slides (anthropics/skills `thumbnail.py` pattern).

**Files:**
- Modify: `scripts/render_slides.py:115-135` (`create_thumbnail_grid`)
- Test: `tests/test_features.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_features.py` (it already sets up `sys.path` to `scripts/` at the top — reuse that):

```python
def test_thumbnail_grid_is_labeled(tmp_path):
    from PIL import Image
    from render_slides import create_thumbnail_grid
    imgs = []
    for i in range(3):
        p = tmp_path / f"slide-{i + 1:02d}.png"
        Image.new("RGB", (400, 225), (40, 40, 60)).save(p)
        imgs.append(p)
    out = tmp_path / "grid.png"
    create_thumbnail_grid(imgs, out, cols=2)
    grid = Image.open(out)
    # badge: each cell's top-left corner carries an opaque label strip
    corner = grid.getpixel((4, 4))
    assert corner != (40, 40, 60), "no label badge drawn on cell 1"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_features.py::test_thumbnail_grid_is_labeled -q`
Expected: FAIL (corner pixel equals the slide background — no badge)

- [ ] **Step 3: Implement**

In `create_thumbnail_grid`, replace the paste loop with:

```python
    from PIL import ImageDraw, ImageFont
    try:
        font = ImageFont.load_default(size=max(14, th // 12))
    except TypeError:  # Pillow < 10 has no size kwarg
        font = ImageFont.load_default()

    for idx, img_path in enumerate(imgs):
        img = Image.open(img_path).resize((tw, th), Image.LANCZOS)
        draw = ImageDraw.Draw(img)
        label = str(idx + 1)
        pad = 6
        bbox = draw.textbbox((0, 0), label, font=font)
        bw, bh = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.rectangle((0, 0, bw + 2 * pad, bh + 2 * pad), fill=(0, 0, 0))
        draw.text((pad, pad), label, fill=(255, 255, 255), font=font)
        row, col = divmod(idx, cols)
        grid.paste(img, (col * (tw + 4), row * (th + 4)))
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_features.py -q`
Expected: all pass

- [ ] **Step 5: Commit**

```bash
git add scripts/render_slides.py tests/test_features.py
git commit -m "feat: slide-number badges on QA thumbnail grid"
```

---

## Phase 2 — Chart annotations & new variants

### Task 5: Difference bracket on waterfall + CAGR arrow on native charts

think-cell's two signature annotations. The waterfall bracket uses exact shape geometry; the chart CAGR arrow uses the same plot-box approximation as the existing `add_benchmark_line` (`scripts/charts.py:138`).

**Files:**
- Modify: `scripts/build_deck.py:34-46` (FIELD_KEYS: add `Bracket`, `CAGR`)
- Modify: `scripts/builders_consulting.py:20-72` (`build_waterfall_slide`)
- Modify: `scripts/charts.py` (add `add_cagr_arrow`)
- Modify: `scripts/builders.py:159-197` (`_place_visual` wiring)
- Test: `tests/test_annotations.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_annotations.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_annotations.py -q`
Expected: FAIL — `slides[0]["bracket"]` KeyError (directive not parsed)

- [ ] **Step 3: Parse the directives**

In `scripts/build_deck.py`, add to `FIELD_KEYS`:

```python
    "Bracket": "bracket", "CAGR": "cagr",
```

(Place them on a new line after `"Recommendation": "recommendation",`.)

- [ ] **Step 4: Implement the waterfall bracket**

In `scripts/builders_consulting.py`, append after `_fmt_num`:

```python
def _bracket_on_waterfall(slide, pal, p, bars, bottom, scale, bar_w, gap):
    """Difference bracket between two named bars: ┌────┐ + centered label."""
    spec = p.get("bracket", "")
    parts = [s.strip().strip('"') for s in spec.split(",") if s.strip()]
    if len(parts) < 2:
        warn(f"Bracket needs two bar labels, got: {spec!r}")
        return
    labels = [b[0].lower() for b in bars]
    try:
        ia, ib = labels.index(parts[0].lower()), labels.index(parts[1].lower())
    except ValueError:
        warn(f"Bracket labels not found among bars: {spec!r}")
        return
    (la, _, ha, _, va), (lb, _, hb, _, vb) = bars[ia], bars[ib]
    xa = 0.7 + ia * (bar_w + gap) + bar_w / 2
    xb = 0.7 + ib * (bar_w + gap) + bar_w / 2
    y = max(min(bottom - ha * scale, bottom - hb * scale) - 0.62, 1.55)
    label = parts[2] if len(parts) > 2 else (
        f"{(vb - va) / va:+.0%}" if va else _fmt_num(vb - va, signed=True))
    B.add_rect(slide, xa, y, xb - xa, 0.018, pal["text_muted"])      # beam
    B.add_rect(slide, xa, y, 0.018, 0.14, pal["text_muted"])          # left tick
    B.add_rect(slide, xb - 0.018, y, 0.018, 0.14, pal["text_muted"])  # right tick
    B.add_tb(slide, label, xa, y - 0.36, xb - xa, 0.32, size=13, bold=True,
             color=pal["text"], align=PP_ALIGN.CENTER, font=pal["font_body"])
```

In `build_waterfall_slide`, before the final `B.add_rect(slide, 0.7, bottom, ...)` baseline line, add:

```python
    if p.get("bracket"):
        _bracket_on_waterfall(slide, pal, p, bars, bottom, scale, bar_w, gap)
```

- [ ] **Step 5: Implement the CAGR arrow for native charts**

Append to `scripts/charts.py`:

```python
def _add_arrowhead(conn):
    ln = conn.line._get_or_add_ln()
    tail = etree.SubElement(ln, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")


def add_cagr_arrow(slide, chart, pal, left, top, w, h):
    """CAGR arrow from the first to the last column of a single-series chart.

    Uses the same plot-box approximation as add_benchmark_line — confirm
    placement in visual QA.
    """
    values = [v for v in chart.plots[0].series[0].values if v is not None]
    if len(values) < 2 or values[0] <= 0 or values[-1] <= 0:
        return
    periods = len(values) - 1
    cagr = (values[-1] / values[0]) ** (1 / periods) - 1

    axis_max = _nice_ceil(max(values) * 1.05)
    va = chart.value_axis
    va.maximum_scale = float(axis_max)
    va.minimum_scale = 0.0

    px, pw = left + 0.09 * w, w * 0.88
    py, ph = top + 0.04 * h, h * 0.80
    n = len(values)
    x0 = px + pw * (0.5 / n)
    x1 = px + pw * ((n - 0.5) / n)
    y0 = py + (1 - values[0] / axis_max) * ph - 0.22
    y1 = py + (1 - values[-1] / axis_max) * ph - 0.22

    from pptx.enum.shapes import MSO_CONNECTOR
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    conn.line.color.rgb = hex_rgb(pal["accent1"])
    conn.line.width = Pt(2.0)
    _add_arrowhead(conn)

    from pptx.enum.text import PP_ALIGN
    tb = slide.shapes.add_textbox(
        Inches((x0 + x1) / 2 - 1.1), Inches(min(y0, y1) - 0.34),
        Inches(2.2), Inches(0.3))
    run = tb.text_frame.paragraphs[0].add_run()
    tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    run.text = f"CAGR {cagr:+.1%}"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = hex_rgb(pal["text"])
    run.font.name = pal["font_body"]
```

- [ ] **Step 6: Wire into `_place_visual`**

In `scripts/builders.py` `_place_visual`, in the single-series branch directly after the existing `add_benchmark_line` block (which ends at the `add_benchmark_line(...)` call), add:

```python
            if p.get("cagr") and value in ("bar", "column", "line"):
                from charts import add_cagr_arrow
                add_cagr_arrow(slide, chart, pal, *_sc(left, top, w, h))
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python3 -m pytest tests/test_annotations.py tests/test_consulting.py -q`
Expected: all pass (consulting tests guard against waterfall regressions)

- [ ] **Step 8: Commit**

```bash
git add scripts/build_deck.py scripts/builders_consulting.py scripts/charts.py scripts/builders.py tests/test_annotations.py
git commit -m "feat: waterfall difference bracket and chart CAGR arrow annotations"
```

### Task 6: `Axis-Max` directive (same-scale chart groups)

Honest cross-slide comparison requires identical axis scales. A per-slide `- Axis-Max: 50` pins the value axis; using the same value on sibling slides gives think-cell's "same scale" behavior.

**Files:**
- Modify: `scripts/build_deck.py:34-46` (FIELD_KEYS)
- Modify: `scripts/builders.py:159-197` (`_place_visual`)
- Test: `tests/test_annotations.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_annotations.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_annotations.py::test_axis_max_pins_value_axis -q`
Expected: FAIL with KeyError `'axis_max'`

- [ ] **Step 3: Implement**

In `scripts/build_deck.py` `FIELD_KEYS`, add `"Axis-Max": "axis_max",` next to the Task 5 additions.

In `scripts/builders.py` `_place_visual`: capture the chart in the multi-series branch too (`chart = add_native_chart(...)` instead of bare call), then after both branches (still inside `if kind == "chart" and p.get("data"):`, before the `return`):

```python
        if p.get("axis_max"):
            try:
                chart.value_axis.maximum_scale = float(p["axis_max"])
                chart.value_axis.minimum_scale = 0.0
            except (ValueError, TypeError):
                warn(f"Axis-Max not numeric: {p['axis_max']!r}")
```

(`Axis-Max` is applied last, so it wins over the max set by `benchmark`/`cagr` — document that in Task 17.)

- [ ] **Step 4: Run tests, commit**

Run: `python3 -m pytest tests/test_annotations.py -q` → all pass

```bash
git add scripts/build_deck.py scripts/builders.py tests/test_annotations.py
git commit -m "feat: Axis-Max directive pins chart value axis for same-scale comparisons"
```

### Task 7: `bar-mekko` layout

Variable-width bars: width ∝ Size (e.g. segment revenue), height = Value (e.g. margin %) — the profit-pool chart. Shape-based like the existing mekko.

**Files:**
- Modify: `scripts/builders_consulting.py` (new builder + LAYOUTS registration)
- Modify: `scripts/build_deck.py` (validation + ACTION_TITLE_LAYOUTS)
- Test: `tests/test_consulting.py` (append)

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_consulting.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_consulting.py -q`
Expected: FAIL with ImportError (`build_bar_mekko_slide`)

- [ ] **Step 3: Implement the builder**

In `scripts/builders_consulting.py`, append after `build_gantt_slide` (before the LAYOUTS dict at the bottom of the file):

```python
# ── bar mekko (profit pool: width = size, height = value) ───────────────────
def build_bar_mekko_slide(prs, p, pal, ctx):
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    rows = []
    for b in p.get("bars", []):
        try:
            rows.append((b.get("label", "?"), float(b["size"]), float(b["value"])))
        except (KeyError, ValueError):
            warn(f"bar-mekko row needs numeric Size and Value: {b}")
    if len(rows) < 2:
        warn("bar-mekko has <2 valid bars; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)

    L, bottom, total_w, top = 0.7, 6.15, 11.9, 1.95
    gap = 0.08
    total_size = sum(s for _, s, _ in rows) or 1
    vmax = max(v for _, _, v in rows) or 1
    scale = (bottom - top) / vmax
    accents = [pal["accent1"], pal["accent2"], pal["accent3"]]

    x = L
    for i, (label, size, value) in enumerate(rows):
        w = (total_w - gap * (len(rows) - 1)) * size / total_size
        h = max(value * scale, 0.04)
        B.add_rect(slide, x, bottom - h, w, h, accents[i % len(accents)])
        B.add_tb(slide, _fmt_num(value), x - 0.1, bottom - h - 0.34, w + 0.2,
                 0.3, size=12, bold=True, color=pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        B.add_tb(slide, f"{label}\n{_fmt_num(size)}", x - 0.1, bottom + 0.10,
                 w + 0.2, 0.75, size=11, color=pal["text_muted"],
                 align=PP_ALIGN.CENTER, font=pal["font_label"])
        x += w + gap
    B.add_rect(slide, L, bottom, total_w, 0.02, pal["surface"])
    return slide
```

Find the `LAYOUTS = {` dict at the bottom of `builders_consulting.py` and add:

```python
    "bar-mekko": build_bar_mekko_slide,
```

- [ ] **Step 4: Add validation**

In `scripts/build_deck.py`:
- Add `"bar-mekko"` to `ACTION_TITLE_LAYOUTS` (line ~227).
- In `validate()`, after the `gantt` block, add:

```python
        if layout == "bar-mekko":
            ok_bars = [b for b in p.get("bars", [])
                       if "size" in b and "value" in b]
            if len(ok_bars) < 2:
                errors.append(f"{where}: bar-mekko needs 2+ '- Bar: "
                              'Label=".." Size=".." Value=".."\' rows')
```

- [ ] **Step 5: Run tests, commit**

Run: `python3 -m pytest tests/test_consulting.py -q` → all pass

```bash
git add scripts/builders_consulting.py scripts/build_deck.py tests/test_consulting.py
git commit -m "feat: bar-mekko layout (profit-pool chart: width=size, height=value)"
```

### Task 8: Bubble sizing on `matrix-2x2` (BCG growth–share style)

`- Item: Name=".." X=".." Y=".." Size="40"` scales the dot area; the matrix becomes a bubble chart.

**Files:**
- Modify: `scripts/builders_consulting.py:118-132` (item loop in `build_matrix_slide`)
- Test: `tests/test_consulting.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_consulting.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_consulting.py::test_matrix_bubble_sizes_scale -q`
Expected: FAIL — both ovals are the fixed 0.18" dot

- [ ] **Step 3: Implement**

In `build_matrix_slide` (`scripts/builders_consulting.py`), replace the item loop with:

```python
    items = p.get("matrix_items", [])[:12]
    sizes = []
    for it in items:
        try:
            sizes.append(float(it["size"]))
        except (KeyError, ValueError):
            sizes.append(None)
    smax = max((s for s in sizes if s), default=0)
    for item, s in zip(items, sizes):
        try:
            fx, fy = float(item.get("x", 0.5)), float(item.get("y", 0.5))
        except ValueError:
            continue
        cx = L + fx * W
        cy = T + (1 - fy) * H
        d = 0.18 if not (s and smax) else 0.24 + 0.66 * (s / smax) ** 0.5
        dot = B.add_circle(slide, cx - d / 2, cy - d / 2, d, pal["accent1"])
        if s and smax:
            from helpers import set_fill_alpha
            set_fill_alpha(dot, 80)  # bubbles overlap; keep grid visible
        off = d / 2 + 0.06
        if cx + off + 2.2 > L + W:  # label would cross the right border
            B.add_tb(slide, item.get("name", ""), cx - off - 2.2, cy - 0.16,
                     2.2, 0.35, size=11, color=pal["text"],
                     align=PP_ALIGN.RIGHT, font=pal["font_body"])
        else:
            B.add_tb(slide, item.get("name", ""), cx + off, cy - 0.16, 2.2,
                     0.35, size=11, color=pal["text"], font=pal["font_body"])
```

- [ ] **Step 4: Run tests to verify nothing regressed**

Run: `python3 -m pytest tests/test_consulting.py -q`
Expected: all pass (matrix without Size renders identically: fixed 0.18 dot, same label offsets within 0.01")

- [ ] **Step 5: Commit**

```bash
git add scripts/builders_consulting.py tests/test_consulting.py
git commit -m "feat: optional Size= bubbles on matrix-2x2 (BCG growth-share)"
```

---

## Phase 3 — Deck automation

### Task 9: Auto agenda/tracker slides

`**Auto-Agenda:** on` inserts an agenda slide after the title listing all sections. `**Auto-Agenda:** track` additionally inserts a current-section-highlighted agenda right after each section divider (think-cell chapter-tracker pattern). Sections are derived from `section-divider` slides.

**Files:**
- Modify: `scripts/build_deck.py` (META_KEYS, new `apply_auto_agenda`, call in `build`)
- Test: `tests/test_build_deck.py` (append)

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_build_deck.py` (reuse its existing imports/path setup):

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_build_deck.py -q`
Expected: FAIL with ImportError (`apply_auto_agenda`)

- [ ] **Step 3: Implement**

In `scripts/build_deck.py`:

Add to `META_KEYS`: `"Auto-Agenda": "auto_agenda",`

Add after `parse_outline` (module level):

```python
def _agenda_slide(sections, current=None):
    slide = {"bullets": list(sections), "stats": [], "cards": [], "items": [],
             "data": [], "table_rows": [], "steps": [], "tiles": [],
             "matrix_items": [], "bars": [], "milestones": [],
             "left_bullets": [], "right_bullets": [],
             "layout": "agenda", "heading": "Agenda",
             "notes": "Auto-generated agenda tracker.", "_auto": True}
    if current:
        slide["current"] = current
        slide["heading"] = "Where we are"
    return slide


def apply_auto_agenda(meta, slides):
    """**Auto-Agenda:** on  -> overview agenda after the title slide.
                       track -> + current-highlighted agenda after each divider.
    Sections come from section-divider headings; no dividers -> no-op."""
    mode = meta.get("auto_agenda", "").lower()
    if mode not in ("on", "track"):
        return slides
    sections = [s.get("heading") or s.get("title", "")
                for s in slides if s.get("layout") == "section-divider"]
    if not sections:
        return slides
    out = []
    for i, s in enumerate(slides):
        out.append(s)
        if i == 0 and not s.get("_appendix"):
            out.append(_agenda_slide(sections))
        if mode == "track" and s.get("layout") == "section-divider" \
                and not s.get("_appendix"):
            out.append(_agenda_slide(
                sections, current=s.get("heading") or s.get("title", "")))
    return out
```

In `build()`, directly after `meta, slides_data = parse_outline(...)`:

```python
    slides_data = apply_auto_agenda(meta, slides_data)
```

(It runs before `validate`, so inserted agenda slides are validated like any other; they carry `notes`, so no missing-notes warnings. Footer page numbers stay sequential because numbering happens at build time over the final slide list.)

- [ ] **Step 4: Run tests, commit**

Run: `python3 -m pytest tests/test_build_deck.py -q` → all pass

```bash
git add scripts/build_deck.py tests/test_build_deck.py
git commit -m "feat: Auto-Agenda meta inserts overview and per-section tracker slides"
```

### Task 10: `Stamp` meta (DRAFT / CONFIDENTIAL tag)

`**Stamp:** DRAFT` renders a small bordered tag at the top-left of every slide (top-right is taken by the section kicker).

**Files:**
- Modify: `scripts/build_deck.py` (META_KEYS, `_add_stamp`, build loop)
- Test: `tests/test_build_deck.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_build_deck.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_build_deck.py::test_stamp_appears_on_every_slide -q`
Expected: FAIL — "DRAFT" not in texts

- [ ] **Step 3: Implement**

In `scripts/build_deck.py`:

Add to `META_KEYS`: `"Stamp": "stamp",`

Add after `_add_source`:

```python
def _add_stamp(slide, pal, text):
    """Bordered status tag (DRAFT, CONFIDENTIAL) top-left, above the heading."""
    import builders
    from pptx.enum.text import PP_ALIGN
    box = builders.add_rect(slide, 0.7, 0.14, 1.5, 0.32, pal["bg"],
                            line_hex=pal["accent3"], line_pt=1.0)
    box.fill.background()  # outline only — works on photos and gradients
    builders.add_tb(slide, text.upper(), 0.7, 0.16, 1.5, 0.28, size=11,
                    bold=True, color=pal["text_muted"], font=pal["font_label"],
                    align=PP_ALIGN.CENTER)
```

In the `build()` slide loop, after the `if not templated and p.get("source"):` block:

```python
            if not templated and meta.get("stamp"):
                _add_stamp(slide, pal, meta["stamp"])
```

(Stamps go on every slide including title/closing — that's the point of a DRAFT stamp.)

- [ ] **Step 4: Run tests, commit**

Run: `python3 -m pytest tests/test_build_deck.py -q` → all pass

```bash
git add scripts/build_deck.py tests/test_build_deck.py
git commit -m "feat: Stamp meta renders DRAFT/CONFIDENTIAL tag on every slide"
```

### Task 11: Ghost-deck mode (`--ghost`)

Skeleton deck for stakeholder alignment: real action titles + grey placeholder boxes labeled with layout/visual intent + dimmed planned content. Reuses all the build/QA plumbing.

**Files:**
- Modify: `scripts/builders.py` (add `build_ghost_slide`)
- Modify: `scripts/build_deck.py` (`--ghost` flag, build loop)
- Test: `tests/test_build_deck.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_build_deck.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_build_deck.py::test_ghost_mode_renders_layout_labels -q`
Expected: FAIL with TypeError (`build()` has no `ghost` parameter)

- [ ] **Step 3: Implement the ghost builder**

In `scripts/builders.py`, add after `_visual_placeholder`:

```python
GHOST_KEEP_REAL = {"title", "section-divider", "closing", "agenda"}


def build_ghost_slide(prs, p, pal, ctx):
    """Skeleton slide: real heading + dashed placeholder describing the
    planned exhibit. Used by build_deck --ghost for storyline alignment."""
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    kind, value = parse_visual(p.get("visual", ""))
    label = f"[ {p.get('layout', 'bullet-list')} ]"
    if kind:
        label += f"   planned visual — {kind}: {value}"
    elif p.get("data"):
        label += "   planned exhibit from Data block"
    box = add_rect(slide, 0.7, 1.9, 11.9, 4.6, pal["surface"],
                   line_hex=pal["text_muted"], line_pt=1.0)
    from helpers import set_fill_alpha
    set_fill_alpha(box, 40)
    add_tb(slide, label, 0.9, 2.05, 11.5, 0.4, size=14, bold=True,
           color=pal["text_muted"], font=pal["font_label"])
    for i, b in enumerate(p.get("bullets", [])[:6]):
        add_tb(slide, f"—  {split_icon(b)[1]}", 1.0, 2.7 + i * 0.55, 11.2,
               0.5, size=12, color=pal["text_muted"], font=pal["font_body"])
    return slide
```

- [ ] **Step 4: Wire the flag**

In `scripts/build_deck.py`:
- `build(...)` signature: add `ghost=False` after `variant=None`.
- In the slide loop, replace `if slide is None:` body:

```python
            if slide is None:
                if ghost and layout_name not in __import__("builders").GHOST_KEEP_REAL:
                    slide = __import__("builders").build_ghost_slide(prs, p, pal, ctx)
                else:
                    slide = LAYOUT_MAP[layout_name](prs, p, pal, ctx)
```

(The file already imports `builders` lazily inside `build`; reuse that import instead of `__import__` if a `builders` name is in scope at that point — it is: `import builders` happens earlier in `build()`. So write it as `slide = builders.build_ghost_slide(...)` / `layout_name not in builders.GHOST_KEEP_REAL`.)

- argparse: `parser.add_argument("--ghost", action="store_true", help="Build a skeleton deck: real titles, placeholder exhibits (storyline alignment)")`
- Pass through: `ok = build(args.outline, args.output, args.palette, args.template, args.assets_dir, args.check, args.size, args.density, args.variant, ghost=args.ghost)`

- [ ] **Step 5: Run tests, commit**

Run: `python3 -m pytest tests/test_build_deck.py -q` → all pass

```bash
git add scripts/builders.py scripts/build_deck.py tests/test_build_deck.py
git commit -m "feat: --ghost mode builds skeleton deck with real titles and placeholder exhibits"
```

---

## Phase 4 — Outline ergonomics & theming

### Task 12: Heading attributes `{layout=x palette=y}`

Quarto/pandoc-style attributes on the slide heading: `## Slide 7: Margin bridge {layout=waterfall palette=aurora}` — keeps directives off the content bullets.

**Files:**
- Modify: `scripts/build_deck.py:112-125` (`## Slide` handling in `parse_outline`)
- Test: `tests/test_build_deck.py` (append)

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_build_deck.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_build_deck.py -q`
Expected: FAIL — heading keeps the `{...}` text, layout not set

- [ ] **Step 3: Implement**

In `parse_outline` (`scripts/build_deck.py`), the `## Slide` branch currently ends with:

```python
            m = re.match(r"## Slide \d+:\s*(.*)", line)
            if m:
                current["heading"] = m.group(1).strip()
            continue
```

Replace with:

```python
            m = re.match(r"## Slide \d+:\s*(.*)", line)
            if m:
                heading = m.group(1).strip()
                m_attr = re.search(r"\{([^{}]*)\}\s*$", heading)
                if m_attr:
                    for k, v in re.findall(r"(\w[\w-]*)=([\w./#-]+)",
                                           m_attr.group(1)):
                        if k == "layout":
                            current["layout"] = v.lower()
                        elif k == "palette":
                            current["palette"] = v.lower()
                    heading = heading[:m_attr.start()].strip()
                current["heading"] = heading
            continue
```

An explicit `**Layout:**` line later in the section simply overwrites the dict key, which gives the override behavior the second test asserts.

- [ ] **Step 4: Run tests, commit**

Run: `python3 -m pytest tests/test_build_deck.py -q` → all pass

```bash
git add scripts/build_deck.py tests/test_build_deck.py
git commit -m "feat: heading attributes {layout=.. palette=..} on slide headings"
```

### Task 13: `chart_series` token + custom JSON palettes

Two theming upgrades from the reveal.js token-schema pattern: (a) chart series colors become a palette token instead of a hardcoded accent triple; (b) users can drop brand palettes as JSON files in `<assets>/palettes/`.

**Files:**
- Modify: `scripts/palettes.py` (token default + `load_custom_palettes`)
- Modify: `scripts/charts.py:100` (use the token)
- Modify: `scripts/build_deck.py` (load customs; relax `--palette` argparse choices)
- Test: `tests/test_features.py` (append)

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_features.py`:

```python
def test_every_palette_has_chart_series_token():
    from palettes import PALETTES
    for key, pal in PALETTES.items():
        assert len(pal.get("chart_series", [])) >= 3, key


def test_load_custom_palette(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    (pdir / "acme-brand.json").write_text(json.dumps({
        "bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
        "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
        "text": "F4F4F4", "text_muted": "9DB2BF", "dark": True}))
    loaded = load_custom_palettes(pdir)
    assert "acme-brand" in loaded and "acme-brand" in PALETTES
    pal = PALETTES["acme-brand"]
    assert pal["font_title"]                      # font defaults filled
    assert len(pal["chart_series"]) >= 3          # token derived
    del PALETTES["acme-brand"]                    # don't leak into other tests


def test_invalid_custom_palette_rejected(tmp_path):
    import json
    from palettes import PALETTES, load_custom_palettes
    pdir = tmp_path / "palettes"
    pdir.mkdir()
    (pdir / "broken.json").write_text(json.dumps({"bg": "101820"}))
    loaded = load_custom_palettes(pdir)
    assert "broken" not in loaded and "broken" not in PALETTES
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m pytest tests/test_features.py -q`
Expected: FAIL — no `chart_series`, no `load_custom_palettes`

- [ ] **Step 3: Implement in `palettes.py`**

Replace the existing setdefault loop (`for _key, _pal in PALETTES.items(): _pal.setdefault("motif", ...)`) with:

```python
FONT_DEFAULTS = {"font_title": "Gill Sans MT", "font_body": "Calibri",
                 "font_label": "Calibri Light"}
REQUIRED_KEYS = {"bg", "bg_deep", "surface", "accent1", "accent2", "accent3",
                 "text", "text_muted", "dark"}


def _fill_defaults(pal):
    pal.setdefault("motif", "icon-circle")
    for k, v in FONT_DEFAULTS.items():
        pal.setdefault(k, v)
    pal.setdefault("chart_series",
                   [pal["accent1"], pal["accent2"], pal["accent3"]])
    return pal


for _key, _pal in PALETTES.items():
    _fill_defaults(_pal)


def load_custom_palettes(palettes_dir):
    """Merge user palette JSONs (one palette per <name>.json) into PALETTES.

    Validates the required token schema; invalid files are skipped with a
    warning (never trust external data). Returns the list of loaded names.
    """
    import json
    import sys
    from pathlib import Path
    palettes_dir = Path(palettes_dir)
    loaded = []
    if not palettes_dir.is_dir():
        return loaded
    for f in sorted(palettes_dir.glob("*.json")):
        try:
            pal = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            print(f"  [WARN] palette {f.name}: unreadable ({e})", file=sys.stderr)
            continue
        missing = REQUIRED_KEYS - set(pal)
        if missing or not isinstance(pal.get("dark"), bool):
            print(f"  [WARN] palette {f.name}: missing keys "
                  f"{sorted(missing)} — skipped", file=sys.stderr)
            continue
        PALETTES[f.stem] = _fill_defaults(pal)
        loaded.append(f.stem)
    return loaded
```

- [ ] **Step 4: Consume the token in `charts.py`**

In `add_native_chart` (`scripts/charts.py:100`), replace:

```python
    accents = [pal["accent1"], pal["accent2"], pal["accent3"]]
```

with:

```python
    accents = pal.get("chart_series") or [pal["accent1"], pal["accent2"],
                                          pal["accent3"]]
```

- [ ] **Step 5: Wire into `build_deck.py`**

In `build()`, directly after `ctx = {...}` is defined:

```python
    from palettes import load_custom_palettes
    custom = load_custom_palettes(ctx["assets_dir"] / "palettes")
    if custom:
        print(f"  Loaded custom palettes: {', '.join(custom)}")
```

In the argparse setup, the `--palette` choices list is now stale (customs load at runtime). Replace:

```python
    parser.add_argument("--palette", default=None,
                        choices=sorted(PALETTES),
                        help="Color palette (overrides outline front-matter)")
```

with:

```python
    parser.add_argument("--palette", default=None,
                        help="Color palette: built-ins "
                             f"({', '.join(sorted(PALETTES))}) or a custom "
                             "palette JSON name from <assets>/palettes/")
```

And in `validate()` the existing per-slide unknown-palette warning already covers typos; add the same check for the deck-level palette at the top of `validate`:

```python
    deck_pal = meta.get("palette")
    if deck_pal and deck_pal not in PALETTES:
        warnings.append(f"unknown deck palette '{deck_pal}' — default will be used")
```

(Note: `load_custom_palettes` runs before `validate` in `build()`, so custom names pass this check.)

- [ ] **Step 6: Run tests, commit**

Run: `python3 -m pytest tests/ -q` → all pass

```bash
git add scripts/palettes.py scripts/charts.py scripts/build_deck.py tests/test_features.py
git commit -m "feat: chart_series palette token and custom JSON palettes from assets/palettes/"
```

### Task 14: Line numbers in `--check` diagnostics

Slidev-parser pattern: every slide records its source line; validation messages become `Slide 3 (line 41): ...` so outline fixes are jump-to-able.

**Files:**
- Modify: `scripts/build_deck.py` (`parse_outline` line tracking, `validate` message prefix)
- Test: `tests/test_build_deck.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_build_deck.py`:

```python
def test_validation_errors_carry_line_numbers():
    from build_deck import parse_outline, validate
    md = ("## Slide 1: Costs have outgrown revenue for years\n"
          "- a bullet\n"
          "- Notes: n.\n"
          "\n"
          "## Slide 2: Bad chart slide misses its data block\n"
          "**Layout:** waterfall\n"
          "- Notes: n.\n")
    _, slides = parse_outline(md)
    assert slides[0]["_line"] == 1 and slides[1]["_line"] == 5
    errors, _ = validate(slides, {"outline_dir": Path("."),
                                  "assets_dir": Path("assets")})
    assert any("Slide 2 (line 5)" in e for e in errors), errors
```

Also add `from pathlib import Path` to the test file's imports if not present.

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_build_deck.py::test_validation_errors_carry_line_numbers -q`
Expected: FAIL with KeyError `'_line'`

- [ ] **Step 3: Implement**

In `parse_outline`, change the loop header from:

```python
    for raw in md_text.splitlines():
```

to:

```python
    for lineno, raw in enumerate(md_text.splitlines(), 1):
```

and in the `## Slide ` branch, right after `current = {...}` is created, add:

```python
            current["_line"] = lineno
```

In `validate`, change:

```python
        where = f"Slide {n}"
```

to:

```python
        where = f"Slide {n} (line {p['_line']})" if p.get("_line") else f"Slide {n}"
```

(Auto-inserted agenda slides from Task 9 have no `_line` — the fallback covers them.)

- [ ] **Step 4: Run the full suite — message-text assertions may need updating**

Run: `python3 -m pytest tests/ -q`
Any test asserting on exact `"Slide N:"` prefixes in validation messages will now see `"Slide N (line L):"`. Fix those assertions to use `in`-style substring checks on the message tail (e.g. assert on `"unknown layout"` rather than the full prefix). Do not weaken what they verify.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_deck.py tests/
git commit -m "feat: validation messages carry outline line numbers"
```

---

## Phase 5 — Edit mode, LLM QA stage, docs

### Task 15: `inventory` / `replace` subcommands in `edit_deck.py`

The tfriedel/claude-office-skills pattern: dump every text run with a stable address as JSON, batch-edit the JSON, apply it back preserving run formatting. Makes template-text swaps surgical instead of regex-on-XML.

**Files:**
- Modify: `scripts/edit_deck.py` (two functions + two subparsers)
- Test: `tests/test_edit_deck.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_edit_deck.py` (it already builds/unpacks a fixture deck — reuse its helpers; if its fixture is named differently, adapt the deck-creation lines to the existing helper):

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_edit_deck.py -q`
Expected: FAIL with AttributeError (`edit_deck.inventory`)

- [ ] **Step 3: Implement**

In `scripts/edit_deck.py`, add after `clean_orphans` (reusing the module's existing `_xml`/`_write_xml` helpers and its `qn`-style namespace handling — match how `duplicate()` reads slide XML):

```python
A_T = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"


def _slide_xml_files(src_dir):
    """slideN.xml paths sorted by N."""
    slides_dir = Path(src_dir) / "ppt" / "slides"
    files = [p for p in slides_dir.glob("slide*.xml") if p.stem[5:].isdigit()]
    return sorted(files, key=lambda p: int(p.stem[5:]))


def inventory(src_dir):
    """Every text run in every slide as [{slide, run, text}] (document order).

    'run' is the index of the <a:t> within its slide — the stable address
    replace_runs() uses. Returns the list (and the CLI prints it as JSON).
    """
    out = []
    for f in _slide_xml_files(src_dir):
        n = int(f.stem[5:])
        tree = _xml(f)
        for i, t in enumerate(tree.iter(A_T)):
            out.append({"slide": n, "run": i, "text": t.text or ""})
    return out


def replace_runs(src_dir, edits_json):
    """Apply [{slide, run, text}] edits; formatting (rPr) is untouched."""
    import json
    edits = json.loads(Path(edits_json).read_text(encoding="utf-8"))
    by_slide = {}
    for e in edits:
        by_slide.setdefault(int(e["slide"]), {})[int(e["run"])] = e["text"]
    for f in _slide_xml_files(src_dir):
        n = int(f.stem[5:])
        if n not in by_slide:
            continue
        tree = _xml(f)
        runs = list(tree.iter(A_T))
        for idx, new_text in by_slide[n].items():
            if idx >= len(runs):
                raise SystemExit(
                    f"slide {n}: run {idx} out of range (has {len(runs)})")
            runs[idx].text = new_text
        _write_xml(tree, f)
        print(f"slide {n}: {len(by_slide[n])} run(s) replaced")
```

In `main()`, add subparsers next to the existing ones:

```python
    p_inv = sub.add_parser("inventory"); p_inv.add_argument("dir")
    p_rep = sub.add_parser("replace"); p_rep.add_argument("dir")
    p_rep.add_argument("edits", help="JSON: [{slide, run, text}, ...]")
```

and dispatch (matching the existing dispatch style in `main()`):

```python
    elif args.cmd == "inventory":
        import json
        print(json.dumps(inventory(args.dir), indent=2, ensure_ascii=False))
    elif args.cmd == "replace":
        replace_runs(args.dir, args.edits)
```

If `_xml`/`_write_xml` signatures differ from this sketch (check `scripts/edit_deck.py:37-44`), adapt the calls — the helpers exist precisely for this.

- [ ] **Step 4: Run tests, commit**

Run: `python3 -m pytest tests/test_edit_deck.py tests/test_edit_extras.py -q` → all pass

```bash
git add scripts/edit_deck.py tests/test_edit_deck.py
git commit -m "feat: edit_deck inventory/replace JSON workflow for surgical text edits"
```

### Task 16: `--numbers` dump + LLM consistency-check stage

UpSlide's killer feature, LLM-native: extract all numeric tokens per slide so Claude can cross-check totals vs components, repeated KPIs, and title claims vs chart data.

**Files:**
- Modify: `scripts/qa_check.py` (add `dump_numbers`, `--numbers` flag)
- Modify: `references/qa-guide.md` (new stage)
- Test: `tests/test_features.py` (append)

- [ ] **Step 1: Write the failing test**

Append to `tests/test_features.py`:

```python
def test_dump_numbers_extracts_per_slide(capsys, tmp_path):
    from pptx import Presentation
    from pptx.util import Inches
    from qa_check import dump_numbers
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    tb.text_frame.text = "Revenue grew 12% to $4.2B in FY25"
    p = tmp_path / "n.pptx"
    prs.save(str(p))
    dump_numbers(p)
    out = capsys.readouterr().out
    assert "Slide 1" in out and "12%" in out and "$4.2B" in out and "FY25" in out
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_features.py::test_dump_numbers_extracts_per_slide -q`
Expected: FAIL with ImportError (`dump_numbers`)

- [ ] **Step 3: Implement**

In `scripts/qa_check.py`, add after `dump_text`:

```python
NUM_TOKEN_RX = re.compile(
    r"[$€£]?\d[\d,.]*\s*(?:%|bn|B|M|k|x|pp|bps)?|FY\d{2,4}|Q[1-4]\s?\d{2,4}",
    re.I)


def dump_numbers(pptx_path):
    """All numeric/period tokens per slide — input for the LLM
    cross-slide consistency check (see references/qa-guide.md)."""
    prs = Presentation(str(pptx_path))
    for n, slide in enumerate(prs.slides, 1):
        tokens = []
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                tokens += [t.strip() for t in
                           NUM_TOKEN_RX.findall(shape.text_frame.text)
                           if t.strip()]
        title = _slide_title_text(slide)
        print(f"Slide {n} [{title[:60]}]: {', '.join(tokens) if tokens else '—'}")
```

In `main()`, add before the `--text` handling:

```python
    if "--numbers" in sys.argv:
        dump_numbers(path)
        sys.exit(0)
```

and update the usage string to `qa_check.py deck.pptx [--text] [--numbers] [--accessibility]`.

- [ ] **Step 4: Add the consistency stage to `references/qa-guide.md`**

Append this section to `references/qa-guide.md`:

```markdown
## Cross-slide consistency check (LLM stage)

Deterministic checks can't catch a title that contradicts its chart or an
exec-summary KPI that doesn't match the body. After the programmatic checks:

1. Run `python3 scripts/qa_check.py deck.pptx --numbers` and
   `python3 scripts/build_deck.py outline.md --titles`.
2. Review the two outputs together and check:
   - **Internal arithmetic** — components sum to stated totals (waterfall
     deltas vs endpoints, funnel stages, "3 levers" vs 3 items).
   - **Repeated KPIs** — the same metric quoted on multiple slides has the
     same value everywhere (exec summary vs body vs appendix).
   - **Title vs exhibit** — each action title's claim ("doubled", "-39%",
     "largest segment") is actually supported by that slide's numbers.
   - **Periods** — FY/quarter labels are consistent (no FY25 vs 2025 drift).
3. Fix discrepancies in the outline (never hand-edit the deck), rebuild,
   re-run. This stage is mandatory for consulting/board decks.
```

- [ ] **Step 5: Run tests, commit**

Run: `python3 -m pytest tests/test_features.py -q` → all pass

```bash
git add scripts/qa_check.py references/qa-guide.md tests/test_features.py
git commit -m "feat: qa_check --numbers dump powering LLM cross-slide consistency stage"
```

### Task 17: Documentation + lint wired into the QA workflow

Every new feature must be discoverable from SKILL.md and the reference guides, and `pptx_lint` becomes a required QA step.

**Files:**
- Modify: `SKILL.md`
- Modify: `references/generation-guide.md`
- Modify: `references/charts-guide.md`
- Modify: `references/qa-guide.md`
- Modify: `references/editing.md`
- Modify: `README.md`

- [ ] **Step 1: Update SKILL.md Phase 4**

Replace the four-item QA list in SKILL.md with (adding lint as item 2 and consistency as item 5):

```markdown
1. **Programmatic:** `python3 scripts/qa_check.py deck.pptx` (add `--accessibility` for WCAG AA strict mode)
2. **Deck lint:** `python3 scripts/pptx_lint.py deck.pptx --palette <palette>` — cross-slide consistency (jiggle, page sequence, off-palette colors, missing fonts)
3. **Content diff:** `python3 scripts/diff_deck.py outline.md deck.pptx` — catches missing text vs outline
4. **Visual:** `python3 scripts/render_slides.py deck.pptx --grid --out assets/qa-thumbs/` + fresh-eyes subagent (grid cells are numbered)
5. **Consistency (LLM):** `python3 scripts/qa_check.py deck.pptx --numbers` + titles test — cross-check totals, repeated KPIs, title claims (see `references/qa-guide.md`)
6. **Fix loop:** edit outline → rebuild → re-run all checks until clean
```

Also in SKILL.md Phase 3, extend the build command line to mention the new flags:

```markdown
2. Run `python3 scripts/build_deck.py outline.md --output deck.pptx [--palette X] [--template T.pptx] [--assets-dir DIR] [--density compact|comfortable] [--variant a|b|c] [--ghost]`
   - `--ghost` builds a skeleton deck (real action titles, grey labeled exhibit placeholders) for storyline sign-off before investing in content.
   - Custom brand palettes: drop `<name>.json` into `<assets>/palettes/` and use `--palette <name>` (schema: `references/generation-guide.md`).
```

- [ ] **Step 2: Update `references/generation-guide.md`**

Append a section (verbatim):

```markdown
## Newer directives (June 2026)

**Heading attributes** — set layout/palette on the heading line, keeping
directives off content bullets:

    ## Slide 7: Margin bridge tells the story {layout=waterfall palette=aurora}

An explicit `**Layout:**` line still wins if both are present.

**Deck meta:**
- `**Auto-Agenda:** on` — auto-insert an agenda slide (sections = your
  section-divider headings) after the title slide.
- `**Auto-Agenda:** track` — additionally insert a "where we are" agenda with
  the current section highlighted after every section divider.
- `**Stamp:** DRAFT` — bordered status tag on every slide (also:
  CONFIDENTIAL, FOR DISCUSSION).

**Chart/exhibit directives (per slide):**
- `- Axis-Max: 50` — pin the chart value axis; use the same value on sibling
  slides for honest same-scale comparison.
- `- CAGR: on` — growth arrow from first to last column (bar/column/line,
  single series).
- `- Bracket: FY25, FY27` — waterfall difference bracket between two named
  bars; auto-labels the % change. Optional third part = custom label:
  `- Bracket: FY25, FY27, "Run-rate reset"`.

**New layouts:**
- `bar-mekko` — profit pool: `- Bar: Label="EMEA" Size="30" Value="6"`
  (width ∝ Size, height = Value). 2+ bars required.
- `matrix-2x2` bubbles — add `Size="40"` to `- Item:` rows; bubble area
  scales with Size (BCG growth–share style).

**Custom palettes:** drop `<name>.json` into `<assets>/palettes/`:

    {"bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
     "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
     "text": "F4F4F4", "text_muted": "9DB2BF", "dark": true}

Optional keys: `font_title`, `font_body`, `font_label`, `chart_series`
(list of hex — chart series colors), `motif`. Then `--palette <name>`.
```

- [ ] **Step 3: Update `references/charts-guide.md`**

Append:

```markdown
## Annotations

- **Benchmark line:** `- Benchmark: 120 Industry average` (existing).
- **CAGR arrow:** `- CAGR: on` — computed from first/last values of a
  single-series bar/column/line chart. Plot-box placement is approximate:
  always confirm in visual QA.
- **Same scale:** `- Axis-Max: N` pins the value axis (applied last, so it
  wins over the axis max set by Benchmark/CAGR). Use the same N across
  slides being compared.
- **Waterfall bracket:** `- Bracket: <bar A>, <bar B>[, "label"]` — exact
  geometry (shape-drawn), no QA caveat.
```

- [ ] **Step 4: Update `references/editing.md`**

Append:

```markdown
## Batch text edits: inventory → replace

For text-only changes across many slides, prefer the JSON workflow over
hand-editing XML:

    python3 scripts/edit_deck.py unpack deck.pptx work/
    python3 scripts/edit_deck.py inventory work/ > inventory.json
    # edit the "text" fields you want to change (keep slide/run addresses),
    # save the changed entries (only those) as edits.json
    python3 scripts/edit_deck.py replace work/ edits.json
    python3 scripts/edit_deck.py pack work/ deck-edited.pptx

Run formatting (bold/size/color) is preserved — only the text changes.
Then run the standard Phase 4 QA on the result.
```

- [ ] **Step 5: Update README.md feature list**

Add to the README's feature overview (match its existing list style): deck lint (`pptx_lint.py`), font preflight, numbered QA grid, CAGR/bracket annotations, `Axis-Max`, `bar-mekko`, matrix bubbles, `Auto-Agenda`, `Stamp`, `--ghost` mode, heading attributes, custom JSON palettes, `edit_deck inventory/replace`, `qa_check --numbers`.

- [ ] **Step 6: Commit**

```bash
git add SKILL.md references/ README.md
git commit -m "docs: document lint stage, annotations, new layouts, automation and palette features"
```

### Task 18: Final verification

**Files:** none

- [ ] **Step 1: Full test suite + smoke test**

```bash
python3 -m pytest tests/ -q
python3 scripts/smoke_test.py
```

Expected: all pass; smoke deck builds.

- [ ] **Step 2: End-to-end with new features**

Create `/tmp/e2e-outline.md`:

```markdown
**Palette:** midnight-executive
**Footer:** Acme · Confidential
**Page-Numbers:** on
**Auto-Agenda:** track
**Stamp:** DRAFT

## Slide 1: Acme FY27 cost reset {layout=title}
- Title: Acme FY27 Cost Reset
- Subtitle: Steering committee · June 2026
- Notes: Opening.

## Slide 2: Diagnosis
**Layout:** section-divider
- Subtitle: Where the money goes

## Slide 3: Three levers bridge run-rate down to target {layout=waterfall}
**Data:**
- FY25: 46
- Tiering: -8
- Exit: -6
- Discounts: -4
- FY27: total
- Bracket: FY25, FY27
- Source: Acme finance, May 2026
- Notes: The bridge.

## Slide 4: EMEA is the margin outlier despite its size {layout=bar-mekko}
- Bar: Label="Americas" Size="55" Value="14"
- Bar: Label="EMEA" Size="30" Value="6"
- Bar: Label="APAC" Size="15" Value="11"
- Source: Acme finance
- Notes: Profit pool.
```

```bash
python3 scripts/build_deck.py /tmp/e2e-outline.md --check
python3 scripts/build_deck.py /tmp/e2e-outline.md --output /tmp/e2e.pptx
python3 scripts/qa_check.py /tmp/e2e.pptx
python3 scripts/pptx_lint.py /tmp/e2e.pptx --palette midnight-executive
python3 scripts/build_deck.py /tmp/e2e-outline.md --output /tmp/e2e-ghost.pptx --ghost
python3 scripts/render_slides.py /tmp/e2e.pptx --grid --out /tmp/e2e-thumbs/
```

Expected: check OK; build saves 6 slides (4 authored + overview agenda + 1 tracker — note slide 4 divider also gets a tracker if a second divider section is added, with this outline there is 1 divider → 6 slides total); qa_check 0 errors; lint exit 0; ghost build saves; grid renders with numbered cells. Visually inspect `/tmp/e2e-thumbs/grid.png`: DRAFT tag top-left, tracker agenda highlights "Diagnosis", waterfall bracket reads "-39%", bar-mekko widths 55/30/15.

- [ ] **Step 3: Update the smoke outline to exercise one new feature**

Add to `assets/example-outline.md` one `- Bracket:` line on its waterfall slide (if it has one) or one `Axis-Max` on a chart slide, so `smoke_test.py` permanently covers an annotation. Re-run `python3 scripts/smoke_test.py`.

- [ ] **Step 4: Commit and present the branch**

```bash
git add assets/example-outline.md
git commit -m "test: exercise chart annotation in smoke outline"
git log --oneline main..HEAD
```

Then use superpowers:finishing-a-development-branch to decide merge/PR.

---

## Deferred (explicitly out of scope — YAGNI for this round)

- **ChartEx (`cx:`) native waterfall XML injection** — high effort, no library support; shape-based waterfall is correct and editable enough. Revisit only on user demand.
- **Horizontal/stacked waterfall and value-chain layouts** — straightforward variants of existing builders (`build_waterfall_slide`, `build_process_flow_slide`); add when a real deck needs one rather than speculatively.
- **SVG dual-blip embedding & `p14:sectionLst` sections** — well-bounded lxml features, but no current workflow needs them.
- **think-cell `.ppttc` export** — only valuable to users with think-cell licenses; needs a user request to justify.
- **Brand-from-URL ingestion** — custom JSON palettes (Task 13) cover the need; URL scraping adds fragility.
- **Per-layout deterministic reflow engine** (Beautiful.ai pattern) — `smart_layout.py` + density tables + QA loop already approximate this; a constraint solver is a separate project.
- **Round-trippable outline serializer** — line numbers (Task 14) deliver most of the diagnostic value; full parse→mutate→save is a refactor to schedule separately.
