# QA Guide

## Philosophy

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Treat QA as a bug hunt, not a confirmation step. If you find zero issues on first inspection, look harder.

---

## Step 0 — Programmatic Checks (run first, takes seconds)

```bash
python3 scripts/qa_check.py output.pptx
python3 scripts/diff_deck.py outline.md output.pptx   # content vs outline
```

`qa_check.py` catches mechanically-detectable defects: shapes outside slide bounds, text under 10pt, leftover placeholder text, opaque chart backgrounds, WCAG contrast failures (<3:1 error, <4.5:1 small-text warning), word-dense slides, empty slides, duplicate slide titles, and likely text overflow. Exit code 1 = errors must be fixed.

**Accessibility / projection mode** (mirrors the MS Accessibility Checker):
```bash
python3 scripts/qa_check.py output.pptx --accessibility
```

Beyond the stricter contrast (4.5:1) and body-font floor (18pt), accessibility mode adds:
- **Slide titles:** every slide must have a detectable title (error), and duplicate titles (case-insensitive) across the deck are reported once with all slide numbers — screen-reader navigation depends on unique titles.
- **Tables:** must be marked with a header row (`firstRow`) — error; merged cells get a warning ("screen readers may misread — prefer simple structure"). Decks built by this skill's table builder pass both.
- **Alt text:** missing alt text escalates to error, and alt text that is just a filename (`hero.png`, `image3`) is an error too — describe the image instead.
- **Reading order:** warns when the title shape is not the first text-bearing shape in document order (screen readers announce content in spTree order, not visual order).

**OOXML integrity** (optional schema validation):
```bash
python3 scripts/qa_check.py output.pptx --integrity
```

With the optional `openxml-audit` package installed (`pip install openxml-audit`), `--integrity` validates the .pptx OOXML structure and reports schema problems as errors prefixed `integrity:` (they fail the gate like any other error). Without the package it prints a single `[INFO] pip install openxml-audit …` line and is skipped — never a failure. The CI smoke test runs this as an advisory step.

For a proofreading pass (typos, wrong numbers, leftover draft phrasing):
```bash
python3 scripts/qa_check.py output.pptx --text
```

`diff_deck.py` compares outline phrases to extracted deck text (MarkItDown when installed, else python-pptx). Use `--strict` to fail CI on missing copy.

---

## Step 0b — Deck Lint (cross-slide consistency)

```bash
python3 scripts/pptx_lint.py output.pptx [--palette <palette>]
```

Complements `qa_check.py` with deck-wide checks that per-slide inspection misses:
- **Anti-jiggle:** recurring elements (page numbers, footers, kickers) must sit at identical coordinates on every slide.
- **Page-number sequence:** consecutive, no gaps or duplicates.
- **Font inventory:** flags decks with more distinct fonts than the threshold (inconsistent branding).
- **Palette whitelist:** with `--palette`, any run or fill color outside the palette (plus known extras) is reported as an error.
- **Google Slides compatibility** (`--gslides`): when the deck will be imported into Google Slides, warns on fonts outside the Google Slides set (they get substituted), SmartArt (flattens to a static image), non-fade transitions (dropped or changed), and embedded audio/video (does not import). Run this only when the user mentions Google Slides — the default palettes intentionally use PowerPoint-native fonts.

For custom-brand decks pass `--assets-dir` so the palette whitelist resolves: `python3 scripts/pptx_lint.py deck.pptx --palette <name> --assets-dir <assets>`. Note: slides using per-slide `{palette=...}` overrides will flag as off-palette against the deck palette — lint against the dominant palette and review those slides by eye.

Run this after `qa_check.py`; fix jiggle and off-palette colors in the outline before visual inspection.

---

## Step 1 — Render All Slides as Images

```bash
python3 scripts/render_slides.py output.pptx --grid --out assets/qa-thumbs/
```

This creates `slide-01.png`, `slide-02.png`, … plus a stitched `grid.png`.

If LibreOffice is not available the script falls back to a Pillow renderer — **the fallback dumps text only and cannot be used to judge layout, images, or charts**. Do not treat visual QA as complete without LibreOffice — rely on `qa_check.py` and opening the `.pptx` directly until LibreOffice is installed.

**Optional higher-fidelity renderers.** LibreOffice remains the default, but two alternates render closer to real PowerPoint when available: ONLYOFFICE's `x2t` converter (ships with ONLYOFFICE Document Server / Desktop Editors; converts .pptx → PDF with a rendering engine built for OOXML fidelity), and Microsoft Graph PDF export (`GET /drive/items/{id}/content?format=pdf` — uploads the deck to OneDrive/SharePoint and uses Microsoft's own rendering service; requires an M365 tenant and auth). Reach for either when LibreOffice output looks suspect — font metrics, chart label placement, effects — then `pdftoppm` the PDF into thumbnails as usual.

---

## Step 1.5 — Geometry Self-Check (instant — actually run it *before* Step 1's render)

```bash
python3 scripts/geometry_report.py output.pptx [--json] [--slides 2,5]
```

Deterministic per-slide layout metrics — no LibreOffice, no images, pure geometry. It is cheap enough to run before spending a render cycle, and it pinpoints exact shapes and offsets that thumbnails only show fuzzily, feeding the fix loop precise targets.

Per slide it reports:
- **Overlaps:** pairs of shapes that intersect by more than 0.02 sq in, excluding containment (a shape fully inside another is the intentional card/background pattern).
- **Uneven spacing:** rows/columns of similar-sized shapes whose gap sequence is inconsistent ("uneven row spacing: gaps 0.31/0.29/0.55in").
- **Near-misses:** shape edges that are *almost* aligned — off by 0.04–0.08in. Smaller offsets are invisible at render size; larger ones are presumed intentional.
- **Whitespace ratio** (1 − covered/slide area, rasterized so overlaps aren't double-counted) plus **visual-mass imbalance** between left/right and top/bottom halves.
- **Word count** — more than 90 words is flagged as text overload.

Human output lists only slides with findings and ends with a summary line; `--json` emits full metrics for every slide (useful even when nothing is flagged — e.g. to compare whitespace ratios across sibling slides). Decks built by this skill's builders should report zero findings; anything it prints is worth fixing in the outline before rendering and inspecting.

---

## Step 2 — Visual Inspection (use a fresh-eyes subagent)

**Dispatch a subagent to inspect the thumbnails — even for a 3-slide deck.** The agent that generated the slides systematically under-reports its own defects; a fresh context judging only the rendered images catches what the generator rationalizes away. Give the subagent the thumbnail paths and the flaw checklist below.

**Use the checklist verbatim — do not ask for a holistic impression.** Binary per-flaw questions catch significantly more real defects than open-ended "review this slide" prompts (SlideAudit, UIST'25). For **each numbered grid cell**, the subagent answers yes/no per flaw, with one line of evidence for every yes:

1. **Overlap / collision** — do any two elements intersect (text on text, label on bar, icon on card edge)?
2. **Misalignment** — do element edges sit off the implied grid shared by their siblings (cards, columns, bullets)?
3. **Crowding / whitespace imbalance** — is one half dense while the other is empty, or content pinched against edges (< 0.4" margin)?
4. **Weak title hierarchy** — is the title *not* the most visually dominant text on the slide?
5. **Off-palette color** — does any element use a color that doesn't belong to the deck palette?
6. **Text overload** — paragraph walls, more than 5 bullets, or body text rendered too small to read?
7. **Image quality** — stretched, pixelated, washed-out, oddly cropped, or missing imagery (incl. dark-on-dark logos)?
8. **Sibling inconsistency** — do fonts, margins, or title positions differ from neighboring slides?

`geometry_report.py` (Step 1.5) already detects flaws 1–3 deterministically — the subagent's value is flaws 4–8, plus any geometry the report missed (it cannot see rendered text wrap, fonts, or image content).

**Deck-level rubric.** After the per-cell pass, the subagent scores the whole deck 1–5 on three dimensions (PPTEval):

- **Content** — text quality, every title's claim supported by its slide, images relevant.
- **Design** — color harmony, visual hierarchy, layout craft.
- **Coherence** — the storyline flows logically slide to slide.

Record all three scores in your QA summary. **Any dimension below 4 routes back to the fix loop (Step 3)** — use the per-cell yes answers to locate the offending slides and fix them in the outline.

---

## Step 3 — Fix & Re-render Loop

1. Fix all flagged issues in the **outline** (or edited slide XML for Mode D)
2. Re-run `python3 scripts/build_deck.py` (or `edit_deck.py pack` for Mode D)
3. Re-render affected slides only:
   ```bash
   python3 scripts/render_slides.py output.pptx --slides 3,5,7 --out assets/qa-thumbs/
   ```
4. Re-inspect until a full pass finds zero issues

**Do not deliver until at least one full fix-and-verify cycle is complete.**

---

## Step 4 — Final Gate

Re-run `python3 scripts/qa_check.py output.pptx` after the last fix — it must exit 0 (no errors) before delivery.

---

## Revision regression (before sending a revised deck)

A revision should change exactly the slides you touched — anything else is collateral damage (reflowed text, a shifted chart, a renumbered footer). Keep the thumbnails of the last-delivered deck as a baseline and diff every revision against it:

```bash
python3 scripts/render_slides.py revised.pptx --out assets/qa-current/
python3 scripts/visual_regress.py assets/qa-baseline/ assets/qa-current/ [--threshold N]
```

- Same-named slide PNGs are compared by perceptual hash — pHash via the optional `imagehash` package when installed, else a built-in Pillow-only average hash. Hamming distance > 5 (tune with `--threshold`) flags the slide and reports a coarse pixel-diff percentage.
- **New slides** (present only in current) are reported as `new (no baseline)` — a warning, not a failure.
- **Deleted slides** (present only in baseline) are failures — removing content must be deliberate.
- Exit 1 on any changed or missing slide. Verify every flagged slide is an intended edit before sending.
- After delivering, bless the new render as the next baseline:
  ```bash
  python3 scripts/visual_regress.py assets/qa-baseline/ assets/qa-current/ --update
  ```
  `--update` also initializes a missing baseline on first delivery (and a run against a missing baseline without `--update` exits with instructions to do exactly that).

---

## Slide Count Per Delivery

Always deliver in one message:
1. The `.pptx` file path (attach or share via your environment's file delivery)
2. A thumbnail grid image showing all slides
3. A brief summary: "Here's your [N]-slide deck on [topic]. Covers: [list key slides]. Ready to edit in PowerPoint."

---

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
