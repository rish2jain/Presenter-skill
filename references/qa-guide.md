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

**Accessibility / projection mode** (stricter body-font floor, alt-text on images):
```bash
python3 scripts/qa_check.py output.pptx --accessibility
```

For a proofreading pass (typos, wrong numbers, leftover draft phrasing):
```bash
python3 scripts/qa_check.py output.pptx --text
```

`diff_deck.py` compares outline phrases to extracted deck text (MarkItDown when installed, else python-pptx). Use `--strict` to fail CI on missing copy.

---

## Step 1 — Render All Slides as Images

```bash
python3 scripts/render_slides.py output.pptx --grid --out assets/qa-thumbs/
```

This creates `slide-01.png`, `slide-02.png`, … plus a stitched `grid.png`.

If LibreOffice is not available the script falls back to a Pillow renderer — **the fallback dumps text only and cannot be used to judge layout, images, or charts**. Do not treat visual QA as complete without LibreOffice — rely on `qa_check.py` and opening the `.pptx` directly until LibreOffice is installed.

---

## Step 2 — Visual Inspection (use a fresh-eyes subagent)

**Dispatch a subagent to inspect the thumbnails — even for a 3-slide deck.** The agent that generated the slides systematically under-reports its own defects; a fresh context judging only the rendered images catches what the generator rationalizes away. Give the subagent the thumbnail paths and the checklist below, and ask it to return a numbered defect list per slide.

Flag any of the following:

**Layout Issues**
- [ ] Text overflowing outside slide boundaries
- [ ] Text boxes overlapping each other
- [ ] Images cropped unexpectedly or stretched
- [ ] Elements too close to slide edges (< 0.4" margin)
- [ ] Cards/columns misaligned or uneven widths

**Typography Issues**
- [ ] Title text too small (should be ≥ 32pt)
- [ ] Body text too small (should be ≥ 14pt)
- [ ] Low contrast: light text on light background, dark text on dark background
- [ ] Placeholder text not replaced ("Click to add title", "XXXX", "Lorem ipsum")
- [ ] Inconsistent font usage across slides

**Image Issues**
- [ ] User images not appearing (missing file path)
- [ ] Images covering text unintentionally
- [ ] Logo appears dark-on-dark or washed out
- [ ] Image aspect ratio distorted (stretched/squashed)

**Chart Issues**
- [ ] Chart axes labels unreadable
- [ ] Chart data doesn't match the outline
- [ ] Chart background is white on a dark slide (fix: set plot area fill to transparent)

**Content Issues**
- [ ] More than 5 bullet points on a single slide (split into two)
- [ ] Slide looks empty or has too little content
- [ ] No visual element on a content slide (add icon, chart, or image)

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
