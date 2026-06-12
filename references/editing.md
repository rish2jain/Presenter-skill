# Editing Existing Decks (Mode D)

Use this when the user provides a .pptx and wants changes — fix text, swap an image, add/remove/reorder slides — while preserving the existing design exactly. Never rebuild a deck the user already has; edit it surgically.

## Workflow

```bash
python3 scripts/render_slides.py original.pptx --grid --out review/   # see what exists
python3 scripts/edit_deck.py unpack original.pptx unpacked/           # extract + pretty-print
python3 scripts/edit_deck.py list unpacked/                           # slide order + titles
# 1) structural changes FIRST (N = 1-based position from list, NOT slide file number):
python3 scripts/edit_deck.py duplicate unpacked/ 3                    # clone slide at position 3
python3 scripts/edit_deck.py remove unpacked/ 7                       # drop slide at position 7
python3 scripts/edit_deck.py reorder unpacked/ 3,1,2,4               # new slide order by position
python3 scripts/edit_deck.py clean unpacked/                          # remove orphan media files
# 2) then content edits: Edit tool on unpacked/ppt/slides/slideN.xml
python3 scripts/edit_deck.py pack unpacked/ edited.pptx               # validate + re-zip
python3 scripts/qa_check.py edited.pptx                               # QA gate
```

`pack` refuses to build if any XML fails to parse, so syntax mistakes surface immediately instead of as a corrupt file PowerPoint rejects.

## Editing slide XML

- Text lives in `<a:t>` elements inside runs (`<a:r>`). Edit the text between the tags; never touch the surrounding `<a:rPr>` (run formatting) unless changing formatting is the goal.
- One paragraph = one `<a:p>`. To bold a header run: `<a:rPr b="1" .../>`.
- **Never insert literal unicode bullets (•)** — bullet glyphs come from `<a:buChar>`/`<a:buAutoNum>` in paragraph properties.
- Leading/trailing spaces need `xml:space="preserve"` on the `<a:t>`.
- Smart quotes in new text: use XML entities `&#x201C; &#x201D; &#x2018; &#x2019;`.
- To reorder slides, rearrange `<p:sldId>` elements in `ppt/presentation.xml` → `<p:sldIdLst>` (the order there IS the deck order).
- To swap an image: replace the file in `ppt/media/` keeping the same filename and format.

## Gotchas

| Issue | Rule |
|-------|------|
| Manual slide file copying | Never — `duplicate` updates Content_Types, rels, and sldIdLst; hand-copying misses them and corrupts the deck |
| Structural vs content edits | Do ALL duplicate/remove/reorder operations BEFORE text edits |
| Position vs file number | `duplicate` / `remove` take the **1-based position** shown by `list`, not `slideN.xml` file numbers |
| Template slots ≠ your items | When a layout has 4 boxes and you have 3 items, delete the whole leftover group (shape + text), not just its text |
| Longer replacement text | May overflow the original box — always run visual QA after edits |
| Notes on duplicated slides | `duplicate` intentionally drops the notes link; add new notes if needed |

## Quick text-only fixes

For a pure find/replace (a typo, a number), prefer python-pptx in place of unpacking:

```python
from pptx import Presentation
prs = Presentation("original.pptx")
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.text = run.text.replace("2025", "2026")
prs.save("edited.pptx")
```

This preserves all formatting because only run text changes.

## Batch text edits: inventory → replace

For text-only changes across many slides, prefer the JSON workflow over
hand-editing XML:

    python3 scripts/edit_deck.py unpack deck.pptx work/
    python3 scripts/edit_deck.py inventory work/ > inventory.json
    # edit the "text" fields you want to change (keep slide/run addresses),
    # save the changed entries (only those) as edits.json
    python3 scripts/edit_deck.py replace work/ edits.json
    python3 scripts/edit_deck.py pack work/ deck-edited.pptx

Addresses are slideN.xml file numbers (these differ from deck position
after a reorder — correlate with the `list` subcommand). Run formatting
(bold/size/color) is preserved — only the text changes. Then run the
standard Phase 4 QA on the result.

## Slide operations playbook

Precise recipes for common edit requests, built from the existing
primitives. Nothing here needs new tooling — the discipline is in the
sequencing.

### condense / rewrite / formalize / translate

All four are the same mechanical recipe — only the text you write differs
(shorter, reworded, formal register, target language):

```bash
python3 scripts/edit_deck.py unpack deck.pptx work/
python3 scripts/edit_deck.py inventory work/ > inventory.json
```

Inventory output (snippet) — every text run with its stable address:

```json
[
  {"slide": 2, "run": 0, "text": "Three priorities driving $14.2M in value"},
  {"slide": 2, "run": 1, "text": "Phase 1 GPU cluster deployed ahead of schedule, saving $2.1M annually"},
  {"slide": 2, "run": 2, "text": "Hybrid cloud migration 62% complete — on track for Q4"}
]
```

Write `edits.json` containing ONLY the runs you change, keeping the
`slide`/`run` addresses (here: condensing run 1):

```json
[
  {"slide": 2, "run": 1, "text": "Phase 1 GPUs live early — $2.1M/yr saved"}
]
```

```bash
python3 scripts/edit_deck.py replace work/ edits.json
python3 scripts/edit_deck.py pack work/ deck-edited.pptx
python3 scripts/qa_check.py deck-edited.pptx
```

**Run-granularity caveat:** one visual sentence can be split across several
runs (e.g. a bold figure mid-sentence produces three runs: before / bold /
after). `replace` works at run level, so rewrite each piece in place — or
put the whole new sentence in the first run and set the leftover runs of
that sentence to `""`. Never merge runs by hand-editing XML for this; you
lose the formatting boundaries. For translate, keep numbers/units as their
own runs when they already are — only the prose runs change.

### split (one overloaded slide → two)

1. `python3 scripts/edit_deck.py list work/` — find the slide's position N.
2. `python3 scripts/edit_deck.py duplicate work/ N` — the clone lands at
   position N+1 (notes link intentionally dropped).
3. `inventory` → write two edit sets: in the original, blank (`""`) the runs
   for the half that moves; in the clone (a NEW slideN.xml file number —
   re-run `inventory` to see it), blank the half that stayed, and retitle
   both (e.g. "… (1/2)" / "… (2/2)" or sharper per-half titles).
4. `replace` → `pack` → visual QA. Blanked runs leave empty shapes behind;
   if a whole box/group is now empty, delete that shape's `<p:sp>` in the
   slide XML rather than leaving a ghost slot.

### remix (re-layout a slide)

- **Deck built from an outline (Mode A/B):** edit that slide's section in
  the outline — change `**Layout:**` to the new layout — and rebuild. For a
  surgical single-slide rebuild instead of a full one: build a one-slide
  outline with the same palette, then merge and swap:

  ```bash
  python3 scripts/build_deck.py remix-slide.md --output one.pptx --palette <same>
  python3 scripts/edit_deck.py append deck.pptx one.pptx --output tmp.pptx
  python3 scripts/edit_deck.py unpack tmp.pptx work/
  python3 scripts/edit_deck.py remove work/ N          # the old slide
  python3 scripts/edit_deck.py reorder work/ ...       # move new slide into N's place
  python3 scripts/edit_deck.py pack work/ deck-remixed.pptx
  ```

- **Template-mode decks:** `add_slide.py deck.pptx out.pptx --layout K`
  adds a blank slide from the template's own layout K; fill placeholders
  with python-pptx, then `remove`/`reorder` the old slide as above.
- **Foreign decks (no outline, no usable template layouts):** there is no
  automated remix — the slide's design exists only as its own XML. Rebuild
  the slide content on the nearest skill layout (it will carry skill
  styling, not the source deck's), or restructure within the existing
  layout via `replace` + shape edits.

## Splitting a deck: extract

Pull a contiguous or cherry-picked subset into a new standalone deck —
positions are the 1-based order `list` shows:

```bash
python3 scripts/edit_deck.py extract deck.pptx 3-7 --output sub.pptx
python3 scripts/edit_deck.py extract deck.pptx 1,3-5,9 --output sub.pptx
```

The subset keeps deck order regardless of how the selection is written.
Unreferenced leftovers (notes, charts, embedded workbooks, media of dropped
slides) are garbage-collected so the output is clean. `--output` must
differ from the input. Speaker notes of KEPT slides survive.

## Merging decks: append

Copy slides from one deck into another, after the destination's last slide,
preserving the source slides' look:

```bash
python3 scripts/edit_deck.py append dst.pptx src.pptx --output merged.pptx
python3 scripts/edit_deck.py append dst.pptx src.pptx --slides 2-4 --output merged.pptx
```

Each copied slide brings its full dependency graph — slide layout, slide
master, theme, images, charts (including chart colors/style and the
embedded .xlsx data) — renamed to avoid collisions and re-registered in
content types and `presentation.xml`. Slides sharing a layout/master are
deduplicated within one append run.

**Known limitations:**

- Source slides keep their own master/theme — faithful import, no
  restyling. The merged deck shows each slide as it looked in its source
  deck; expect a visible style seam unless both decks share a design. To
  restyle instead, re-author the content on the destination deck's layouts.
- Speaker notes on copied slides are dropped (same policy as `duplicate`).
- If the decks have different slide sizes (16:9 vs 4:3), copied shapes keep
  their source geometry on the destination's slide size — check visually.
- Slide-to-slide hyperlinks pointing outside the copied selection are
  dropped (a warning is printed).
- Appended slides keep their original baked page numbers — re-number via
  `inventory`/`replace` if needed.

After any merge, run the standard QA gate (`qa_check.py`, render review) —
and flip through the seam slides specifically.
