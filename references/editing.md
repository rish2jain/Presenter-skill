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
