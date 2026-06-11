---
name: presentation-skill
description: |
  Use this skill whenever the user wants to create, generate, build, or export a presentation, slide deck, pitch deck, PowerPoint file, or .pptx file. Also trigger for consulting decks, strategy decks, client decks, steering-committee or board decks, exec readouts, and consulting visuals (waterfall/bridge charts, 2x2 matrix, harvey balls, mekko, gantt roadmaps). Also trigger when the user says "make slides", "build a deck", "create a presentation", "generate a PowerPoint", or uploads images and asks to use them in slides, or provides a .pptx template to match.
  Do NOT trigger for: HTML slide decks only (handled by slides skill), simple one-off diagram generation, or image editing without a slide output.
---

# Presentation Deck Skill

Generate professional, visually rich native PowerPoint (.pptx) presentations. See `references/decision-tree.md` for workflow routing.

## Quick Reference

| Goal | Path |
|------|------|
| Pick workflow | Read `references/decision-tree.md` |
| Create from scratch | Read `references/generation-guide.md` |
| Pitch / strategy / talk arcs | Read `references/narratives/*.md` |
| Consulting storyline (SCQA, action titles, titles test) | Read `references/storyline.md` |
| Edit an existing .pptx | Read `references/editing.md` |
| Use a user's .pptx template | Read `references/template-mode.md` |
| Embed user-supplied images | Read `references/image-handling.md` |
| Generate charts natively | Read `references/charts-guide.md` |
| QA and render check | Read `references/qa-guide.md` |

---

## Phase 0 — Detect Mode

Determine the correct mode before proceeding:

**Mode A: From Scratch**
No template file provided. Claude generates a fully styled deck using the outline + design system.

**Mode B: Template-Based**
User provides an existing .pptx file to use as the visual template. Read `references/template-mode.md`.

**Mode C: Image-Rich Deck**
User has shared one or more images (photos, diagrams, screenshots, brand assets) to embed in specific slides. Read `references/image-handling.md`. This mode can be combined with A or B.

**Mode D: Edit Existing Deck**
User provides a finished .pptx and wants changes (fix text, swap images, add/remove/reorder slides). Do NOT rebuild — edit surgically. Read `references/editing.md`. Skip Phases 1–3 and go straight to the editing workflow + Phase 4 QA.

---

## Phase 1 — Content & Style Discovery

Ask ONE clarifying message covering all of the below before starting. Combine into a single natural prompt, not a numbered list.

1. **Topic / title** — What is the presentation about? Who is the audience?
2. **Length** — Approximately how many slides? (Short 5–10 / Medium 10–20 / Long 20+)
3. **Purpose** — Pitch deck / Internal strategy / Client presentation / Conference talk / Tutorial
4. **Tone** — Bold & dark / Clean & corporate / Warm & editorial / Technical & minimal
5. **User images** — Has the user shared any images? If yes, ask which slides to place them on (see `references/image-handling.md`).
6. **Template** — Has the user provided a .pptx template file? If yes, switch to Mode B.

After gathering answers, research the topic if needed using available tools.

---

## Phase 2 — Outline (User Approval Gate)

0. **Strategy/consulting decks first draft the storyline** (`references/storyline.md`): SCQA the executive summary, dot-dash the body, then run the titles test (`build_deck.py outline.md --titles`). Action titles only — full-sentence takeaways, never topic labels.
1. Write the slide outline in markdown (one `## Slide N:` section per slide) using the exact syntax in `references/generation-guide.md`. Follow its Design Rules: vary layouts, use icons (`icon:name`) on bullet slides, write `- Notes:` speaker notes for every slide, and `- Source:` on every exhibit.
2. Validate it: `python3 scripts/build_deck.py outline.md --check` — fix errors; address warnings (missing notes, etc.).
3. **Show the outline to the user and ask for approval or edits before building.** The outline is the reviewable artifact; changing a slide here is cheap, regenerating a deck is not. Skip this gate only if the user explicitly asked for a one-shot build.

---

## Phase 3 — Generation

1. If user images are involved, normalize first: `python3 scripts/prep_images.py assets/user-images/`
2. Run `python3 scripts/build_deck.py outline.md --output deck.pptx [--palette X] [--template T.pptx] [--assets-dir DIR] [--density compact|comfortable] [--variant a|b|c] [--ghost]`
   - `--ghost` builds a skeleton deck (real action titles, grey labeled exhibit placeholders) for storyline sign-off before investing in content.
   - Custom brand palettes: drop `<name>.json` into `<assets>/palettes/` and use `--palette <name>` (schema: `references/generation-guide.md`).
3. The build fails fast on validation errors. If any slide fails during build, **no .pptx is written**.

Optional: `python3 scripts/gen_appendix.py outline.md` for pitch/strategy appendix skeleton.

---

## Phase 4 — QA (Required)

Never skip QA. Six complementary checks (see `references/qa-guide.md`):

1. **Programmatic:** `python3 scripts/qa_check.py deck.pptx` (add `--accessibility` for WCAG AA strict mode)
2. **Deck lint:** `python3 scripts/pptx_lint.py deck.pptx --palette <palette>` — cross-slide consistency (jiggle, page sequence, off-palette colors, missing fonts)
3. **Content diff:** `python3 scripts/diff_deck.py outline.md deck.pptx` — catches missing text vs outline
4. **Visual:** `python3 scripts/render_slides.py deck.pptx --grid --out assets/qa-thumbs/` + fresh-eyes subagent (grid cells are numbered)
5. **Consistency (LLM):** `python3 scripts/qa_check.py deck.pptx --numbers` + titles test — cross-check totals, repeated KPIs, title claims (see `references/qa-guide.md`)
6. **Fix loop:** edit outline → rebuild → re-run all checks until clean

Deliver the `.pptx` and the thumbnail grid together.

---

## Dependencies

```bash
pip install -r requirements.txt
pip install pillow-heif   # only needed for HEIC (iPhone) image inputs
pip install cairosvg      # optional: Lucide SVG icon fallback (Tabler PNG needs nothing)
# Optional for high-fidelity QA rendering:
# LibreOffice (soffice) + poppler (pdftoppm)
# Optional: pip install unoserver, then run `unoserver` in the background —
# render_slides.py auto-uses it (much faster repeated QA renders)
```

Smoke test (validate + build + QA on the example outline):

```bash
python3 scripts/smoke_test.py
pytest tests/
```
