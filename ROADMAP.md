# Enhancement Roadmap

Research-driven roadmap (June 2026) from a five-stream survey: Anthropic's official pptx
skill + OSS peers, the python-pptx ecosystem, professional consulting tooling
(think-cell / Mekko Graphics / Macabacus / UpSlide), commercial AI deck generators
(Gamma / Beautiful.ai / Plus AI / Presentations.ai), and pptx QA/accessibility research
(SlideAudit UIST'25, AeSlides, PPTAgent/PPTEval, openxml-audit).

Validated bets (no change needed): native pptx over image-gen (AutoPresent arXiv 2501.00912;
Tome's 2025 shutdown), markdown outline as source of truth, layout DSL, render-based QA.

Explicitly rejected after research: font *embedding* (Google Slides ignores it; whitelist
linting wins), SmartArt/animation/OMML XML hacks (every serious attempt failed),
python-pptx forks (behind upstream), pptx template engines (abandoned),
matplotlib→native-chart converters (none viable), Gotenberg (wraps LibreOffice, no gain).

Each task below is a self-contained implementation spec. Status: `[ ]` open, `[x]` done.

---

## T2. Chart intelligence pack — `[ ]`

think-cell-parity label/scale behaviors. Files: `scripts/charts.py`, `scripts/build_deck.py`
(syntax), `scripts/builders_consulting.py` (waterfall/funnel/mekko labels),
`references/charts-guide.md`, `tests/test_annotations.py` (extend).

- **Largest-remainder rounding.** New `charts.round_to_sum(values, total=100, decimals=0)`
  implementing largest-remainder so displayed percentages always sum to the stated total.
  Apply wherever the code derives % labels from raw values (mekko column shares, funnel
  conversion rates, bracket % change is a single value — exempt). Unit-test: `[33.33, 33.33,
  33.33] → [33, 33, 34]`, preserves order, handles negatives by passthrough.
- **Value line annotation.** Outline: `- Value-Line: <label>, <value>` inside **Data:**
  context (FIELD_KEYS entry `Value-Line → value_line`). Draws a dashed horizontal reference
  line + small right-aligned label across the plot area for bar/column/line chart visuals,
  reusing the geometry math of `add_benchmark_line` (factor shared helper; `Benchmark:`
  stays as-is for chart-callout). Value must lie within axis range, else build warning.
- **Dual labels on stacked-100.** For `chart:stacked-100` (add to CHART_TYPES if absent as
  XL.COLUMN_STACKED_100): `- Labels: pct|abs|both` field (`labels_mode`). `both` renders
  data labels "42% ($8.4M)" via per-point data_labels text. Percentages via round_to_sum
  per column so each column's labels sum to 100.
- **Same-scale groups.** Deck meta `**Scale-Group:** auto|off` (default off). When `auto`,
  after parsing, group chart slides whose series unit prefix/suffix matches (currency
  symbol or % detected via NUM context in data labels) AND whose chart kind matches;
  compute shared axis max (`_nice_ceil` of global max) and inject as `axis_max` for every
  slide in the group lacking an explicit `- Axis-Max:`. Print info line listing groups.
  Lint side (`pptx_lint.py`): warn when two native charts of the same XL type in one deck
  have value-axis maxima differing by >10x (possible dishonest scale).

Acceptance: new tests pass; `smoke_test.py` green; charts-guide documents all four.

## T3. Stamps, stickers, kicker box, title lint, AI-tell lint — `[x]`

Files: `scripts/build_deck.py`, `scripts/builders.py` (small helpers),
`scripts/pptx_lint.py`, `references/generation-guide.md`, `tests/test_features.py`,
`tests/test_lint.py`.

- **Sticker library.** Per-slide `- Sticker: ILLUSTRATIVE` (also NON-EXHAUSTIVE,
  PRELIMINARY, BACKUP, FOR DISCUSSION — free text, uppercased) renders a small bordered
  tag top-right (mirror of the deck-level `_add_stamp`, which stays top-left). Coexists
  with section kicker text: sticker sits below it if both present.
- **Kicker/takeaway box.** Per-slide `- Kicker: "sentence"` renders a full-width
  accent-bordered box at the bottom content area (above source/footer band) with the
  sentence in bold body size. Rename the existing `_add_kicker` (section tracker) to
  `_add_section_tracker` to free the name. Validation warning when kicker text shares
  >60% of its words with the heading (restatement).
- **Action-title lint extensions** (in `validate()`): warn when heading >15 words; error
  never. Warn when a chart-bearing slide's heading contains no digit. Warn when heading
  contains " and " (split-the-slide signal). All scoped to ACTION_TITLE_LAYOUTS, skipping
  appendix.
- **AI-tell lint** (`pptx_lint.py`): flag (a) thin rectangle (height ≤ 6pt, width ≥ 1in)
  within 0.25in below a title-position text shape — "accent line under title";
  (b) centered body paragraphs of ≥ 2 lines outside title/closing/section/big-number/
  quote layouts (heuristic: PP_ALIGN.CENTER on a text frame taller than 1in below y=1.5in);
  (c) placeholder-text regex extended with `this.*(page|slide).*layout` and `lorem ipsum`.

## T4. New chart layouts: heatmap-table, tornado, football-field — `[x]`

Files: `scripts/builders_consulting.py` (register in LAYOUTS), `scripts/build_deck.py`
(validation), `references/generation-guide.md`, `tests/test_consulting.py`.

- **`heatmap-table`.** Markdown table where numeric body cells get background fills
  interpolated bg→accent1 across the column's min–max (per-column normalization; non-numeric
  cells get surface fill). Text color auto-switches black/white by luminance (reuse
  qa-style contrast math). Optional `- Scale: rag` switches to red/amber/green from
  palette (accent3/accent2-ish mapping; define rag_bad/rag_mid/rag_good fallbacks in
  palette defaults). Validation: needs table_rows with ≥2 numeric columns.
- **`tornado`.** Two-sided horizontal bar chart for sensitivity/comparison. Data rows:
  `- Driver name: -12, +18` (low, high deltas vs a center base). Central spine with
  driver labels, left bars accent2, right bars accent1, value labels at bar ends, sorted
  by total span descending (preserve input order if `- Sort: off`). Built from shapes
  (not native chart). Validation: ≥2 rows, each 2 numeric values.
- **`football-field`.** Valuation-range floating bars. Rows: `- Method: low, high`
  (e.g. `- DCF: 42, 58`). Optional `- Marker: label, value` vertical dashed line spanning
  the chart (e.g. current price). Shape-built horizontal floating bars on a light value
  axis with gridlines and min/max labels. Validation: ≥2 rows, each 2 numerics,
  low < high; marker within global range.

All three: ghost-mode placeholder support comes free via build_ghost_slide; add each to
ACTION_TITLE_LAYOUTS.

## T5. New diagram layouts: driver-tree, stakeholder-map, raci — `[x]`

Files: same as T4.

- **`driver-tree`.** Left-to-right value decomposition. Rows:
  `- Node: Id="rev" Label="Revenue" Value="$120M" Delta="+8%" Parent=""` /
  `Parent="rev"`. Layered tree layout (root left, children right), max depth 3, boxes
  with label+value+optional delta (delta colored green/red by sign), elbow connectors.
  Validation: exactly one root, parents must exist, depth ≤ 3, ≤ 12 nodes.
- **`stakeholder-map`.** Influence(Y) × support(X) grid reusing the matrix-2x2 engine:
  `- Item: Name=".." X="0-1" Y="0-1" TargetX=".." TargetY=".."` — when Target present,
  draws a thin arrow current→target. Axis labels default "Support →" / "Influence →",
  overridable via existing X-axis/Y-axis fields.
- **`raci`.** Markdown table: header `| Activity | Alice | Bob |`, body cells containing
  only letters from RACI (or blank). Renders colored letter chips (R=accent1, A=accent2,
  C=accent3, I=muted). Validation *warning* when a row has zero or 2+ `A`s ("exactly one
  Accountable per activity").

## T6. Section tracker tabs — `[x]`

Files: `scripts/build_deck.py`, `scripts/builders.py`, `references/generation-guide.md`,
`tests/test_build_deck.py`.

- Deck meta `**Tracker:** tabs` (default off; requires section dividers). On every content
  slide after the first divider, replaces the plain-text section kicker with a compact
  tab strip top-right: one small rounded chip per section, current section chip filled
  accent1 with bg-colored text, others outlined muted. Chips truncate section names to
  ~14 chars; ≥6 sections → fall back to "n/N · Section" text form (warn once).
  Mutually compatible with `**Auto-Agenda:**`.

## T7. QA pack: accessibility, Google Slides compat, geometry report — `[ ]`

Files: `scripts/qa_check.py`, `scripts/pptx_lint.py`, new `scripts/geometry_report.py`,
`references/qa-guide.md`, `tests/` (extend test_lint.py, new test_geometry.py).

- **Accessibility extensions** (`qa_check.py --accessibility`, mirroring MS Accessibility
  Checker): every slide has a title-ish text (existing `_slide_title_text`); duplicate
  slide titles (case-folded) flagged; tables missing `firstRow` header marking; tables
  containing merged cells (gridSpan/rowSpan/hMerge/vMerge) → warn "simple structure";
  alt text equal to a filename pattern (`.*\.(png|jpe?g|gif)$` or `image\d+`) flagged
  even when present; reading-order heuristic — warn when the title shape is not first
  in spTree document order.
- **Google Slides compatibility lint** (`pptx_lint.py --gslides`): fonts outside a
  bundled Google-fonts-available whitelist (Arial, Georgia, Times New Roman, Verdana,
  Trebuchet MS, Courier New, + Google-library names list constant); presence of `dgm:`
  (SmartArt) namespaces; `p:transition` beyond fade/none; embedded media parts.
- **Geometry report** (`scripts/geometry_report.py deck.pptx [--json]`): per-slide
  deterministic layout metrics for LLM self-verification *before* any render —
  (a) pairwise overlap of visible shapes (area + ids), excluding intentional
  background/card containment (shape fully inside another = containment, not overlap);
  (b) column/row gap consistency: cluster shapes by row/column, report gap variance;
  (c) edge-alignment deviations < 0.08in ("almost aligned" jiggle within a slide);
  (d) whitespace ratio and left/right + top/bottom visual-mass imbalance (AeSlides
  metrics); (e) text density (words per slide). Output human-readable by default,
  `--json` for tooling. Document in qa-guide as QA step 1.5 (run before rendering;
  feed output to the fix loop).

## T8. Build-time overflow guard (capacity budgets + autofit) — `[ ]`

Files: `scripts/builders.py` (or new `scripts/textfit.py`), `scripts/build_deck.py`
(validation), `tests/test_textfit.py`.

- **Text measurement.** `textfit.estimate_lines(text, font_pt, box_w_in, bold=False)`
  using PIL `ImageFont` when a matching TTF is found (font_manager lookup), else a
  per-character width table fallback (average widths for common sans faces). Must be
  deterministic and dependency-tolerant (PIL already required by prep_images? verify —
  if not, fallback table only).
- **Capacity budgets in `validate()`**: per-layout char/line budgets (title ≤ 2 lines at
  its render size; bullet block total estimated lines ≤ layout capacity; card/tile body
  fields ≤ budget). Emit *errors* only for catastrophic overflow (>140% of capacity),
  warnings between 100–140%.
- **normAutofit fallback.** For bullet/two-column text frames between 100–140% capacity,
  write `<a:normAutofit fontScale="..%" lnSpcReduction="10%"/>` into the body's bodyPr
  computing fontScale to fit, floor 80% (i.e. 16pt→12.8pt min). Below floor → keep the
  validation warning ("split the slide").

## T9. OOXML wraps: image transparency/duotone, sections, table banding — `[ ]`

Files: `scripts/helpers.py`, `scripts/builders.py` (table builder), new logic in
`scripts/build_deck.py` (sections), `references/generation-guide.md`,
`tests/test_features.py`.

- **Image alpha + duotone.** Visual spec extensions: `image:photo.png|alpha=85` and
  `image:photo.png|duotone` parsed by `parse_visual` (pipe-separated options dict as
  third return; keep 2-tuple callers working via wrapper). alpha → `a:alphaModFix`
  appended to the picture's `a:blip`; duotone → `a:duotone` from palette bg→accent1
  (brand-tinted photos). Both verified to render in LibreOffice (add to smoke outline).
- **Slide sections (`p14:sectionLst`).** After build, when the deck has section-divider
  slides, inject PowerPoint sections named after each divider (first section "Opening"
  for pre-divider slides; appendix slides → "Backup" section). Deterministic GUIDs
  derived from md5(section name + index) so rebuilds are stable. Behind meta
  `**Sections:** on` (default on when ≥2 dividers; `off` disables).
- **Native table banding.** In the table builder, set `firstRow=True` and
  `horz_banding=True` on the `a:tbl` properties so accessibility checkers see a marked
  header row (visual styling stays custom).

## T10. Pipeline outputs: references slide, handout, CSV data, CJK stacks — `[ ]`

Files: `scripts/build_deck.py`, new `scripts/gen_handout.py`, `scripts/palettes.py`
(CJK), `references/generation-guide.md`, `SKILL.md` (Phase 1 doc-to-deck note),
`tests/test_build_deck.py`.

- **Auto-references appendix.** Meta `**References:** on` → appends a backup slide
  "Sources" (bullet layout, appendix-flagged) aggregating unique `- Source:` values
  with their exhibit numbers / slide numbers.
- **Handout generator.** `python3 scripts/gen_handout.py outline.md --output handout.md`:
  emits a readable pre-read doc — per slide: heading as H2, bullets/body as prose
  bullets, data tables as markdown tables, speaker notes as a "Talk track" blockquote,
  sources as footnotes. Pure markdown transformation, no pptx dependency.
- **CSV chart data.** Slide field `- Data-File: data/q3.csv` (resolved like images
  against outline dir / assets dir): two-column `label,value` or multi-column with
  header row mapping to **Series:** names. Loaded into the same `data` structure at
  parse time; file-missing → validation error.
- **CJK font stacks.** When any slide text contains CJK codepoints, swap per-palette
  fonts to a CJK-safe stack (PingFang SC on macOS / Noto Sans CJK SC elsewhere — pick
  by `fc-list` availability, warn when neither found) for those runs' fonts at minimum
  deck-wide font fallback. Implement as palette post-processor in `get_palette` path
  guarded by a `set_cjk(True)` toggle decided in `build()` after parsing.

## T11. Brand kit ingestion — `[ ]`

New `scripts/brand_kit.py`, `references/generation-guide.md` (palette section),
`tests/test_brand_kit.py` (offline: HTML/CSS fixtures only).

- `python3 scripts/brand_kit.py <domain-or-url> --name acme [--assets-dir assets]`.
- Source 1: Brandfetch API when `BRANDFETCH_API_KEY` env is set (colors, fonts, logo URL).
- Source 2 (fallback, no key): fetch homepage HTML + linked CSS (stdlib urllib, 10s
  timeout, useragent set); regex hex/rgb() colors; dedupe near-identical colors
  (Euclidean RGB distance < 30); rank by frequency; classify into bg/surface/accents by
  luminance + saturation; derive text colors by contrast (reuse palettes/qa math).
- Output: `<assets>/palettes/<name>.json` with all 9 REQUIRED_KEYS + `dark` flag,
  validated through `load_custom_palettes` before writing final file; prints a swatch
  summary table (hex + role + contrast ratios) and downloads logo to
  `<assets>/brand/<name>-logo.png` when available.
- Network failures → clear error, no partial files. Tests run entirely on local fixture
  HTML/CSS (no network).

## T12. Deck integrity + visual regression — `[ ]`

New `scripts/visual_regress.py`; `scripts/qa_check.py` (integrity hook);
`requirements.txt` comments; `references/qa-guide.md`; `tests/test_regress.py`.

- **openxml-audit gate.** `qa_check.py` gains `--integrity`: if `openxml-audit`
  importable, validate the package and report errors as QA issues; if not installed,
  print one-line "pip install openxml-audit for OOXML schema validation" note (not a
  failure). Wire into `smoke_test.py` as advisory.
- **Visual regression.** `visual_regress.py baseline/ current/ [--update] [--threshold N]`
  comparing same-named slide PNGs: pHash hamming distance via `imagehash` when
  installed, else Pillow-based average-hash fallback implemented locally (so zero new
  hard deps); flagged slides get a pixel-diff percentage; `--update` blesses current as
  baseline. Exit 1 when any slide exceeds threshold. Document workflow: render → regress
  vs last delivered deck before sending revisions.

## T13. Edit mode: verbs, remix, deck merge/split — `[ ]`

Files: `scripts/edit_deck.py`, `references/editing.md`, `tests/test_edit_deck.py`.

- **Edit verbs documentation** (`editing.md`): a "Slide operations playbook" section
  defining condense / rewrite / formalize / translate / split / remix as documented
  recipes over existing primitives (inventory → targeted replace; remix = re-author that
  slide's outline section with a different layout and rebuild single slide via add_slide
  flow). No new code beyond what exists; precise step lists.
- **`extract` subcommand** (split): `edit_deck.py extract deck.pptx 3-7 --output sub.pptx`
  — copies the package, removes all slides outside the range (reusing `remove` + clean),
  repacks. Pure reuse of existing machinery.
- **`append` subcommand** (merge): `edit_deck.py append dst.pptx src.pptx [--slides 2-4]
  --output merged.pptx` — copies selected slide parts from src package into dst:
  slide XML + its layout + master + theme + media/charts it references, renaming parts
  to avoid collisions, remapping all rIds, registering content types, appending to
  sldIdLst. Validate XML before repack (existing validator). Known-limitation note:
  source slides keep their own master/theme (faithful import, no restyling).

## T14. Docs: VLM rubric QA, SKILL.md/README refresh — `[ ]`

Files: `references/qa-guide.md`, `SKILL.md`, `README.md`, `references/decision-tree.md`.

- **Taxonomy-guided VLM critique** (qa-guide): replace the open-ended fresh-eyes prompt
  with a binary per-flaw checklist (SlideAudit finding: checklists beat holistic scores):
  overlap/collision, misalignment, crowding/whitespace imbalance, weak title hierarchy,
  off-palette color, text overload, stretched/pixelated imagery, inconsistency with
  sibling slides — each answered per numbered grid cell, plus a 1–5 Content/Design/
  Coherence rubric (PPTEval) recorded in the QA summary; score <4 routes to fix loop.
- **SKILL.md**: Phase 1 gains doc-to-deck note ("user may supply PDF/docx/xlsx/CSV —
  draft the outline from it"); Phase 3/4 mention geometry_report, --integrity, handout,
  brand_kit, tracker tabs, new layouts count; Phase 4 adds visual_regress for revision
  cycles. README feature list + layout table updated; ROADMAP statuses flipped.
- **Optional renderers note** (qa-guide): ONLYOFFICE x2t and MS Graph PDF export as
  higher-fidelity alternates worth using when available; LibreOffice remains default.

---

## Deferred (researched, intentionally not in this cycle)

- **Template-mode slide clustering / clone-from-finished-deck** (PPTAgent Stage I) —
  high effort; revisit after template-mode usage feedback.
- **Label collision avoidance engine** (think-cell-grade) — needs per-label bounding
  geometry across all chart builders; do after T2 lands and real decks expose cases.
- **Stacked-segment waterfalls with auto subtotals** — extend waterfall once demand
  appears; bracket + CAGR already cover most bridge stories.
- **ONLYOFFICE second renderer spike** — documented as optional in qa-guide (T14);
  full integration deferred.
- **Axis breaks ("breaking bars")** — recognizable but fiddly; revisit with shape-chart
  unification.
