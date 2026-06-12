# Generation Guide

## Core Build Stack

| Tool | Purpose |
|------|---------|
| `python-pptx` | Native .pptx creation — real shapes, placeholders, charts |
| `Pillow` | Image normalization and aspect-ratio math |
| `scripts/build_deck.py` | Builds the full deck from a markdown outline (also `--check` validation) |
| `scripts/qa_check.py` | Programmatic QA of the built .pptx |
| `scripts/render_slides.py` | Converts .pptx → thumbnail images for visual QA |

`python-pptx` produces native PowerPoint XML with proper placeholders and fully editable elements. Always use the scripts rather than hand-writing pptx code — the layout builders live in `scripts/builders.py`, palettes in `scripts/palettes.py`, chart helpers in `scripts/charts.py`.

---

## Outline Format (exact syntax)

One `## Slide N:` section per slide. The text after `Slide N:` becomes the heading unless overridden with `- Heading:` or `- Title:` (recommended on `title` / `closing` layouts; if omitted, the `## Slide N:` text is used as the title). Always validate with `build_deck.py --check` before building.

**Deck front-matter** (optional, before the first `## Slide`):

```markdown
**Palette:** midnight-executive        ← default for the whole deck
**Purpose:** pitch                     ← pitch|strategy|client|talk|tutorial (narrative hints)
**Takeaway:** "One sentence investors should remember"
**Variant:** b                         ← preset: a=default, b=aurora+comfortable, c=swiss-light+compact
**Footer:** "Confidential · Q3 2026"   ← bottom-left on content slides
**Page-Numbers:** on                   ← bottom-right on content slides
**Size:** 16:9                         ← or 4:3
**Density:** compact                   ← compact (default) or comfortable
```

Omit `**Layout:**` on a slide and the builder auto-selects from content (`smart_layout.py`). Explicit layout always wins. See `references/narratives/*.md` for purpose-specific slide arcs.

**Inline rich text** works in bullets, points, captions, and card bodies:
`**bold**` for bold, `{accent}text{/}` for accent-colored emphasis.

**Icons**: prefix any bullet with `icon:<name> ` (kebab-case [Tabler](https://tabler.io/icons)/[Lucide](https://lucide.dev) names) to render an icon-in-circle marker. Cards take `Icon="name"`. Icons are fetched once and cached in `assets/icons/`.

```markdown
## Slide 1: Title
**Layout:** title
**Visual:** user-image:hero.png        ← optional hero image, right side
**Palette:** swiss-light               ← optional per-slide override
- Title: "Accelerating AI Infrastructure"
- Subtitle: "Q4 Strategy Review · December 2026"
- Notes: "Open with the cost-savings story."   ← speaker notes (write these for EVERY slide)

## Slide 2: Executive Summary
**Layout:** exec-summary
- Headline: "Three strategic priorities for H1 2027"
- Point 1: "icon:server Deploy GPU cluster Phase 2 by March (saving **$2.1M/yr**)"
- Point 2: "icon:cloud Migrate {accent}80% of workloads{/} to hybrid cloud by April"
- Point 3: "icon:rocket Launch developer platform beta with 500 design partners"

## Slide 3: Market Opportunity
**Layout:** two-column-split
**Visual:** chart:bar                  ← bar|hbar|line|pie|doughnut|area|scatter
**Series:** Market Size ($B)
**Data:**
- 2024: 42
- 2025: 67
- 2026: 104
- 2027: 158
- Heading: "AI infrastructure market growing 32% YoY"
- Enterprise AI spend reached $104B globally in 2026
- Infrastructure layer captures 41% of total AI spend

## Slide 3b: Revenue vs Costs (multi-series)
**Layout:** two-column-split
**Visual:** chart:line
**Series:** Revenue, Costs
**Data:**
- Q1: 4.2, 3.1
- Q2: 5.1, 3.0
- Q3: 6.3, 2.8
- Heading: "Margin expansion through cloud migration"

## Slide 4: Product Screenshot
**Layout:** full-image
**Visual:** user-image:product-screenshot.png
- Caption: "Live dashboard — 47ms p99 latency"

## Slide 5: Team
**Layout:** cards-3
- Card 1: Name="Alex Kim" Title="CEO & Co-founder" Bio="Ex-Google Brain, 2x founder"
- Card 2: Name="Jordan Lee" Title="CTO" Bio="Built infra at Scale AI"
- Card 3: Name="Sam Patel" Title="Head of Product" Bio="Former PM at Stripe"

## Slide 6: Key Metrics
**Layout:** stat-callout
- Value="$14.2M" Label="Cost Savings" Sublabel="vs. legacy stack"
- Value="99.97%" Label="Uptime SLA" Sublabel="Jan–Jun 2026"

## Slide 7: Roadmap
**Layout:** timeline
- Date="Q3 2026" Title="Cloud migration" Desc="Complete hybrid migration"
- Date="Q4 2026" Title="Phase 2 GPUs" Desc="256 H100s operational"
- Date="Q1 2027" Title="Platform GA" Desc="Billing + RBAC"

## Slide 8: Before / After
**Layout:** comparison
**Visual-Left:** user-image:before.png      ← image OR bullets per side
**Visual-Right:** user-image:after.png
- Left label: "Legacy Stack (2023)"
- Right label: "New Architecture (2026)"
- Left: "Monolith on-prem"                  ← used when no Visual-Left image
- Right: "Active-active multi-region"

## Slide 9: Pricing
**Layout:** table
- Heading: "Plan comparison"
| Plan | Price | Seats |
|------|-------|-------|
| Starter | $99 | 5 |
| Scale | $499 | 50 |
```

Generic bullets (`- text`) work on `bullet-list`, `two-column-split`, `exec-summary` (or use `- Point N:`). Timeline also accepts plain `- Q3 2026: description` bullets.

---

## Layout Types (all implemented in `scripts/builders.py`)

| Layout Name | When to Use | Content keys |
|-------------|-------------|--------------|
| `title` | Opening slide | Title, Subtitle, optional Visual |
| `section-divider` | Chapter breaks | heading (from `## Slide N:`), Subtitle |
| `exec-summary` | Key takeaways | Headline + up to 4 Point N cards |
| `bullet-list` | Main points | Heading + up to 14 bullets (2 columns; 7 in comfortable) |
| `two-column-split` | Text + visual | Bullets left; chart/image right |
| `cards-3` / `cards-4` | Team, features | Card N with Name/Title/Bio (or Title/Body) |
| `stat-callout` | Big numbers | 2–3 Value/Label/Sublabel stats |
| `timeline` | Chronological steps | Date/Title/Desc items (max 5) |
| `comparison` | Before/after, A vs B | Left/Right labels + images or bullets |
| `table` | Structured data | Markdown table rows |
| `full-image` | Hero visual | Visual (cover-cropped) + Caption |
| `closing` | Thank you / CTA | Title, Subtitle, Contact |
| `exec-summary-scqa` | Consulting exec summary | Situation, Finding N, Recommendation |
| `agenda` | Section navigation | Section names as bullets, optional Current |
| `waterfall` | Value bridges, cost walks | **Data:** start, ± deltas, `total` row |
| `matrix-2x2` | Prioritization quadrants | X-axis/Y-axis, Q1–Q4, Item rows |
| `harvey-scorecard` | Vendor/capability comparison | Markdown table, cells 0–4 |
| `process-flow` | Operating model steps | Step N: Title/Desc (chevrons, max 6) |
| `big-number` | One hero metric | Value, Label, Context |
| `chart-callout` | Chart + key insight box | chart Data + Callout |
| `dashboard` | KPI tiles with trends | Tile N: Value/Label/Delta/Trend |
| `quote-evidence` | Voice + proof | Quote, Attribution, one Value/Label stat |
| `funnel` | Pipeline drop-off | **Data:** stage rows (auto conversion %) |
| `next-steps` | Conclusions for Q&A | Step N: Action/Owner/When |
| `mekko` | Market maps, share-of-wallet | **Series:** segments + multi-value **Data:** |
| `gantt` | Phased roadmaps, swimlanes | **Periods:** + Bar/Milestone rows |
| `heatmap-table` | Benchmarks, scorecards by intensity | Markdown table, numeric cells heat-filled; optional `- Scale: rag` |
| `tornado` | Sensitivity analysis | **Series:** Low, High + `- Driver: low, high` rows; optional `- Sort: off` |
| `football-field` | Valuation ranges | **Series:** Low, High + `- Method: low, high` rows; optional `- Marker:` |

### Consulting layout syntax

```markdown
## Slide 2: We recommend consolidating onto two strategic cloud platforms
**Layout:** exec-summary-scqa
- Situation: "Cloud spend doubled to $46M while delivery velocity stalled"
- Finding 1: "61% of spend sits on workloads with <30% utilization"
- Finding 2: "Tooling fragmentation adds 11 days to median release cycle"
- Recommendation: "Consolidate to two platforms; tier workloads — $18M savings by FY27"

## Slide 5: Three levers bridge EBITDA from $42M to $58M
**Layout:** waterfall
**Data:**
- FY24: 42
- Pricing: +9
- Volume: +6
- Cost inflation: -7
- Efficiency: +8
- FY25: total
- Source: "Company financials; team analysis"

## Slide 6: Workload tiering is the fastest path to value
**Layout:** matrix-2x2
- X-axis: "Implementation effort →"
- Y-axis: "Value at stake →"
- Q1: "Quick wins"
- Item: Name="Workload tiering" X="0.2" Y="0.85"
- Item: Name="Platform exit" X="0.85" Y="0.7"

## Slide 8: Vendor B leads on the criteria that matter most
**Layout:** harvey-scorecard
| Criterion | Vendor A | Vendor B |
|---|---|---|
| Scalability | 2 | 4 |
| Ecosystem | 3 | 3 |

## Slide 9: Spend passes budget in 2026 without intervention
**Layout:** chart-callout
**Visual:** chart:line
**Benchmark:** 60 "FY budget ceiling"
**Data:**
- 2024: 38
- 2025: 52
- 2026: 71
- Callout: "Crossover hits in Q2 2026 — decision needed this quarter"

## Slide 12: Decision required: approve Phase 1 by month end
**Layout:** next-steps
- Step 1: Action="Approve Phase 1 budget ($2.4M)" Owner="CFO" When="This week"
- Step 2: Action="Stand up tiger team" Owner="CTO" When="Within 2 weeks"

## Appendix
## Slide 13: Detailed cost model assumptions
**Layout:** table
...backup slides: marked BACKUP, numbered B·n, exempt from variety/notes warnings
```

```markdown
## Slide 10: Two platforms capture 73% of segment value
**Layout:** mekko
**Series:** Platform A, Platform B, Other
**Data:**
- Compute: 18, 22, 6        ← column width ∝ total, heights = share
- Storage: 9, 11, 4
- Networking: 5, 4, 3

## Slide 11: Three workstreams deliver over six quarters
**Layout:** gantt
**Periods:** Q3 26, Q4 26, Q1 27, Q2 27
- Bar: Row="Infrastructure" Label="Tiering" Start="1" End="2"
- Bar: Row="Infrastructure" Label="Exit" Start="3" End="4"
- Bar: Row="Platform" Label="Pilot" Start="1" End="1"
- Milestone: Row="Platform" Label="GA" At="3"
```

```markdown
## Slide 12: EMEA unit costs run 3x the NA baseline on every metric
**Layout:** heatmap-table
| Region | Cost | Churn % | NPS |
|---|---|---|---|
| NA | $10 | 5% | 62 |
| EMEA | $30 | 7% | 41 |
| APAC | $20 | 6% | 55 |
- Scale: rag                ← optional: red/amber/green terciles instead of
                              the default bg→accent1 heat per column
- Notes: "Heat is normalized per column; text/$/% cells parse as numbers."

## Slide 13: Margin and WACC swings drive a 35-point valuation range
**Layout:** tornado
**Series:** Downside, Upside     ← REQUIRED: exactly 2 names so the rows
**Data:**                          parse as [low, high] pairs
- WACC +/-1pt: -12, +18
- Volume growth: -8, +10
- Gross margin: -15, +20
- Sort: off                 ← optional: keep input order (default sorts by
                              |low|+|high| descending)
- Notes: "Left bars = values[0] (accent2), right = values[1] (accent1)."

## Slide 14: Three valuation methods converge on the $45-52 band
**Layout:** football-field
**Series:** Low, High            ← REQUIRED: 2 names → [low, high] pairs
**Data:**
- DCF: 42, 58
- Trading comps: 45, 52
- Precedent deals: 40, 50
- Marker: Current price, 47 ← optional dashed vertical line + label; must
                              fall inside the data range or it is skipped
- Notes: "Floating bars low→high on a shared nice-interval axis."
```

`heatmap-table` needs a header row plus body rows with 2+ columns and at
least one numeric body column (`--check` errors otherwise); non-numeric
cells get a plain surface fill and cell text auto-switches black/white by
fill luminance. `tornado` and `football-field` reuse the multi-series data
parser — declare `**Series:**` with exactly two names or the rows won't
parse as pairs (`--check` errors). `football-field` additionally requires
low < high per row. Put `- Scale:` / `- Sort:` / `- Marker:` lines *after*
the data rows — like any `- Key:` field they end the **Data:** block.

`**Benchmark:** <value> "label"` adds a dashed reference line to bar/column/line/area charts. `**Exhibits:** on` (front-matter) numbers every sourced slide's footnote "Exhibit N · Source: ...". `agenda` slides highlight the row matching `- Current: "Section name"`. Section dividers automatically set the top-right section tracker label on subsequent content slides (`- Kicker:` is the bottom takeaway band — a different element).

---

## Palette Selection Guide

Palettes (colors + per-theme font stacks, 60-30-10 dominance encoded) are defined once in `scripts/palettes.py`.

| Palette Key | Best For | Theme | Title font |
|-------------|----------|-------|------------|
| `midnight-executive` | Finance, strategy, luxury | Dark | Gill Sans MT |
| `aurora` | AI, tech, innovation | Dark | Gill Sans MT |
| `venture-pitch` | Startups, product launches | Dark | Trebuchet MS |
| `forest` | Sustainability, ESG | Dark | Gill Sans MT |
| `teal-trust` | Healthcare, fintech, services | Dark | Gill Sans MT |
| `charcoal-minimal` | Design, architecture, monochrome | Dark | Arial Black |
| `ocean-gradient` | Data, cloud, deep tech | Dark | Gill Sans MT |
| `swiss-light` | Corporate, education, tutorials | Light | Palatino Linotype |
| `warm-terracotta` | Consumer, hospitality, editorial | Light | Georgia |
| `berry-cream` | Lifestyle, brand, creative | Light | Georgia |

Pick by topic domain, not habit — defaulting to blue is an AI tell. If the user does not specify, use `midnight-executive` for business decks and `aurora` for tech/AI. Light palettes intentionally avoid pure `#FFFFFF` backgrounds.

---

## Design Rules (enforced or checked)

- **Layout variety**: never use the same layout 3+ times in a row (`--check` warns). Alternate text layouts with cards, stats, charts, images.
- **No accent lines under titles** — a hallmark of AI-generated decks. Hierarchy comes from size contrast and whitespace (builders already comply).
- **Every content slide gets a visual**: icon bullets, a chart, an image, or stat cards. `qa_check.py` warns on text-only slides.
- **Sandwich structure**: title/section/closing slides render with a dark gradient for a premium feel; for a light-content deck, set a light deck palette and per-slide override dark on slide 1 and the closer.
- **Left-align body text** (builders comply); center only titles and stat values.
- **Density**: default is `compact` (double-density) — up to 14 bullets flowing into two columns, 8 exec-summary cards in a 2-column grid, 8 points beside a chart/image, 9 comparison rows (~220-word qa warning). `**Density:** comfortable` gives single-column spacing (7 bullets, 5 exec/two-column points). Prefer filling a slide over splitting content across sparse ones.
- **Speaker notes on every slide** — write `- Notes:` lines in the outline; `--check` warns when missing.

**Data blocks:** blank lines inside `**Data:**` are ignored. Each data row must be `- label: value` (multi-series: `- label: v1, v2`).

---

## Slide Dimensions & Typography

Widescreen **13.33" × 7.5"** by default; `**Size:** 4:3` or `--size 4:3` gives 10" × 7.5" (builders rescale automatically).

Body/bullets render at 16–17pt, captions/labels 10–13pt, titles 30–54pt in the palette's title font. Never go below 10pt (enforced by `qa_check.py`). Fonts are web-safe PowerPoint defaults; PowerPoint substitutes automatically when a font is missing (python-pptx cannot embed fonts).

---

## Charts

Declare `**Visual:** chart:<type>` plus a `**Data:**` block (see Slide 3 above) and the build pipeline generates a native, palette-styled chart with a transparent background — fully editable in PowerPoint. For custom chart code or shape-based visuals (progress bars, stat cards), see `references/charts-guide.md`.

## Newer directives (June 2026)

**Heading attributes** — set layout/palette on the heading line, keeping
directives off content bullets:

    ## Slide 7: Margin bridge tells the story {layout=waterfall palette=aurora}

Values must be unquoted (`layout=two-column-split`, not `layout="..."`);
unknown keys emit a warning. An explicit `**Layout:**` line still wins if
both are present.

**Deck meta:**
- `**Auto-Agenda:** on` — auto-insert an agenda slide (sections = your
  section-divider headings, appendix excluded) after the title slide.
- `**Auto-Agenda:** track` — additionally insert a "where we are" agenda with
  the current section highlighted after every section divider.
- `**Stamp:** DRAFT` — bordered status tag on every slide (also:
  CONFIDENTIAL, FOR DISCUSSION).

**Per-slide tags:**
- `- Sticker: ILLUSTRATIVE` — small bordered tag top-right (free text,
  uppercased; e.g. NON-EXHAUSTIVE, PRELIMINARY, BACKUP, FOR DISCUSSION).
  Visual mirror of the deck-level `**Stamp:**`, which stays top-left.
  Coexists with the section label: the sticker drops below it when both
  are present.
- `- Kicker: "So-what sentence"` — bold takeaway band across the bottom of
  the content area, just above the source/footer. Make it advance the
  argument — `--check` warns when it merely restates the slide title.
  Skipped on title/closing/section-divider/full-image layouts.

**Action-title lint (`--check`, warnings):** on action-title layouts,
headings over 15 words, exhibit headings (chart visuals, waterfall, mekko,
bar-mekko, chart-callout) with no number, and headings joining two messages
with " and " are flagged.

**AI-tell lint (`pptx_lint.py`):** thin accent lines under titles and long
centered body copy — both hallmarks of AI-generated decks — are warned on
any .pptx, including imported ones.

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
  (width ∝ Size, height = Value; Size>0, Value>=0). 2+ bars required.
- `matrix-2x2` bubbles — add `Size="40"` to `- Item:` rows; bubble area
  scales with Size (BCG growth–share style).
- `heatmap-table` — markdown table with per-column heat fills (bg→accent1);
  `- Scale: rag` switches to red/amber/green terciles (palette keys
  `rag_bad`/`rag_mid`/`rag_good`, overridable in custom palettes).
- `tornado` — sensitivity bars off a central spine; needs `**Series:**` with
  2 names + `- Driver: low, high` rows; `- Sort: off` keeps input order.
- `football-field` — valuation range bars on a shared axis; needs
  `**Series:**` with 2 names + `- Method: low, high` rows (low < high);
  `- Marker: label, value` adds a dashed reference line.

**Custom palettes:** drop `<name>.json` into `<assets>/palettes/`:

    {"bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
     "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
     "text": "F4F4F4", "text_muted": "9DB2BF", "dark": true}

All nine keys required; colors must be 6-char hex. Optional keys:
`font_title`, `font_body`, `font_label`, `chart_series` (list of hex —
chart series colors), `motif`. Then `--palette <name>`.
