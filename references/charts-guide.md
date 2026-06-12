# Charts Guide

## Preferred Path: Outline-Driven Charts

You normally don't write chart code. Declare the chart in the outline and `build_deck.py` generates it via `scripts/charts.py`:

```markdown
**Visual:** chart:bar        ← bar | hbar | line | pie | doughnut | area | scatter | stacked-100
**Series:** Market Size ($B)
**Data:**
- 2024: 42
- 2025: 67
- 2026: 104
```

The pipeline automatically applies palette colors, axis/tick label styling, and — critically — transparent chart and plot areas.

### Multi-series (outline)

Comma-separate series names and values per row:

```markdown
**Visual:** chart:line
**Series:** Revenue, Costs
**Data:**
- Q1: 4.2, 3.1
- Q2: 5.1, 3.0
- Q3: 6.3, 2.8
```

`--check` validates that every data row has the same number of comma-separated values as series names.

## The White-Box Defect (and the fix)

Default python-pptx charts render with an **opaque white chart area** — the most common visual defect in AI-generated decks on dark slides. python-pptx has no public API for the chart-area fill, so `scripts/charts.py` clears it at the XML level:

```python
from charts import make_chart_transparent, add_native_chart

# add_native_chart() already calls this. If you hand-roll a chart, you MUST:
make_chart_transparent(chart)   # inserts <c:spPr><a:noFill/> on chartSpace + plotArea
```

`scripts/qa_check.py` flags any chart that is missing this.

## Custom Charts

For scatter plots or chart placement the outline can't express, call the helper directly:

```python
from charts import add_native_chart

add_native_chart(slide, pal, "line",
                 categories=["Q1", "Q2", "Q3"],
                 series={"Revenue": (4.2, 5.1, 6.3), "Costs": (3.1, 3.0, 2.8)},
                 left=7.0, top=1.5, w=5.5, h=4.5)
```

Multi-series gets a legend automatically; pie/doughnut slices cycle through the palette accents.

---

## Shape-Based Visuals (better than charts for stats)

For stat callouts, progress bars, and KPI displays, custom shapes look more polished than native chart objects. The `stat-callout` layout covers cards; for progress bars:

```python
from helpers import add_rect, add_tb

def add_progress_bar(slide, pal, label, percent, value_label, left, top,
                     width=9.0, height=0.45):
    add_tb(slide, label, left, top, 3.5, height,
           size=14, color=pal["text"], font=pal["font_body"])
    track_w = width - 4.0
    add_rect(slide, left + 3.7, top + 0.05, track_w, height - 0.1, pal["surface"])
    add_rect(slide, left + 3.7, top + 0.05, track_w * percent / 100,
             height - 0.1, pal["accent1"])
    add_tb(slide, value_label, left + width - 0.25, top, 0.9, height,
           size=13, bold=True, color=pal["accent1"], font=pal["font_body"])
```

## Annotations

- **Benchmark line:** `- Benchmark: 120 Industry average` (existing).
- **Value line:** `- Value-Line: <label>, <value>` — see below.
- **CAGR arrow:** `- CAGR: on` — computed from first/last values of a
  single-series bar/column/line chart. Plot-box placement is approximate:
  always confirm in visual QA.
- **Same scale:** `- Axis-Max: N` pins the value axis (applied last, so it
  wins over the axis max set by Benchmark/CAGR). Use the same N across
  slides being compared — or let `**Scale-Group:** auto` do it (below).
- **Waterfall bracket:** `- Bracket: <bar A>, <bar B>[, "label"]` — exact
  geometry (shape-drawn), no QA caveat.

## Chart Intelligence (think-cell-parity behaviors)

### Largest-remainder rounding

Displayed percentage sets always sum to the stated total: `charts.round_to_sum(
values, total=100, decimals=0)` implements largest-remainder rounding —
`[33.33, 33.33, 33.33]` renders as `33% / 33% / 34%`, never `33+33+33=99`.
Applied automatically to mekko column shares and stacked-100 `pct` labels
(per column, so each column's labels sum to 100). Any negative input falls
back to plain rounding. Single-value % labels (waterfall bracket, funnel
stage-to-stage conversion) are exempt — there is no group to reconcile.

### Value line

```markdown
**Visual:** chart:bar
- Value-Line: Target, 40
**Data:**
- 2023: 30
- 2024: 44
```

Draws a dashed horizontal reference line with a small right-aligned label
across bar/column/line chart visuals. Unlike `- Benchmark:` (which widens the
axis to fit), a value line never changes an already-set axis: if the value
falls outside the computed axis range, the build prints a warning and skips
the line. Place it after `- Axis-Max:` decisions — it is drawn against the
final axis range.

### Dual labels on 100%-stacked columns

```markdown
**Visual:** chart:stacked-100
**Series:** Subs, Services, Hardware
- Labels: both                ← pct | abs | both
**Data:**
- 2023: 1, 1, 1
- 2024: 8.4, 7.6, 4
```

`pct` renders per-segment percentages (largest-remainder per column, so each
column shows exactly 100), `abs` renders the raw values, `both` renders
`42% (8.4)`. Without `- Labels:` the chart renders unlabeled (default
behavior unchanged).

### Same-scale groups

```markdown
**Scale-Group:** auto         ← deck front-matter; default off
```

When `auto`, slides whose chart visuals share a kind (bar/column/hbar/line/
area), have numeric data, and have no explicit `- Axis-Max:` are grouped per
kind; each group of 2+ slides gets one shared axis maximum (`_nice_ceil` of
the group's global data max), printed at build time:

```
  Scale-group: column charts on slides 4, 7 share axis max 50
```

The reverse check lives in `pptx_lint.py`: same-type native charts whose
explicit value-axis maxima differ by more than 10x are flagged as a
"possible dishonest scale comparison" warning.
