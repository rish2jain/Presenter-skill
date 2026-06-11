# Charts Guide

## Preferred Path: Outline-Driven Charts

You normally don't write chart code. Declare the chart in the outline and `build_deck.py` generates it via `scripts/charts.py`:

```markdown
**Visual:** chart:bar        ← bar | hbar | line | pie | doughnut | area | scatter
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
