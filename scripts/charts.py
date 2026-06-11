"""Native python-pptx chart helpers, styled for dark/light palette themes.

The critical detail: default charts render with an opaque white chart/plot
area, which looks broken on dark slides. make_chart_transparent() clears both.
"""
from lxml import etree
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn

from palettes import hex_rgb

CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}

_NO_FILL_SPPR = (
    '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    "<a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>"
)


def _clear_fill(element, insert_after=None):
    """Ensure element has a <c:spPr> with noFill (replacing any existing fill)."""
    spPr = element.find(qn("c:spPr"))
    if spPr is not None:
        element.remove(spPr)
    spPr = etree.fromstring(_NO_FILL_SPPR)
    if insert_after is not None:
        insert_after.addnext(spPr)
    else:
        element.append(spPr)
    return spPr


def make_chart_transparent(chart):
    """Clear the chart-area and plot-area backgrounds (white-box fix)."""
    chart_space = chart._chartSpace
    chart_el = chart_space.find(qn("c:chart"))
    # chartSpace schema: spPr must directly follow <c:chart>
    _clear_fill(chart_space, insert_after=chart_el)
    plot_area = chart_el.find(qn("c:plotArea"))
    if plot_area is not None:
        _clear_fill(plot_area)


def _style_axes(chart, pal):
    chart.font.size = Pt(12)
    chart.font.color.rgb = hex_rgb(pal["text_muted"])
    chart.font.name = pal["font_body"]
    for axis_name in ("value_axis", "category_axis"):
        try:
            axis = getattr(chart, axis_name)
        except ValueError:
            continue  # pie/doughnut have no axes
        axis.tick_labels.font.size = Pt(12)
        axis.tick_labels.font.color.rgb = hex_rgb(pal["text_muted"])
        axis.format.line.color.rgb = hex_rgb(pal["surface"])
        if axis_name == "value_axis":
            axis.has_major_gridlines = True
            axis.major_gridlines.format.line.color.rgb = hex_rgb(pal["surface"])


def add_native_chart(slide, pal, chart_kind, categories, series,
                     left, top, w, h, series_name=""):
    """Add a palette-styled native chart.

    series: list of (value, ...) floats for a single series, or
            dict {name: values} for multiple series.
    """
    xl_type = CHART_TYPES.get(chart_kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
    chart_data = ChartData()
    chart_data.categories = categories
    if isinstance(series, dict):
        for name, values in series.items():
            chart_data.add_series(name, values)
        n_series = len(series)
    else:
        chart_data.add_series(series_name or "Series 1", series)
        n_series = 1

    frame = slide.shapes.add_chart(
        xl_type, Inches(left), Inches(top), Inches(w), Inches(h), chart_data
    )
    chart = frame.chart
    chart.has_title = False

    make_chart_transparent(chart)
    _style_axes(chart, pal)

    accents = [pal["accent1"], pal["accent2"], pal["accent3"]]
    if chart_kind in ("pie", "doughnut"):
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        for i, point in enumerate(chart.plots[0].series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = hex_rgb(accents[i % len(accents)])
    elif chart_kind == "line":
        chart.has_legend = n_series > 1
        for i, s in enumerate(chart.plots[0].series):
            s.smooth = True
            line = s.format.line
            line.color.rgb = hex_rgb(accents[i % len(accents)])
            line.width = Pt(2.5)
    else:
        chart.has_legend = n_series > 1
        for i, s in enumerate(chart.plots[0].series):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = hex_rgb(accents[i % len(accents)])

    return chart


# ── benchmark / target reference line ────────────────────────────────────────
def _nice_ceil(v):
    """Round up to a 'nice' axis maximum (1/2/2.5/5 x 10^k)."""
    import math
    if v <= 0:
        return 1
    exp = math.floor(math.log10(v))
    for mult in (1, 2, 2.5, 5, 10):
        cand = mult * 10 ** exp
        if cand >= v:
            return cand
    return 10 ** (exp + 1)


def add_benchmark_line(slide, chart, pal, spec, left, top, w, h):
    """Dashed reference line across a bar/column/line chart.

    spec: '120 Industry average'. Sets an explicit value-axis maximum so the
    line height is computable; plot-area insets are approximations — confirm
    placement in visual QA.
    """
    import re
    m = re.match(r"\s*\$?([\d,.]+)\s*(.*)", spec)
    if not m:
        return
    value = float(m.group(1).replace(",", ""))
    label = m.group(2).strip().strip('"') or "Benchmark"

    data_max = value
    for s in chart.plots[0].series:
        data_max = max(data_max, max(v for v in s.values if v is not None))
    axis_max = _nice_ceil(data_max * 1.05)
    va = chart.value_axis
    va.maximum_scale = float(axis_max)
    va.minimum_scale = 0.0

    # approximate plot box within the chart frame
    px, pw = left + 0.09 * w, w * 0.88
    py, ph = top + 0.04 * h, h * 0.80
    y = py + (1 - value / axis_max) * ph

    from pptx.enum.shapes import MSO_CONNECTOR
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(px), Inches(y), Inches(px + pw), Inches(y))
    conn.line.color.rgb = hex_rgb(pal["accent3"])
    conn.line.width = Pt(1.75)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    from pptx.enum.text import PP_ALIGN
    tb = slide.shapes.add_textbox(Inches(px + 0.05), Inches(y - 0.32),
                                  Inches(2.4), Inches(0.28))
    tf = tb.text_frame
    run = tf.paragraphs[0].add_run()
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    run.text = label
    run.font.size = Pt(11)
    run.font.bold = True
    # text_muted clears 4.5:1 on every palette (accent3 does not)
    run.font.color.rgb = hex_rgb(pal["text_muted"])
    run.font.name = pal["font_label"]
