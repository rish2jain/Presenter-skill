#!/usr/bin/env python3
"""
build_deck.py — Build a .pptx from a markdown outline.

Usage:
    python scripts/build_deck.py outline.md --output deck.pptx
        [--palette midnight-executive] [--template template.pptx]
        [--assets-dir assets] [--check]

--check validates the outline (layouts, images, chart data) without building.
Outline format: references/generation-guide.md.
"""
import argparse
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from builders import LAYOUT_MAP, ctx_blank_layout
from helpers import add_speaker_notes, parse_visual, resolve_image_path
from narrative import check_unsourced_stats, validate_narrative
from palettes import PALETTES, apply_variant, get_palette
from smart_layout import auto_layout
from template_helpers import (find_best_layout, label_placeholder_names,
                              load_template_config, palette_from_theme,
                              validate_template_mapping, LAYOUT_CATEGORIES)

KV_RE = re.compile(r'(\w+)="([^"]*)"')
NUM_RE = re.compile(r"-?\d[\d,]*\.?\d*")

# "- Key: value" prefixes mapped to slide-dict fields
FIELD_KEYS = {
    "Title": "title", "Subtitle": "subtitle", "Heading": "heading",
    "Headline": "heading", "Caption": "caption", "Contact": "contact",
    "Notes": "notes", "Series": "series", "Source": "source",
    "Left label": "left_label", "Right label": "right_label",
    # consulting layouts
    "Quote": "quote", "Attribution": "attribution", "Context": "context",
    "Value": "value", "Label": "label", "Callout": "callout",
    "Current": "current", "Situation": "situation",
    "Recommendation": "recommendation",
    "Bracket": "bracket", "CAGR": "cagr", "Axis-Max": "axis_max",
    "X-axis": "x_axis", "Y-axis": "y_axis",
    "Q1": "q1", "Q2": "q2", "Q3": "q3", "Q4": "q4",
}


def _parse_kv(line):
    return {k.lower(): v for k, v in KV_RE.findall(line)}


def _series_count(slide):
    """Number of chart series declared via **Series:** (comma-separated names)."""
    series = slide.get("series", "")
    if not series or "," not in series:
        return 1
    return len([n for n in series.split(",") if n.strip()])


def _parse_data_point(text, n_series=1):
    """'2024: $42B' -> ('2024', 42.0); multi: '2024: 42, 38' -> ('2024', [42.0, 38.0])."""
    label, sep, value = text.partition(":")
    if not sep:
        return None
    if n_series > 1:
        nums = []
        for part in value.split(","):
            m = NUM_RE.search(part)
            if not m:
                return None
            nums.append(float(m.group().replace(",", "")))
        if len(nums) != n_series:
            return None
        return label.strip().strip('"'), nums
    if value.strip().lower() in ("total", "end"):  # waterfall closing bar
        return label.strip().strip('"'), "total"
    m = NUM_RE.search(value)
    if not m:
        return None
    return label.strip().strip('"'), float(m.group().replace(",", ""))


META_KEYS = {"Palette": "palette", "Footer": "footer",
              "Page-Numbers": "page_numbers", "Size": "size",
              "Density": "density", "Purpose": "purpose",
              "Variant": "variant", "Motif": "motif", "Takeaway": "takeaway",
              "Exhibits": "exhibits", "Auto-Agenda": "auto_agenda",
              "Stamp": "stamp"}


def parse_outline(md_text):
    """Parse a markdown outline into (deck_meta, list of slide dicts).

    Deck-level front-matter: `**Palette:** aurora`, `**Footer:** "Confidential"`,
    `**Page-Numbers:** on`, `**Size:** 16:9|4:3` before the first `## Slide`.
    """
    meta, slides, current, in_data = {}, [], None, False

    for raw in md_text.splitlines():
        line = raw.strip()
        if current is None and line.startswith("**"):
            key, _, value = line.lstrip("*").partition(":**")
            if key.strip() in META_KEYS:
                meta[META_KEYS[key.strip()]] = value.strip().strip('"')
                continue
        if line.startswith("## Appendix"):
            if current:
                slides.append(current)
                current = None
            meta["_appendix_from"] = len(slides)  # subsequent slides are backup
            continue
        if line.startswith("## Slide "):
            if current:
                slides.append(current)
            current = {"bullets": [], "stats": [], "cards": [], "items": [],
                       "data": [], "table_rows": [], "steps": [], "tiles": [],
                       "matrix_items": [], "bars": [], "milestones": [],
                       "left_bullets": [], "right_bullets": []}
            if "_appendix_from" in meta:
                current["_appendix"] = True
            in_data = False
            m = re.match(r"## Slide \d+:\s*(.*)", line)
            if m:
                current["heading"] = m.group(1).strip()
            continue
        if current is None:
            continue
        if not line:
            continue  # blank lines inside **Data:** blocks are ignored

        if line.startswith("**"):
            in_data = False
            key, _, value = line.lstrip("*").partition(":**")
            key, value = key.strip(), value.strip()
            if key == "Layout":
                current["layout"] = value.lower()
            elif key == "Visual":
                current["visual"] = value
            elif key == "Visual-Left":
                current["visual_left"] = value
            elif key == "Visual-Right":
                current["visual_right"] = value
            elif key == "Palette":
                current["palette"] = value.lower()
            elif key == "Series":
                current["series"] = value
            elif key == "Benchmark":
                current["benchmark"] = value
            elif key == "Periods":
                current["periods"] = [s.strip() for s in value.split(",")
                                      if s.strip()]
            elif key == "Data":
                in_data = True
            continue

        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if not all(set(c) <= {"-", ":", " "} for c in cells):  # skip ---- row
                current["table_rows"].append(cells)
            continue

        if not line.startswith("- "):
            continue
        item = line[2:].strip()

        # Known field keys take precedence over (and terminate) a Data block,
        # so '- Heading: "32% YoY"' after **Data:** isn't read as a data point.
        matched = False
        for prefix, field in FIELD_KEYS.items():
            # Q1-Q4 are quadrant labels only on matrix slides; elsewhere they
            # are ordinary labels (e.g. '- Q1: 4.2, 3.1' chart data rows)
            if field in ("q1", "q2", "q3", "q4") and \
                    current.get("layout") != "matrix-2x2":
                continue
            if item.startswith(prefix + ":"):
                current[field] = item[len(prefix) + 1:].strip().strip('"')
                matched = True
                in_data = False
                break
        if matched:
            continue

        if in_data:
            point = _parse_data_point(item, _series_count(current))
            if point:
                current["data"].append(point)
                continue
            in_data = False  # fall through: not a data line after all

        if item.startswith("Left:"):
            current["left_bullets"].append(item[5:].strip().strip('"'))
        elif item.startswith("Right:"):
            current["right_bullets"].append(item[6:].strip().strip('"'))
        elif re.match(r"Card \d+:", item):
            current["cards"].append(_parse_kv(item))
        elif re.match(r"Step \d+:", item):
            current["steps"].append(_parse_kv(item))
        elif re.match(r"Tile \d+:", item):
            current["tiles"].append(_parse_kv(item))
        elif item.startswith("Item:"):
            current["matrix_items"].append(_parse_kv(item))
        elif item.startswith("Bar:"):
            current["bars"].append(_parse_kv(item))
        elif item.startswith("Milestone:"):
            current["milestones"].append(_parse_kv(item))
        elif item.startswith("Value=") or item.startswith("Stat"):
            kv = _parse_kv(item)
            if kv:
                current["stats"].append(kv)
        elif "date=" in item.lower() and KV_RE.search(item):
            current["items"].append(_parse_kv(item))
        elif re.match(r"(Point|Finding) \d+:", item):
            current["bullets"].append(item.split(":", 1)[1].strip().strip('"'))
        else:
            current["bullets"].append(item.strip('"'))

    if current:
        slides.append(current)
    for slide in slides:
        if not slide.get("layout"):
            slide["layout"] = auto_layout(slide)
    return meta, slides


def _agenda_slide(sections, current=None):
    slide = {"bullets": list(sections), "stats": [], "cards": [], "items": [],
             "data": [], "table_rows": [], "steps": [], "tiles": [],
             "matrix_items": [], "bars": [], "milestones": [],
             "left_bullets": [], "right_bullets": [],
             "layout": "agenda", "heading": "Agenda",
             "notes": "Auto-generated agenda tracker.", "_auto": True}
    if current:
        slide["current"] = current
        slide["heading"] = "Where we are"
    return slide


def apply_auto_agenda(meta, slides):
    """**Auto-Agenda:** on  -> overview agenda after the title slide.
                       track -> + current-highlighted agenda after each divider.
    Sections come from section-divider headings; no dividers -> no-op."""
    mode = meta.get("auto_agenda", "").lower()
    if mode not in ("on", "track"):
        return slides
    sections = [s.get("heading") or s.get("title", "")
                for s in slides if s.get("layout") == "section-divider"]
    if not sections:
        return slides
    out = []
    for i, s in enumerate(slides):
        out.append(s)
        if i == 0 and not s.get("_appendix"):
            out.append(_agenda_slide(sections))
        if mode == "track" and s.get("layout") == "section-divider" \
                and not s.get("_appendix"):
            out.append(_agenda_slide(
                sections, current=s.get("heading") or s.get("title", "")))
    return out


# ── Validation ───────────────────────────────────────────────────────────────
# Layouts whose headings should pass the consulting "titles test"
ACTION_TITLE_LAYOUTS = {
    "bullet-list", "two-column-split", "exec-summary", "exec-summary-scqa",
    "comparison", "table", "stat-callout", "timeline", "waterfall",
    "matrix-2x2", "chart-callout", "dashboard", "funnel", "harvey-scorecard",
    "process-flow", "mekko", "gantt", "bar-mekko",
}


def validate(slides, ctx, meta=None):
    """Fail-fast outline checks. Returns (errors, warnings)."""
    meta = meta or {}
    errors, warnings = [], []
    if not slides:
        return ["Outline contains no '## Slide N:' sections"], []

    for n, p in enumerate(slides, 1):
        where = f"Slide {n}"
        layout = p.get("layout", "bullet-list")
        if layout not in LAYOUT_MAP:
            errors.append(f"{where}: unknown layout '{layout}' "
                          f"(valid: {', '.join(sorted(LAYOUT_MAP))})")
            continue
        if "palette" in p and p["palette"] not in PALETTES:
            warnings.append(f"{where}: unknown palette '{p['palette']}', "
                            "default will be used")

        for vis_key in ("visual", "visual_left", "visual_right"):
            kind, value = parse_visual(p.get(vis_key, ""))
            if not kind:
                continue
            if kind == "image" and not resolve_image_path(value, ctx):
                errors.append(f"{where}: image not found for {vis_key}: '{value}'")
            elif kind == "chart":
                if value not in __import__("charts").CHART_TYPES:
                    errors.append(f"{where}: unknown chart type '{value}'")
                if vis_key == "visual" and not p.get("data"):
                    errors.append(f"{where}: 'chart:{value}' visual but no "
                                  "**Data:** block (or no parsable '- label: value' lines)")
                elif vis_key != "visual":
                    errors.append(f"{where}: charts on {vis_key} are not supported "
                                  "(use **Visual:** on two-column-split)")

        if not p.get("notes") and not p.get("_appendix") \
                and layout != "section-divider":
            warnings.append(f"{where}: missing '- Notes:' speaker notes")

        # Titles test: content-slide headings should be action titles —
        # full-sentence takeaways, not topic labels ("Leadership" fails).
        if (layout in ACTION_TITLE_LAYOUTS and not p.get("_appendix")
                and len(p.get("heading", "").split()) < 5):
            warnings.append(
                f"{where}: heading {p.get('heading', '')!r} is a topic label, "
                "not an action title — state the takeaway as a sentence "
                "(run --titles for the titles test)")

        if layout in ("title", "closing") and not p.get("title") and p.get("heading"):
            warnings.append(f"{where}: no '- Title:' — using '## Slide N:' heading "
                            f"({p['heading']!r}) as title text")

        n_series = _series_count(p)
        if n_series > 1 and parse_visual(p.get("visual", ""))[0] == "chart":
            if not p.get("data"):
                errors.append(
                    f"{where}: multi-series chart needs **Data:** rows with "
                    f"{n_series} comma-separated values per line "
                    f"(**Series:** {p.get('series')})")
            else:
                for label, vals in p["data"]:
                    if not isinstance(vals, list) or len(vals) != n_series:
                        errors.append(
                            f"{where}: multi-series row {label!r} needs {n_series} "
                            f"comma-separated values (got **Series:** {p.get('series')})")
                        break

        import builders
        cap = builders.density()["bullet_max"]
        if len(p.get("bullets", [])) > cap:
            warnings.append(f"{where}: {len(p['bullets'])} bullets — only the "
                            f"first {cap} render; split the slide")
        if layout == "stat-callout" and not p.get("stats"):
            errors.append(f"{where}: stat-callout requires "
                          '\'- Value="..." Label="..." Sublabel="..."\' lines')
        if layout == "table" and not p.get("table_rows"):
            errors.append(f"{where}: table layout requires markdown table rows")
        if layout in ("cards-3", "cards-4"):
            want = 3 if layout == "cards-3" else 4
            if len(p.get("cards", [])) != want:
                warnings.append(f"{where}: {layout} has {len(p.get('cards', []))} "
                                f"cards (expected {want})")
        if layout == "timeline" and not p.get("items") and not p.get("bullets"):
            errors.append(f"{where}: timeline needs '- Date=\"..\" Title=\"..\"' "
                          "items or 'Date: text' bullets")
        if layout == "full-image" and parse_visual(p.get("visual", ""))[0] != "image":
            warnings.append(f"{where}: full-image without an image visual "
                            "renders a plain background")

        # consulting layout requirements
        numeric_data = [d for d in p.get("data", [])
                        if isinstance(d[1], (int, float, list))]
        if parse_visual(p.get("visual", ""))[0] == "chart" and \
                any(d[1] == "total" for d in p.get("data", [])):
            errors.append(f"{where}: 'total' rows are only valid on waterfall "
                          "layouts, not charts")
        if layout == "waterfall" and len(p.get("data", [])) < 2:
            errors.append(f"{where}: waterfall needs a **Data:** block — start "
                          "value, +/- deltas, optional '- End label: total' row")
        if layout == "funnel" and len(numeric_data) < 2:
            errors.append(f"{where}: funnel needs a **Data:** block with 2+ "
                          "numeric '- Stage: value' rows")
        if layout == "matrix-2x2" and not p.get("matrix_items"):
            errors.append(f"{where}: matrix-2x2 needs '- Item: Name=\"..\" "
                          "X=\"0-1\" Y=\"0-1\"' rows")
        if layout == "harvey-scorecard" and len(p.get("table_rows", [])) < 2:
            errors.append(f"{where}: harvey-scorecard needs a markdown table "
                          "(header row + criteria rows, cells 0-4)")
        if layout in ("process-flow", "next-steps") and not p.get("steps") \
                and not p.get("bullets"):
            errors.append(f"{where}: {layout} needs '- Step N: ...' rows")
        if layout == "big-number" and not p.get("value"):
            errors.append(f"{where}: big-number needs '- Value: \"$4.2B\"'")
        if layout == "chart-callout":
            if not p.get("callout"):
                errors.append(f"{where}: chart-callout needs '- Callout: \"...\"'")
            if not p.get("data") and parse_visual(p.get("visual", ""))[0] != "image":
                errors.append(f"{where}: chart-callout needs a chart **Data:** "
                              "block or an image visual")
        if layout == "dashboard" and not p.get("tiles"):
            errors.append(f"{where}: dashboard needs '- Tile N: Value=\"..\" "
                          "Label=\"..\" Delta=\"..\" Trend=\"up|down\"' rows")
        if layout == "quote-evidence" and not p.get("quote"):
            errors.append(f"{where}: quote-evidence needs '- Quote: \"...\"'")
        if layout == "exec-summary-scqa" and not (
                p.get("situation") and p.get("recommendation")):
            errors.append(f"{where}: exec-summary-scqa needs '- Situation:' and "
                          "'- Recommendation:' (findings via '- Finding N:')")
        if layout == "agenda" and not p.get("bullets"):
            errors.append(f"{where}: agenda needs section names as bullets")
        if layout == "mekko":
            if _series_count(p) < 2 or not any(
                    isinstance(d[1], list) for d in p.get("data", [])):
                errors.append(f"{where}: mekko needs **Series:** with 2+ names "
                              "and multi-value **Data:** rows")
        if layout == "gantt":
            if not p.get("periods"):
                errors.append(f"{where}: gantt needs '**Periods:** Q1, Q2, ...'")
            if not p.get("bars"):
                errors.append(f"{where}: gantt needs '- Bar: Row=\"..\" "
                              "Label=\"..\" Start=\"1\" End=\"2\"' rows")
            for bar in p.get("bars", []):
                try:
                    s, e = float(bar.get("start", 1)), float(bar.get("end", 1))
                    if not (1 <= s <= e <= len(p.get("periods", [])) ):
                        errors.append(f"{where}: Bar {bar.get('label', '')!r} "
                                      f"Start/End outside 1-{len(p['periods'])}")
                except ValueError:
                    errors.append(f"{where}: Bar {bar.get('label', '')!r} has "
                                  "non-numeric Start/End")
        if layout == "bar-mekko":
            ok_bars = [b for b in p.get("bars", [])
                       if "size" in b and "value" in b]
            if len(ok_bars) < 2:
                errors.append(f"{where}: bar-mekko needs 2+ '- Bar: "
                              'Label=".." Size=".." Value=".."\' rows')
            for bar in ok_bars:
                try:
                    size, value = float(bar["size"]), float(bar["value"])
                    if size <= 0 or value < 0:
                        errors.append(f"{where}: Bar {bar.get('label', '')!r} "
                                      "needs Size>0 and Value>=0")
                except ValueError:
                    errors.append(f"{where}: Bar {bar.get('label', '')!r} has "
                                  "non-numeric Size/Value")

    # Narrative + data integrity
    warnings.extend(validate_narrative(meta, slides))
    warnings.extend(check_unsourced_stats(slides))

    template_path = ctx.get("template_path")
    if template_path and Path(template_path).is_file():
        from pptx import Presentation
        config = load_template_config(template_path)
        prs = Presentation(str(template_path))
        warnings.extend(validate_template_mapping(prs, slides, config))

    # Layout variety: 3+ identical consecutive layouts reads as monotonous
    # (appendix/backup slides are deliberately dense and repetitive — exempt)
    layouts = [p.get("layout", "bullet-list") for p in slides
               if not p.get("_appendix")]
    run_start = 0
    for i in range(1, len(layouts) + 1):
        if i == len(layouts) or layouts[i] != layouts[run_start]:
            if i - run_start >= 3 and layouts[run_start] not in ("full-image",):
                warnings.append(
                    f"Slides {run_start + 1}-{i}: '{layouts[run_start]}' used "
                    f"{i - run_start}x consecutively — vary layouts (cards, "
                    "stat-callout, two-column-split...)")
            run_start = i
    return errors, warnings


# ── Template-aware build ─────────────────────────────────────────────────────
TEMPLATE_WANT = LAYOUT_CATEGORIES


def _fill_placeholders(slide, p):
    """Fill title/body/subtitle placeholders from the slide dict."""
    from pptx.enum.shapes import PP_PLACEHOLDER
    title_text = p.get("title") or p.get("heading", "")
    body_lines = p.get("bullets", [])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        ph_type = ph.placeholder_format.type
        if idx == 0 or ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            ph.text = title_text
        elif ph_type == PP_PLACEHOLDER.SUBTITLE or (idx == 1 and not body_lines):
            ph.text = p.get("subtitle", "")
        elif ph.has_text_frame and body_lines:
            tf = ph.text_frame
            tf.clear()
            for i, line in enumerate(body_lines):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                para.text = line
            body_lines = []  # only fill the first body placeholder


def build_template_slide(prs, p, layout_name, ctx, config=None):
    """Try a placeholder-based slide from the user's template."""
    from pptx.enum.shapes import PP_PLACEHOLDER as PH

    kind, value = parse_visual(p.get("visual", ""))
    if layout_name == "two-column-split" and kind == "image":
        layout, _ = find_best_layout(prs, "picture", config)
        if layout is not None:
            path = resolve_image_path(value, ctx)
            if path:
                slide = prs.slides.add_slide(layout)
                label_placeholder_names(slide, layout)
                _fill_placeholders(slide, p)
                for ph in slide.placeholders:
                    if ph.placeholder_format.type == PH.PICTURE:
                        ph.insert_picture(str(path))
                        break
                return slide

    want = TEMPLATE_WANT.get(layout_name)
    if not want:
        return None
    layout, score = find_best_layout(prs, want, config)
    if layout is None or score < 1:
        return None
    slide = prs.slides.add_slide(layout)
    label_placeholder_names(slide, layout)
    _fill_placeholders(slide, p)
    return slide


# ── Main ─────────────────────────────────────────────────────────────────────
SIZES = {"16:9": (13.33, 7.5), "4:3": (10.0, 7.5)}
NO_FOOTER_LAYOUTS = {"title", "closing", "section-divider", "full-image"}


def create_presentation(template_path=None, size="16:9"):
    if template_path:
        return Presentation(template_path)
    prs = Presentation()
    w, h = SIZES.get(size, SIZES["16:9"])
    prs.slide_width = Inches(w)
    prs.slide_height = Inches(h)
    return prs


def _add_footer(slide, i, pal, footer, page_numbers, appendix=False):
    import builders
    if footer:
        builders.add_tb(slide, footer, 0.7, 7.08, 6.0, 0.35, size=11,
                        color=pal["text_muted"], font=pal["font_label"])
    if page_numbers:
        from pptx.enum.text import PP_ALIGN
        label = f"B·{i + 1}" if appendix else str(i + 1)
        builders.add_tb(slide, label, 11.9, 7.08, 0.8, 0.35, size=11,
                        color=pal["text_muted"], font=pal["font_label"],
                        align=PP_ALIGN.RIGHT)


def _add_kicker(slide, pal, section):
    """Section tracker: small current-section label top-right (consulting
    navigation convention for long decks)."""
    if not section:
        return
    import builders
    from pptx.enum.text import PP_ALIGN
    builders.add_tb(slide, section.upper(), 7.6, 0.12, 5.0, 0.32, size=11,
                    bold=True, color=pal["text_muted"], font=pal["font_label"],
                    align=PP_ALIGN.RIGHT)


def _add_source(slide, pal, source, exhibit_no=None):
    """Exhibit attribution footnote (every consulting exhibit cites a source).
    With **Exhibits:** on, sourced slides are numbered 'Exhibit N · Source: ...'."""
    import builders
    text = source if source.lower().startswith("source") else f"Source: {source}"
    if exhibit_no:
        text = f"Exhibit {exhibit_no} · {text}"
    builders.add_tb(slide, text, 0.7, 6.80, 9.5, 0.3, size=11,
                    color=pal["text_muted"], font=pal["font_label"])


def _add_stamp(slide, pal, text):
    """Bordered status tag (DRAFT, CONFIDENTIAL) top-left, above the heading."""
    import builders
    from pptx.enum.text import PP_ALIGN
    box = builders.add_rect(slide, 0.7, 0.14, 1.5, 0.32, pal["bg"],
                            line_hex=pal["accent3"], line_pt=1.0)
    box.fill.background()  # outline only — works on photos and gradients
    builders.add_tb(slide, text.upper(), 0.7, 0.16, 1.5, 0.28, size=11,
                    bold=True, color=pal["text_muted"], font=pal["font_label"],
                    align=PP_ALIGN.CENTER)


def build(outline_path, output_path, palette_key=None,
          template_path=None, assets_dir=None, check_only=False, size=None,
          density=None, variant=None):
    outline_path = Path(outline_path)
    ctx = {
        "outline_dir": outline_path.parent,
        "assets_dir": Path(assets_dir) if assets_dir else Path("assets"),
        "template_path": template_path,
    }
    meta, slides_data = parse_outline(outline_path.read_text(encoding="utf-8"))
    slides_data = apply_auto_agenda(meta, slides_data)

    import builders
    variant_key = variant or meta.get("variant", "")
    palette_key, density = apply_variant(variant_key, palette_key, density)
    builders.set_density(density or meta.get("density", ""))
    errors, warnings = validate(slides_data, ctx, meta)
    for w in warnings:
        print(f"  [WARN] {w}", file=sys.stderr)
    if errors:
        for e in errors:
            print(f"  [ERROR] {e}", file=sys.stderr)
        print(f"\nOutline validation failed: {len(errors)} error(s).", file=sys.stderr)
        return False
    if check_only:
        print(f"Outline OK: {len(slides_data)} slides, "
              f"{len(warnings)} warning(s).")
        return True

    # precedence: CLI flag > outline front-matter > variant preset > default
    pal_default = get_palette(palette_key or meta.get("palette", ""))
    template_config = None
    if template_path:
        template_config = load_template_config(template_path)
        from profile_template import extract_theme_colors
        from pptx import Presentation as _Prs
        theme = extract_theme_colors(_Prs(str(template_path)))
        theme_pal = palette_from_theme(theme)
        if theme_pal:
            for k, v in theme_pal.items():
                if not k.startswith("_"):
                    pal_default[k] = v
    if meta.get("motif"):
        pal_default = {**pal_default, "motif": meta["motif"]}
    size = size or meta.get("size", "16:9")
    footer = meta.get("footer", "")
    page_numbers = meta.get("page_numbers", "").lower() in ("on", "true", "yes")
    exhibits_on = meta.get("exhibits", "").lower() in ("on", "true", "yes")
    exhibit_no = 0

    prs = create_presentation(template_path, size)
    import builders
    builders.set_canvas(prs)  # rescale styled builders to the actual slide size

    failures = 0
    built_slides = []
    current_section = None
    for i, p in enumerate(slides_data):
        pal = get_palette(p["palette"]) if "palette" in p else pal_default
        layout_name = p.get("layout", "bullet-list")
        if layout_name == "section-divider":
            current_section = p.get("heading") or p.get("title", "")
        print(f"  Building slide {i + 1}: "
              f"{p.get('heading', p.get('title', '?'))} [{layout_name}]")
        try:
            slide = None
            templated = False
            if template_path:
                slide = build_template_slide(prs, p, layout_name, ctx, template_config)
                templated = slide is not None
            if slide is None:
                slide = LAYOUT_MAP[layout_name](prs, p, pal, ctx)
            if p.get("notes"):
                add_speaker_notes(slide, p["notes"])
            if not templated and layout_name not in NO_FOOTER_LAYOUTS:
                _add_footer(slide, i, pal, footer, page_numbers,
                            appendix=p.get("_appendix", False))
                _add_kicker(slide, pal,
                            "BACKUP" if p.get("_appendix") else current_section)
            if not templated and p.get("source"):
                exhibit_no += 1
                _add_source(slide, pal, p["source"],
                            exhibit_no if exhibits_on else None)
            if not templated and meta.get("stamp"):
                _add_stamp(slide, pal, meta["stamp"])
            built_slides.append(slide)
        except Exception as e:
            failures += 1
            print(f"  [ERROR] Slide {i + 1} failed: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    if failures:
        print(f"\nBuild aborted: {failures} slide(s) failed — no file written.",
              file=sys.stderr)
        return False

    prs.save(output_path)
    print(f"\nSaved: {output_path} ({len(built_slides)} slides)")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build a .pptx from a markdown outline.")
    parser.add_argument("outline", help="Path to outline.md")
    parser.add_argument("--output", default="deck.pptx", help="Output .pptx path")
    parser.add_argument("--palette", default=None,
                        choices=sorted(PALETTES),
                        help="Color palette (overrides outline front-matter)")
    parser.add_argument("--template", default=None, help="Optional .pptx template")
    parser.add_argument("--assets-dir", default=None,
                        help="Root dir for user-images/ and auto/ (default: ./assets)")
    parser.add_argument("--size", default=None, choices=sorted(SIZES),
                        help="Slide aspect (default 16:9)")
    parser.add_argument("--density", default=None,
                        choices=("compact", "comfortable"),
                        help="Content density (default compact)")
    parser.add_argument("--variant", default=None,
                        choices=("a", "b", "c", "consulting"),
                        help="Palette+density preset (a=default, b=aurora/comfortable)")
    parser.add_argument("--check", action="store_true",
                        help="Validate the outline only; do not build")
    parser.add_argument("--titles", action="store_true",
                        help="Print slide titles in order (the consulting "
                             "'titles test' — they should read as an argument)")
    args = parser.parse_args()

    if args.titles:
        _, _slides = parse_outline(Path(args.outline).read_text(encoding="utf-8"))
        print("Titles test — read top to bottom; it should work as an argument:\n")
        for n, s in enumerate(_slides, 1):
            tag = " [BACKUP]" if s.get("_appendix") else ""
            print(f"  {n:2d}. {s.get('title') or s.get('heading', '?')}{tag}")
        sys.exit(0)

    ok = build(args.outline, args.output, args.palette,
               args.template, args.assets_dir, args.check, args.size,
               args.density, args.variant)
    sys.exit(0 if ok else 1)
