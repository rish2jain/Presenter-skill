"""Color palettes and per-palette font stacks for presentation-deck skill.

Design rules encoded here (per the 60-30-10 dominance rule): bg/bg_deep/surface
carry ~60-70% of visual weight, accent1 is THE accent (~10%), accent2/3 support.
Muted text colors are chosen to clear WCAG 4.5:1 against `bg`.
"""
import re

from pptx.dml.color import RGBColor

PALETTES = {
    # ── Dark themes ──────────────────────────────────────────────────────────
    "midnight-executive": {  # finance, strategy, luxury
        "bg": "0A0F1E", "bg_deep": "060912", "surface": "1A2035",
        "accent1": "C9A84C", "accent2": "E8D5B5", "accent3": "8B6914",
        "text": "F1F5F9", "text_muted": "94A3B8",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "aurora": {  # AI, tech, innovation
        "bg": "0F0A2A", "bg_deep": "080517", "surface": "1E1040",
        "accent1": "7C3AED", "accent2": "06B6D4", "accent3": "A78BFA",
        "text": "F8FAFC", "text_muted": "CBD5E1",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "venture-pitch": {  # startups, product launches
        "bg": "18181B", "bg_deep": "0A0A0A", "surface": "27272A",
        "accent1": "F97316", "accent2": "14B8A6", "accent3": "FEF3C7",
        "text": "FAFAFA", "text_muted": "A1A1AA",
        "dark": True,
        "font_title": "Trebuchet MS", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "forest": {  # sustainability, ESG
        "bg": "0D1B0E", "bg_deep": "080F08", "surface": "1A2E1C",
        "accent1": "22C55E", "accent2": "A3B18A", "accent3": "8B6914",
        "text": "F0FDF4", "text_muted": "86EFAC",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "teal-trust": {  # healthcare, fintech trust, services
        "bg": "0B2B29", "bg_deep": "06201E", "surface": "14403C",
        "accent1": "14B8A6", "accent2": "99F6E4", "accent3": "5EEAD4",
        "text": "F0FDFA", "text_muted": "8CCFC6",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "charcoal-minimal": {  # design, architecture, monochrome minimal
        "bg": "1C1C1E", "bg_deep": "111113", "surface": "2A2A2E",
        "accent1": "FAFAF9", "accent2": "A8A29E", "accent3": "78716C",
        "text": "F5F5F4", "text_muted": "A1A1AA",
        "dark": True,
        "font_title": "Arial Black", "font_body": "Arial", "font_label": "Arial",
    },
    "ocean-gradient": {  # data, cloud, maritime, deep tech
        "bg": "082F49", "bg_deep": "041C30", "surface": "0D4060",
        "accent1": "06B6D4", "accent2": "38BDF8", "accent3": "7DD3FC",
        "text": "F0F9FF", "text_muted": "94C6E0",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    # ── Light themes (backgrounds intentionally never pure #FFFFFF) ─────────
    "swiss-light": {  # corporate, education, tutorials
        "bg": "FAF8F5", "bg_deep": "F0EDE8", "surface": "FFFFFF",
        "accent1": "2563EB", "accent2": "F97066", "accent3": "D4A574",
        "text": "0F172A", "text_muted": "57657B",
        "dark": False,
        "font_title": "Palatino Linotype", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "warm-terracotta": {  # consumer, hospitality, editorial warmth
        "bg": "FAF3EC", "bg_deep": "F1E6DA", "surface": "FFFFFF",
        "accent1": "C2552C", "accent2": "8A9B6E", "accent3": "D9A87E",
        "text": "2D2A26", "text_muted": "6E6258",
        "dark": False,
        "font_title": "Georgia", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "berry-cream": {  # lifestyle, brand, creative
        "bg": "FBF6F0", "bg_deep": "F3E9E0", "surface": "FFFFFF",
        "accent1": "8E2A4F", "accent2": "C98A9E", "accent3": "B5838D",
        "text": "33222B", "text_muted": "75606B",
        "dark": False,
        "font_title": "Georgia", "font_body": "Calibri", "font_label": "Calibri Light",
    },
}

DEFAULT_PALETTE = "midnight-executive"

FONT_DEFAULTS = {"font_title": "Gill Sans MT", "font_body": "Calibri",
                 "font_label": "Calibri Light"}
REQUIRED_KEYS = {"bg", "bg_deep", "surface", "accent1", "accent2", "accent3",
                 "text", "text_muted", "dark"}


def _fill_defaults(pal):
    pal.setdefault("motif", "icon-circle")
    for k, v in FONT_DEFAULTS.items():
        pal.setdefault(k, v)
    pal.setdefault("chart_series",
                   [pal["accent1"], pal["accent2"], pal["accent3"]])
    return pal


for _key, _pal in PALETTES.items():
    _fill_defaults(_pal)


def load_custom_palettes(palettes_dir):
    """Merge user palette JSONs (one palette per <name>.json) into PALETTES.

    Validates the required token schema; invalid files are skipped with a
    warning (never trust external data). Returns the list of loaded names.
    """
    import json
    import sys
    from pathlib import Path
    palettes_dir = Path(palettes_dir)
    loaded = []
    if not palettes_dir.is_dir():
        return loaded
    for f in sorted(palettes_dir.glob("*.json")):
        try:
            pal = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            print(f"  [WARN] palette {f.name}: unreadable ({e})", file=sys.stderr)
            continue
        missing = REQUIRED_KEYS - set(pal)
        if missing or not isinstance(pal.get("dark"), bool):
            print(f"  [WARN] palette {f.name}: missing keys "
                  f"{sorted(missing)} — skipped", file=sys.stderr)
            continue
        hex_keys = REQUIRED_KEYS - {"dark"}
        bad = [k for k in sorted(hex_keys)
               if not re.fullmatch(r"[0-9A-Fa-f]{6}", str(pal.get(k, "")))]
        series = pal.get("chart_series")
        if series is not None:
            if len(series) < 3:
                bad.append("chart_series: too few colors")
            bad += [f"chart_series[{i}]" for i, c in enumerate(series)
                    if not re.fullmatch(r"[0-9A-Fa-f]{6}", str(c))]
        if bad:
            print(f"  [WARN] palette {f.name}: non-hex color values "
                  f"{bad} — skipped", file=sys.stderr)
            continue
        PALETTES[f.stem] = _fill_defaults(pal)
        loaded.append(f.stem)
    return loaded


VARIANT_PRESETS = {
    "a": {},
    "b": {"palette": "aurora", "density": "comfortable"},
    "c": {"palette": "swiss-light", "density": "comfortable"},
    "consulting": {"palette": "midnight-executive", "density": "compact"},
}


def apply_variant(variant_key, palette_key=None, density=None):
    preset = VARIANT_PRESETS.get((variant_key or "").lower(), {})
    if palette_key is None and preset.get("palette"):
        palette_key = preset["palette"]
    if density is None and preset.get("density"):
        density = preset["density"]
    return palette_key, density


def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def get_palette(key):
    return PALETTES.get(key, PALETTES[DEFAULT_PALETTE])
