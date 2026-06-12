"""Build-time text measurement — predict wrapping before the deck renders.

estimate_lines() greedily word-wraps the way PowerPoint will, using PIL glyph
metrics when Pillow plus a known sans TTF is available, else an
average-character-width approximation (within ~10% on typical copy).
bullets_fit() scores a bullet block against a layout's vertical budget so
build_deck.validate() and the builders share one capacity model; mismatched
copies of that arithmetic were the #1 source of render-time overflow.
apply_autofit() writes the <a:normAutofit> shrink hint for the 100-140%
overflow band.
"""
import math
import re

LINE_SPACING = 1.2
WIDTH_FACTOR = 0.50        # fallback: avg glyph width as a fraction of font size
WIDTH_FACTOR_BOLD = 0.53
MARKER_INDENT = 0.30       # accent-square marker offset (builders._draw_marker)
ICON_INDENT = 0.66         # icon-in-circle marker offset

# keep in sync with builders.ICON_RX and helpers._RICH_RX
_ICON_RX = re.compile(r"^icon:([a-z0-9-]+)\s+(.*)$")
_RICH_RX = re.compile(r"\*\*(.+?)\*\*|\{accent\}(.+?)\{/\}")

_FONT_CANDIDATES = {
    False: (
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ),
    True: (
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/System/Library/Fonts/Helvetica.ttc",  # regular face — close enough
    ),
}
_FONT_CACHE = {}


def _load_font(font_pt, bold):
    """PIL font for (size, bold), cached; None when PIL or every TTF is missing."""
    key = (font_pt, bold)
    if key not in _FONT_CACHE:
        font = None
        try:
            from PIL import ImageFont
            for path in _FONT_CANDIDATES[bool(bold)]:
                try:
                    font = ImageFont.truetype(path, int(round(font_pt)))
                    break
                except OSError:
                    continue
        except ImportError:
            pass
        _FONT_CACHE[key] = font
    return _FONT_CACHE[key]


def strip_markup(text):
    """Measurement form of a bullet: drop icon: prefix and rich-text markers.
    Returns (plain_text, has_icon)."""
    text = (text or "").strip()
    m = _ICON_RX.match(text)
    if m:
        text = m.group(2)
    text = _RICH_RX.sub(lambda mm: mm.group(1) or mm.group(2), text)
    return text, bool(m)


def estimate_lines(text, font_pt, box_w_in, bold=False):
    """Estimated wrapped line count of text in a box_w_in-wide frame (greedy
    word wrap; tokens wider than the box count ceil(width/box) lines)."""
    words = (text or "").split()
    if not words or box_w_in <= 0:
        return 0
    font = _load_font(font_pt, bold)
    if font is not None:
        def measure(s):
            return font.getlength(s) / 72.0
    else:
        per_char = font_pt * (WIDTH_FACTOR_BOLD if bold else WIDTH_FACTOR) / 72.0

        def measure(s):
            return len(s) * per_char
    space_w = measure(" ")
    lines, cur = 0, 0.0
    for word in words:
        w = measure(word)
        if w > box_w_in:                       # unbreakable long token
            if cur:
                lines += 1
            lines += math.ceil(w / box_w_in)
            cur = 0.0
            continue
        need = w if not cur else cur + space_w + w
        if need <= box_w_in:
            cur = need
        else:
            lines += 1
            cur = w
    if cur:
        lines += 1
    return lines


def estimate_height_in(text, font_pt, box_w_in, bold=False,
                       line_spacing=LINE_SPACING):
    """Estimated rendered height (inches) of wrapped text."""
    lines = estimate_lines(text, font_pt, box_w_in, bold)
    return lines * font_pt * line_spacing / 72.0


def bullets_fit(bullets, font_pt, col_w_in, avail_h_in, cols=1, step_in=0.0,
                bold=False):
    """Score a bullet block against its vertical budget.

    Splits bullets across cols columns column-major (like build_bullet_slide);
    each bullet consumes at least its fixed step_in slot. Returns
    (ratio, total_lines, overflowing_indices) where:
      ratio               = worst column's estimated height / avail_h_in —
                            >1.0 means autofit territory, >1.4 means split slide.
      total_lines         = sum of wrapped lines across all bullets.
      overflowing_indices = list of bullet indices (0-based) whose wrapped height
                            exceeds step_in AND which are not the last bullet in
                            their column (i.e. they would collide with the next
                            bullet's fixed-y slot).  Empty when step_in == 0.
    """
    bullets = list(bullets or [])
    if not bullets or avail_h_in <= 0:
        return 0.0, 0, []
    cols = max(int(cols), 1)
    per_col = -(-len(bullets) // cols)
    line_h = font_pt * LINE_SPACING / 72.0
    heights, line_counts, total_lines = [], [], 0
    for b in bullets:
        plain, has_icon = strip_markup(b)
        w = col_w_in - (ICON_INDENT if has_icon else MARKER_INDENT)
        lines = max(estimate_lines(plain, font_pt, w, bold), 1)
        total_lines += lines
        heights.append(max(lines * line_h, step_in))
        line_counts.append(lines)
    worst = max(sum(heights[c * per_col:(c + 1) * per_col])
                for c in range(cols))
    # Per-slot collision detection: a non-terminal bullet overflows its slot
    # when its wrapped height exceeds the fixed step_in spacing.
    overflowing_indices = []
    if step_in > 0:
        for c in range(cols):
            col_start = c * per_col
            col_end = min(col_start + per_col, len(bullets))
            for pos in range(col_start, col_end - 1):  # skip last in each col
                if line_counts[pos] * line_h > step_in:
                    overflowing_indices.append(pos)
    return worst / avail_h_in, total_lines, overflowing_indices


def apply_autofit(text_frame, scale_pct, ln_reduction_pct=10):
    """Write <a:normAutofit fontScale=".." lnSpcReduction=".."/> into the
    frame's bodyPr (values are 1000ths of a percent: 85% -> "85000")."""
    from lxml import etree
    from pptx.oxml.ns import qn
    bodyPr = text_frame._txBody.find(qn("a:bodyPr"))
    for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
        for el in bodyPr.findall(qn(tag)):
            bodyPr.remove(el)
    fit = etree.SubElement(bodyPr, qn("a:normAutofit"))
    fit.set("fontScale", str(int(scale_pct) * 1000))
    fit.set("lnSpcReduction", str(int(ln_reduction_pct) * 1000))
    return fit
