"""Auto-select layout when outline omits **Layout:**."""
from helpers import parse_visual


def auto_layout(slide):
    """Pick a layout from slide content keys. Returns layout name string."""
    if slide.get("layout"):
        return slide["layout"]

    if slide.get("stats"):
        return "stat-callout"
    if slide.get("cards"):
        n = len(slide["cards"])
        if n >= 4:
            return "cards-4"
        if n >= 2:
            return "cards-3"
    if slide.get("table_rows"):
        return "table"
    if slide.get("items") or any(":" in b for b in slide.get("bullets", [])[:3]):
        if slide.get("items") or sum(1 for b in slide.get("bullets", [])
                                     if re_date_bullet(b)) >= 2:
            return "timeline"
    if slide.get("visual_left") or slide.get("visual_right"):
        return "comparison"
    if slide.get("left_bullets") or slide.get("right_bullets"):
        return "comparison"

    kind, _ = parse_visual(slide.get("visual", ""))
    if kind == "image" and slide.get("caption"):
        return "full-image"
    if kind in ("chart", "image"):
        return "two-column-split"

    bullets = slide.get("bullets", [])
    if any(b.lower().startswith("point ") for b in bullets) or len(bullets) <= 4:
        if len(bullets) >= 3 and all(len(b) > 40 for b in bullets[:3]):
            return "exec-summary"

    heading = (slide.get("heading") or slide.get("title") or "").lower()
    if slide.get("contact") or "thank" in heading or "questions" in heading:
        return "closing"
    if slide.get("title") and slide.get("subtitle") and not bullets:
        return "title"
    if not bullets and slide.get("subtitle"):
        return "section-divider"

    return "bullet-list"


def re_date_bullet(text):
    import re
    return bool(re.match(r"^Q[1-4]\s+\d{4}|^\d{4}|^[A-Z][a-z]{2}\s+\d{4}", text.strip()))
