"""Style resolution: reads font sizes from spec front-matter ``style`` block."""

from pptx.util import Pt

# Fallback defaults (used when spec omits a value)
_DEFAULTS = {
    "title_font_size": 36,
    "subtitle_font_size": 20,
    "body_font_size": 20,
    "heading_font_size": 32,
    "column_heading_font_size": 22,
    "column_body_font_size": 18,
}


class Style:
    """Immutable bag of resolved font sizes (as ``Pt`` values)."""

    def __init__(self, spec_style: dict | None = None):
        t = {**_DEFAULTS, **(spec_style or {})}
        self.title_font = Pt(int(t["title_font_size"]))
        self.subtitle_font = Pt(int(t["subtitle_font_size"]))
        self.body_font = Pt(int(t["body_font_size"]))
        self.heading_font = Pt(int(t["heading_font_size"]))
        self.col_heading_font = Pt(int(t["column_heading_font_size"]))
        self.col_body_font = Pt(int(t["column_body_font_size"]))
