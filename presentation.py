"""
Presentation generator: reads a .spec.md file and produces a PowerPoint deck.

Usage:
    python presentation.py <spec-file>

Example:
    python presentation.py .speckit/specifications/ai101.spec.md
"""

import argparse
import os
import re
import sys

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Theme constants
# ---------------------------------------------------------------------------
TITLE_FONT = Pt(36)
SUBTITLE_FONT = Pt(20)
BODY_FONT = Pt(20)
HEADING_FONT = Pt(32)
COL_HEADING_FONT = Pt(22)
COL_BODY_FONT = Pt(18)

# ---------------------------------------------------------------------------
# Spec parser
# ---------------------------------------------------------------------------

def parse_spec(path: str) -> dict:
    """Parse a presentation spec markdown file into metadata + slide list."""
    with open(path, encoding="utf-8") as f:
        text = f.read()

    # Split YAML front matter
    fm_match = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.DOTALL)
    if not fm_match:
        sys.exit("Error: spec file must start with YAML front matter (--- … ---)")
    metadata = yaml.safe_load(fm_match.group(1))
    body = text[fm_match.end():]

    # Split slides on horizontal rules (--- on its own line)
    raw_slides = re.split(r"\n---\s*\n", body)
    slides = []
    for raw in raw_slides:
        raw = raw.strip()
        if not raw:
            continue
        slide = _parse_slide(raw)
        if slide:
            slides.append(slide)

    return {"metadata": metadata, "slides": slides}


def _parse_slide(raw: str) -> dict | None:
    """Parse a single slide block into a structured dict."""
    # Header: ## [type] Title
    header_match = re.match(r"^##\s+\[(\w[\w-]*)\]\s+(.+)", raw)
    if not header_match:
        return None
    slide_type = header_match.group(1).strip()
    title = header_match.group(2).strip()
    rest = raw[header_match.end():].strip()

    # Extract notes (everything after **Notes**: to end)
    notes = ""
    notes_match = re.split(r"\*\*Notes\*\*\s*:\s*", rest, maxsplit=1)
    if len(notes_match) == 2:
        rest = notes_match[0].strip()
        notes = notes_match[1].strip()

    slide: dict = {"type": slide_type, "title": title, "notes": notes}

    if slide_type == "title":
        sub_match = re.search(r"\*\*Subtitle\*\*\s*:\s*(.+)", rest)
        slide["subtitle"] = sub_match.group(1).strip() if sub_match else ""

    elif slide_type == "bullets":
        slide["bullets"] = _extract_bullets(rest)

    elif slide_type == "two-column":
        slide.update(_parse_two_column(rest))

    return slide


def _extract_bullets(text: str) -> list[str]:
    """Extract markdown list items from text."""
    return [m.group(1).strip() for m in re.finditer(r"^[-*]\s+(.+)", text, re.MULTILINE)]


def _parse_two_column(text: str) -> dict:
    """Parse two-column slide fields."""
    result: dict = {}

    lt_match = re.search(r"\*\*Left Title\*\*\s*:\s*(.+)", text)
    rt_match = re.search(r"\*\*Right Title\*\*\s*:\s*(.+)", text)
    result["left_title"] = lt_match.group(1).strip() if lt_match else ""
    result["right_title"] = rt_match.group(1).strip() if rt_match else ""

    # Extract left bullets (between **Left**: and **Right Title**: or **Right**:)
    left_match = re.search(
        r"\*\*Left\*\*\s*:\s*\n(.*?)(?=\*\*Right Title\*\*|\*\*Right\*\*|\*\*Notes\*\*|$)",
        text, re.DOTALL,
    )
    result["left_bullets"] = _extract_bullets(left_match.group(1)) if left_match else []

    # Extract right bullets (between **Right**: and **Notes**: or end)
    right_match = re.search(
        r"\*\*Right\*\*\s*:\s*\n(.*?)(?=\*\*Notes\*\*|$)",
        text, re.DOTALL,
    )
    result["right_bullets"] = _extract_bullets(right_match.group(1)) if right_match else []

    return result


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

def add_title_slide(prs: Presentation, title: str, subtitle: str, notes: str):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle
    slide.shapes.title.text_frame.paragraphs[0].font.size = TITLE_FONT
    slide.placeholders[1].text_frame.paragraphs[0].font.size = SUBTITLE_FONT
    slide.notes_slide.notes_text_frame.text = notes


def add_bullets_slide(prs: Presentation, title: str, bullets: list[str], notes: str):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    slide.shapes.title.text_frame.paragraphs[0].font.size = HEADING_FONT
    body = slide.shapes.placeholders[1].text_frame
    body.clear()
    for i, b in enumerate(bullets):
        p = body.paragraphs[0] if i == 0 else body.add_paragraph()
        p.text = b
        p.level = 0
        p.font.size = BODY_FONT
    slide.notes_slide.notes_text_frame.text = notes


def add_two_column_slide(
    prs: Presentation,
    title: str,
    left_title: str,
    left_bullets: list[str],
    right_title: str,
    right_bullets: list[str],
    notes: str,
):
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # title-only layout
    slide.shapes.title.text = title
    slide.shapes.title.text_frame.paragraphs[0].font.size = HEADING_FONT

    # Left column
    left = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(4.4), Inches(4.8))
    ltf = left.text_frame
    ltf.clear()
    lp = ltf.paragraphs[0]
    lp.text = left_title
    lp.font.size = COL_HEADING_FONT
    lp.font.bold = True
    for b in left_bullets:
        p = ltf.add_paragraph()
        p.text = b
        p.level = 1
        p.font.size = COL_BODY_FONT

    # Right column
    right = slide.shapes.add_textbox(Inches(5.1), Inches(1.7), Inches(4.4), Inches(4.8))
    rtf = right.text_frame
    rtf.clear()
    rp = rtf.paragraphs[0]
    rp.text = right_title
    rp.font.size = COL_HEADING_FONT
    rp.font.bold = True
    for b in right_bullets:
        p = rtf.add_paragraph()
        p.text = b
        p.level = 1
        p.font.size = COL_BODY_FONT

    # Subtle divider line
    line = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(4.95), Inches(1.55), Inches(0.03), Inches(5.1),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(230, 230, 230)
    line.line.fill.background()

    slide.notes_slide.notes_text_frame.text = notes


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------

SLIDE_BUILDERS = {
    "title": lambda prs, s: add_title_slide(prs, s["title"], s.get("subtitle", ""), s.get("notes", "")),
    "bullets": lambda prs, s: add_bullets_slide(prs, s["title"], s.get("bullets", []), s.get("notes", "")),
    "two-column": lambda prs, s: add_two_column_slide(
        prs,
        s["title"],
        s.get("left_title", ""),
        s.get("left_bullets", []),
        s.get("right_title", ""),
        s.get("right_bullets", []),
        s.get("notes", ""),
    ),
}


def render(spec: dict, output_dir: str = "output"):
    """Render a parsed spec into a PowerPoint file."""
    metadata = spec["metadata"]
    slides = spec["slides"]
    out_name = metadata.get("output", "presentation.pptx")

    prs = Presentation()

    for slide_data in slides:
        stype = slide_data["type"]
        builder = SLIDE_BUILDERS.get(stype)
        if builder is None:
            print(f"Warning: unknown slide type '{stype}', skipping.")
            continue
        builder(prs, slide_data)

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, out_name)
    prs.save(out_path)
    print(f"Saved {len(slides)} slides -> {out_path}")
    return out_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Generate a PowerPoint presentation from a spec file.")
    parser.add_argument("spec", help="Path to the .spec.md file")
    parser.add_argument("-o", "--output-dir", default="output", help="Output directory (default: output)")
    args = parser.parse_args()

    if not os.path.isfile(args.spec):
        sys.exit(f"Error: spec file not found: {args.spec}")

    spec = parse_spec(args.spec)
    render(spec, args.output_dir)


if __name__ == "__main__":
    main()
