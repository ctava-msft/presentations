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

    elif slide_type == "section-header":
        sub_match = re.search(r"\*\*Subtitle\*\*\s*:\s*(.+)", rest)
        slide["subtitle"] = sub_match.group(1).strip() if sub_match else ""

    elif slide_type == "content":
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

    # Extract left bullets (between **Left**: and **Right**: or **Notes**: or end)
    left_match = re.search(
        r"\*\*Left\*\*\s*:\s*\n(.*?)(?=\*\*Right\*\*|\*\*Notes\*\*|$)",
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
    """Layout 0 – Title Slide: large centered title + subtitle."""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle
    slide.shapes.title.text_frame.paragraphs[0].font.size = TITLE_FONT
    slide.placeholders[1].text_frame.paragraphs[0].font.size = SUBTITLE_FONT
    slide.notes_slide.notes_text_frame.text = notes


def add_content_slide(prs: Presentation, title: str, bullets: list[str], notes: str):
    """Layout 1 – Title and Content: title bar + single content placeholder."""
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


def add_section_header_slide(prs: Presentation, title: str, subtitle: str, notes: str):
    """Layout 2 – Section Header: large title + subtitle for topic transitions."""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    slide.shapes.title.text = title
    slide.shapes.title.text_frame.paragraphs[0].font.size = TITLE_FONT
    if subtitle:
        slide.placeholders[1].text = subtitle
        slide.placeholders[1].text_frame.paragraphs[0].font.size = SUBTITLE_FONT
    slide.notes_slide.notes_text_frame.text = notes


def add_two_column_slide(
    prs: Presentation,
    title: str,
    left_bullets: list[str],
    right_bullets: list[str],
    notes: str,
):
    """Layout 3 – Two Content: title + two side-by-side content placeholders."""
    slide = prs.slides.add_slide(prs.slide_layouts[3])
    slide.shapes.title.text = title
    slide.shapes.title.text_frame.paragraphs[0].font.size = HEADING_FONT

    # Left content – placeholder index 1
    left_ph = slide.placeholders[1].text_frame
    left_ph.clear()
    for i, b in enumerate(left_bullets):
        p = left_ph.paragraphs[0] if i == 0 else left_ph.add_paragraph()
        p.text = b
        p.level = 0
        p.font.size = COL_BODY_FONT

    # Right content – placeholder index 2
    right_ph = slide.placeholders[2].text_frame
    right_ph.clear()
    for i, b in enumerate(right_bullets):
        p = right_ph.paragraphs[0] if i == 0 else right_ph.add_paragraph()
        p.text = b
        p.level = 0
        p.font.size = COL_BODY_FONT

    slide.notes_slide.notes_text_frame.text = notes


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------

SLIDE_BUILDERS = {
    "title": lambda prs, s: add_title_slide(
        prs, s["title"], s.get("subtitle", ""), s.get("notes", "")),
    "content": lambda prs, s: add_content_slide(
        prs, s["title"], s.get("bullets", []), s.get("notes", "")),
    "section-header": lambda prs, s: add_section_header_slide(
        prs, s["title"], s.get("subtitle", ""), s.get("notes", "")),
    "two-column": lambda prs, s: add_two_column_slide(
        prs, s["title"], s.get("left_bullets", []), s.get("right_bullets", []),
        s.get("notes", "")),
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
    out_path = _next_version_path(output_dir, out_name)
    prs.save(out_path)
    print(f"Saved {len(slides)} slides -> {out_path}")
    return out_path


def _next_version_path(output_dir: str, filename: str) -> str:
    """Return a versioned path: filename.pptx, filename_1.pptx, filename_2.pptx, …"""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(output_dir, filename)
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(output_dir, f"{base}_{n}{ext}")
        n += 1
    return candidate


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
