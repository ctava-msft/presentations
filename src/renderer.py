"""Renderer – orchestrates parsing, enrichment, image generation, and slide building."""

from __future__ import annotations

import os

from pptx import Presentation

from .animations import apply_animations
from .enrichment import enrich_content_from_urls, enrich_notes_from_urls
from .images import resolve_image_prompt
from .slides import SLIDE_BUILDERS
from .style import Style

# ---------------------------------------------------------------------------
# Versioned output path
# ---------------------------------------------------------------------------


def _next_version_path(output_dir: str, filename: str) -> str:
    """Return a versioned path: ``file.pptx``, ``file_1.pptx``, …"""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(output_dir, filename)
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(output_dir, f"{base}_{n}{ext}")
        n += 1
    return candidate


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def render(
    spec: dict,
    output_dir: str = "output",
    image_model: str | None = None,
) -> str:
    """Render a parsed spec into a PowerPoint file and return the output path."""
    metadata = spec["metadata"]
    slides = spec["slides"]
    out_name = metadata.get("output", "presentation.pptx")

    # Image model priority: CLI flag > front-matter
    default_model = (
        image_model
        or metadata.get("image_model", "").strip().lower()
    )
    if not default_model:
        print(
            "Warning: no image_model specified in front matter or --image-model flag. "
            "ImagePrompt directives will be skipped."
        )

    # Text model for note enrichment: front-matter > env var > default
    text_model = metadata.get("text_model", "").strip()

    # Build Style from front-matter
    style = Style(metadata.get("style"))

    prs = Presentation()

    for slide_data in slides:
        stype = slide_data["type"]
        builder = SLIDE_BUILDERS.get(stype)
        if builder is None:
            print(f"Warning: unknown slide type '{stype}', skipping.")
            continue

        # Enrich slide content from ContentUrls before building
        enrich_content_from_urls(slide_data, text_model=text_model)

        # Enrich notes from ContentUrls before building
        enrich_notes_from_urls(slide_data, text_model=text_model)

        # Resolve any ImagePrompt → generate images before building
        resolve_image_prompt(slide_data, output_dir, default_model=default_model)

        builder(prs, slide_data, style, apply_animations=apply_animations)

    os.makedirs(output_dir, exist_ok=True)
    out_path = _next_version_path(output_dir, out_name)
    prs.save(out_path)
    print(f"Saved {len(slides)} slides -> {out_path}")
    return out_path
