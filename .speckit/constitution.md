# Microsoft Presentation Constitution

## Project Overview

This constitution establishes the principles, standards, and governance framework for creating professional presentations using specification-driven automation. Presentations are generated from structured markdown specification files, ensuring consistency, reusability, and quality across all presentation deliverables.

**Mission**: Enable scalable, high-quality presentation creation where content is authored as structured specifications and rendered into polished slide decks via automation—eliminating manual formatting and ensuring brand/message consistency.

## Core Principles

### 1. Specification-Driven Content

- **Content as Code**: All presentation content lives in versioned `.spec.md` files, never hardcoded in scripts
- **Separation of Concerns**: Content (spec) is decoupled from rendering (script) and styling (theme)
- **Single Source of Truth**: Each presentation topic has exactly one spec file that defines all slides, notes, and metadata
- **Reproducibility**: Running the same spec always produces the same presentation

### 2. Audience-Centered Design

- **Know Your Audience**: Every presentation spec must declare its target audience (business, technical, executive, mixed)
- **Appropriate Depth**: Content depth matches audience—avoid jargon for business audiences, avoid oversimplification for technical audiences
- **Actionable Takeaways**: Every presentation must end with clear, actionable next steps
- **Speaker Notes**: Every slide must include speaker notes that guide delivery and provide context beyond bullet points

### 3. Microsoft Brand & Messaging Standards

- **Consistent Voice**: Professional, clear, confident—never hyperbolic or speculative
- **Balanced Perspective**: Present capabilities alongside limitations and risks
- **Responsible AI Framing**: When covering AI topics, always include governance, guardrails, and human oversight
- **Customer Value First**: Lead with business value and outcomes, not technology features

### 4. Structured Slide Types

All presentations use a defined set of slide types to maintain consistency:

| Slide Type | Layout | Purpose | Required Fields |
|------------|--------|---------|------------------|
| `title` | 0 – Title Slide | Opening slide with title and subtitle | title, subtitle, notes |
| `content` | 1 – Title and Content | Key points with supporting notes | title, bullets (list), notes |
| `section-header` | 2 – Section Header | Topic transition with title and subtitle | title, subtitle, notes |
| `two-column` | 3 – Two Content | Side-by-side comparison or contrast | title, left_bullets, right_bullets, notes |

### 5. Quality Standards

- **Slide Count**: Target 15–25 slides for a 60-minute presentation; adjust proportionally
- **Bullet Count**: 3–5 bullets per slide; avoid walls of text
- **Bullet Length**: Each bullet should be one concise line (under 80 characters when possible)
- **Notes Depth**: Speaker notes should be 2–5 sentences providing delivery guidance, examples, and transitions
- **Flow**: Slides should follow a logical narrative arc: context → education → application → action

## Specification Format

### File Structure

Spec files use markdown with YAML front matter and structured slide sections:

```
---
title: Presentation Title
subtitle: Descriptive subtitle for the audience
output: Output_Filename.pptx
author: Author Name
duration: 60
audience: business, IT
---

## [title] Slide Title

**Subtitle**: A descriptive subtitle

**Notes**: Speaker notes for delivery guidance.

---

## [content] Slide Title

- Bullet point one
- Bullet point two
- Bullet point three

**Notes**: Speaker notes for delivery guidance.

---

## [section-header] Section Title

**Subtitle**: Optional subtitle for the section transition

**Notes**: Speaker notes for delivery guidance.

---

## [two-column] Slide Title

**Left**:
- Left bullet one
- Left bullet two

**Right**:
- Right bullet one
- Right bullet two

**Notes**: Speaker notes for delivery guidance.
```

### YAML Front Matter

| Field | Required | Description |
|-------|----------|-------------|
| `title` | Yes | Presentation title |
| `subtitle` | Yes | Subtitle / tagline |
| `output` | Yes | Output filename (`.pptx`) |
| `author` | No | Author or team name |
| `duration` | No | Target duration in minutes |
| `audience` | No | Target audience types |

### Slide Definitions

Each slide is separated by `---` and begins with `## [type] Title`:

- **Title slides** (`[title]`): Include a `**Subtitle**:` line
- **Content slides** (`[content]`): Use markdown list items (`- `)
- **Section Header slides** (`[section-header]`): Include an optional `**Subtitle**:` line for topic transitions
- **Two-column slides** (`[two-column]`): Include `**Left**:` (bullets) and `**Right**:` (bullets)
- **All slides**: Must include a `**Notes**:` section

## Development Standards

### Script Architecture

- The generation script accepts a spec file path as a command-line argument
- Output is written to the `output/` directory using the filename from the spec's `output` field
- The script is content-agnostic—it renders any valid spec into a presentation
- No presentation content is hardcoded in the script

### Dependency Management

- `python-pptx` for PowerPoint generation
- `pyyaml` for YAML front matter parsing
- Minimal dependencies; prefer standard library where possible
- All dependencies tracked in `requirements.txt`

### Error Handling

- Validate spec structure before rendering
- Provide clear error messages for malformed specs
- Fail gracefully with actionable guidance

## Governance

### Content Review

- Spec files should be reviewed for accuracy, messaging, and audience appropriateness before generation
- Speaker notes should be reviewed for delivery quality
- Generated presentations should be spot-checked for formatting

### Version Control

- All spec files are version controlled
- Changes to specs are tracked with meaningful commit messages
- Major content revisions warrant version bumps in the spec metadata

### Reusability

- Common slide patterns (e.g., Responsible AI, guardrails) can be extracted into reusable spec fragments
- The generation script supports composing presentations from multiple spec sources in future iterations

## Usage

```bash
# Generate a presentation from a spec file
python presentation.py .speckit/specifications/ai101.spec.md

# Output is written to: output/<output-filename-from-spec>.pptx
```

## Success Criteria

- ✅ Any team member can create a professional presentation by writing a spec file
- ✅ No manual PowerPoint editing required for standard presentations
- ✅ Content changes require only spec file updates, not script changes
- ✅ Consistent quality and formatting across all generated presentations
- ✅ Speaker notes provide genuine delivery value, not just bullet restatements
