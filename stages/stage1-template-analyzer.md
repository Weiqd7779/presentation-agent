# Stage 1: Template Visual and Structural Analysis

## Goal

Analyze a PPTX template and produce reliable artifacts for Stage 2 content injection. The most important output is the exact list of editable text boxes.

## Command

Always quote paths:

```powershell
uv run python scripts/orchestrator.py analyze "<template_file.pptx>" "<output_dir>"
```

## What the command does

1. Unpacks the PPTX into `<output_dir>/unpacked`.
2. Extracts text box names and current text into `placeholders.txt`.
3. Writes `.source_pptx` so Stage 2 can reset from the original template.
4. Converts the PPTX to PDF with LibreOffice.
5. Converts PDF pages to slide JPGs and creates a slide overview image.
6. Extracts template text with `markitdown`.
7. Runs the visual analyzer to create `<template_name>_spec.md`.

## Outputs

- `placeholders.txt`: source of truth for exact editable shape names.
- `placeholder_layout.json`: deterministic XML geometry and fallback role hints for editable text boxes and pictures.
- `placeholder_roles.json`: machine-readable role map, including Vision status per shape.
- `annotated_slides/`: per-slide images with bounding boxes and exact shape names overlaid.
- `spec_slides/`: per-slide visual analysis cache.
- `<template_name>_text_direct.md`: direct PPTX text extraction, useful when `markitdown` produces mojibake for localized decks.
- `<template_name>_spec.md`: advisory visual analysis and mapping guidance.
- `<template_name>_slide_overview.jpg`: overview image for layout decisions.
- `<template_name>_text.md`: extracted template text.
- `unpacked/`: unpacked PPTX package used by Stage 2.

## Checks

Before moving to Stage 2:

- Confirm `placeholders.txt` exists and includes all expected slides.
- For picture replacement, confirm `placeholders.txt` lists the expected picture names under each slide.
- For Canva or exported templates, inspect `annotated_slides/` and `placeholder_roles.json` to map generic names such as `TextBox 17` or `Picture 4` to their visual role.
- Confirm the overview image exists.
- Prefer exact TextBox names from `placeholders.txt` over any LLM-generated description.
- If localized text appears corrupted in `*_text.md`, check `*_text_direct.md` before assuming the PPTX is corrupt.
- If `*_spec.md` contains unreadable text, continue with `placeholders.txt` and the overview image.

## Vision Analysis Batching

Stage 1 sends annotated rendered slide images to the Vision LLM one slide at a time. Defaults:

- `PRESENTATION_AGENT_VISION_TIMEOUT_MS=300000`
- `PRESENTATION_AGENT_SKIP_VISION=0`

Increase the timeout for slow providers. If the provider is unreliable, set `PRESENTATION_AGENT_SKIP_VISION=1`; Stage 1 will still produce XML geometry, annotated images, fallback specs, and role hints.

## Known Warnings

`MuPDF error: format error: No common ancestor in structure tree` can appear during PDF-to-image conversion. It is usually non-fatal if JPG files were still created.
