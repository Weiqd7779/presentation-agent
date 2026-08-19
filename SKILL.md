---
name: presentation-agent
description: Create, transform, and validate PowerPoint decks from a PPTX template plus Markdown/script content. Use when Codex needs to analyze a presentation template, extract exact placeholders, map content into slides, build a draft PPTX, run visual QA, fix layout issues, or finalize a deck.
---

# Presentation Agent

Use this skill to turn a PPTX template and user-provided content into a finished presentation through three stages:

1. Analyze the template and extract exact editable text boxes.
2. Design the slide structure, merge content into the template, and build a draft deck.
3. Run visual QA, iterate if needed, then finalize only after QA passes.

Always quote Windows paths in commands because template filenames often contain spaces.

## Required Inputs

- A `.pptx` template file.
- A Markdown/script file that describes the desired deck content.
- An output directory for generated artifacts.

If the user has not provided a template, offer the bundled templates in `templates/`:

- `template_a_corporate.pptx`
- `template_b_proposal.pptx`
- `template_c_training.pptx`

## Stage 1: Template Analysis

Read [stages/stage1-template-analyzer.md](stages/stage1-template-analyzer.md), then run:

```powershell
uv run python scripts/orchestrator.py analyze "<template_file.pptx>" "<output_dir>"
```

Primary Stage 1 outputs:

- `placeholders.txt`: exact shape names and current text. Treat this as the source of truth.
- `placeholder_layout.json`: deterministic XML geometry for each editable text box and picture, including position, size, region, and a fallback role hint.
- `placeholder_roles.json`: machine-readable role map from Stage 1 visual analysis, with deterministic fallback data when Vision analysis times out.
- `annotated_slides/`: rendered slide images with bounding boxes and exact shape names overlaid. Use these for Canva or exported templates with generic shape names.
- `spec_slides/`: per-slide analysis cache. Re-run only failed or changed slides when needed.
- `*_text_direct.md`: text extracted directly from the PPTX with `python-pptx`; use this when `markitdown` output is mojibake for localized decks.
- `<template_name>_spec.md`: LLM-assisted layout notes. Treat this as advisory.
- `<template_name>_slide_overview.jpg`: visual overview for layout decisions.
- `<template_name>_text.md`: extracted original template text.
- `.source_pptx`: pointer used by Stage 2 to reset the workspace before building.

## Stage 2: Content Merge

Read [stages/stage2-content-merger.md](stages/stage2-content-merger.md).

Preferred path: create a `mapping.json` in the output directory and build from data. Before writing the mapping, decide whether to keep, duplicate, delete, or reorder slides. Do not default to the original slide sequence when the content would be clearer with repeated layouts or fewer slides.

Preflight the mapping before building:

```powershell
uv run python scripts/orchestrator.py preflight "<output_dir>" "<output_dir>\mapping.json"
```

```powershell
uv run python scripts/orchestrator.py build-json "<output_dir>" "<output_dir>\mapping.json"
```

Use this mapping shape:

```json
{
  "slide_order": [1, 2, 5, 5, 10],
  "content_map": {
    "1": {
      "TextBox 7": "Deck title",
      "TextBox 8": "Subtitle"
    },
    "2": {
      "TextBox 6": "Agenda",
      "TextBox 7": ["First point", "Second point", "Third point"]
    },
    "3": {
      "TextBox 17": "Case Study 1"
    },
    "4": {
      "TextBox 17": "Case Study 2"
    }
  }
}
```

`slide_order` uses 1-based original slide numbers. Repeating a number duplicates that original slide; omitting a number deletes that original slide. `content_map` keys always refer to final slide numbers after this structure change.

Fallback path: copy `scripts/stage2/inject_content_template.py` to `<output_dir>/inject_content.py`, edit `SLIDE_ORDER` and `CONTENT_MAP`, then run:

```powershell
uv run python scripts/orchestrator.py build "<output_dir>" "<output_dir>\inject_content.py"
```

After Stage 2, inspect `output_draft_text.md` for missing content and obvious mapping mistakes.
If `output_draft_text.md` contains mojibake for non-English text, inspect `output_draft_text_direct.md` instead.
Stage 2 also validates `output_draft.pptx` through the PowerPoint DOM and records the result in `powerpoint_validation_draft.json`.

## Stage 3: Visual QA

Read [stages/stage3-visual-qa.md](stages/stage3-visual-qa.md), then run:

```powershell
uv run python scripts/orchestrator.py audit "<output_dir>\output_draft.pptx" "<output_dir>"
```

Review `qa/vision_qa_report.json`.

- If `blocking_issue_count` is greater than `0`, fix the mapping/content and rerun Stage 2 and Stage 3.
- Medium and low severity visual issues are advisory; do not keep iterating unless they affect delivery.
- If `pass` is `true`, finalize:

```powershell
uv run python scripts/orchestrator.py finalize "<output_dir>"
```

Finalize validates `output_final.pptx` through the PowerPoint DOM and records the result in `powerpoint_validation_final.json`.

If a user explicitly asks to package despite failing QA, use:

```powershell
uv run python scripts/orchestrator.py finalize "<output_dir>" --force
```

## Operating Rules

- Treat `placeholders.txt` as the authoritative list of editable shape names.
- Never invent TextBox names. Use exact names from Stage 1.
- Use `slide_order` as a design tool when the user's content needs duplicated layouts, removed sections, or a stronger narrative flow.
- Do not edit decorative placeholders such as dates, slide numbers, footers, or headers unless explicitly requested.
- Keep `mapping.json` or `inject_content.py` in the output directory, not inside the skill source.
- Treat `mapping_preflight_report.json`, `powerpoint_validation_draft.json`, `qa/vision_qa_report.json`, and `powerpoint_validation_final.json` as delivery gates.
- Check `run_manifest.json` before reporting completion; it records the latest gate status and artifact paths.
- Do not finalize a normal run until visual QA passes.
- Keep local Vision LLM credentials outside version control. Use `.vision_llm_config.template.json` as the template.

## Complex Custom Template Guidance

For complex decks, first create the PPTX template as a finished visual system, then use this skill only to analyze, map, build, audit, and finalize the content-filled deck.

- Put every editable content area in a named text box with meaningful placeholder text.
- Build diagrams, tables, cards, timelines, and project layouts directly into the template; Stage 2 should replace text, reorder slides, duplicate layouts, or replace existing pictures, not design new geometry from scratch.
- Prefer stable semantic shape names such as `S04_METRIC_1_VALUE` or `S08_POINT_3_BODY` when authoring templates. Stage 1 will preserve these names in `placeholders.txt`.
- Use ASCII placeholder text if the template-generation path may corrupt non-ASCII text. Inject localized final copy through UTF-8 `mapping.json` and verify the final PPTX directly when text extraction tools show mojibake.
- Treat Stage 1 visual analysis as advisory. If the Vision LLM times out, continue from `placeholders.txt`, rendered slide images, and Stage 3 QA.
- Stage 1 visual analysis sends annotated rendered slides one at a time instead of sending the whole deck overview at once. Tune with `PRESENTATION_AGENT_VISION_TIMEOUT_MS`; the default is 300000 ms per slide.
- `placeholders.txt` includes both editable text boxes and picture names for `image_map`. It also warns about generic shape names so template authors can rename them before mapping.
- For Canva or exported templates with generic names, do not rely on names alone. Use `annotated_slides/`, `placeholder_layout.json`, and `placeholder_roles.json` to identify where each `TextBox N` or `Picture N` appears and what role it likely serves.
- If you need a deterministic Stage 1 run without Vision calls, set `PRESENTATION_AGENT_SKIP_VISION=1`. The output will still include geometry, annotated slides, fallback specs, and role hints.

## Troubleshooting

- If `quick_validate.py` fails on Windows with an encoding error, rerun with UTF-8 enabled:

```powershell
$env:PYTHONUTF8="1"; uv run python <quick_validate.py> "<skill_dir>"
```

- If Stage 1 fails with "unrecognized arguments", quote the template and output paths.
- If MuPDF prints `No common ancestor in structure tree`, check whether slide images were still produced. This warning is common for LibreOffice-generated PDFs and is usually non-fatal.
- If Stage 3 cannot call a provider, check `.vision_llm_config.json` and the required environment variable for the selected provider.
