# Stage 3: Visual QA and Finalize

## Goal

Inspect the draft deck visually, fix layout issues, and create the final PPTX only after QA passes.

## Audit Command

```powershell
uv run python scripts/orchestrator.py audit "<output_dir>\output_draft.pptx" "<output_dir>"
```

The audit command:

1. Converts `output_draft.pptx` to PDF.
2. Converts the PDF to slide images.
3. Sends slide images to the configured Vision LLM.
4. Writes `qa/vision_qa_report.json`.

## QA Report

Review:

```text
<output_dir>\qa\vision_qa_report.json
```

Expected fields:

- `pass`: boolean
- `overall_summary`: short summary
- `blocking_issue_count`: number of high-confidence delivery-blocking issues
- `issues`: list of slide-level issues

Issue types include:

- `TEXT_OVERFLOW`
- `MAPPING_COLLISION`
- `LAYOUT_DISPLACEMENT`
- `FONT_INCONSISTENCY`
- `BROKEN_RENDER`

Each issue should include `severity`, `confidence`, `blocking`, `evidence`, `repair_hint`, and `verification`. The audit normalizes missing fields and only treats an issue as blocking when it is high severity, explicitly blocking, and confidence is at least `0.70`.

## Iteration Rule

If `blocking_issue_count` is greater than `0`:

1. Fix `mapping.json` or `inject_content.py`.
2. Rerun Stage 2.
3. Rerun Stage 3.

Medium and low severity issues are advisory. Do not keep iterating on subjective polish unless the user asks for it.

Do not finalize a normal run while blocking QA issues remain.

## Finalize Command

Run only after QA passes:

```powershell
uv run python scripts/orchestrator.py finalize "<output_dir>"
```

The command checks `qa/vision_qa_report.json` and blocks if QA did not pass.

If the user explicitly wants a packaged file despite failing QA:

```powershell
uv run python scripts/orchestrator.py finalize "<output_dir>" --force
```

Finalize outputs:

- `output_final.pptx`

## Provider Setup

Stage 3 requires `.vision_llm_config.json` in the skill directory. Prefer environment variables for API keys. Keep real credentials out of version control.
