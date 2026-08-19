# Stage 2: Content Merger

## Goal

Design the deck structure, then inject user content into the analyzed template without corrupting the PPTX. Use exact shape names from `placeholders.txt`.

Stage 2 is not only text replacement. Before writing `mapping.json`, decide whether the content needs the original slide count and order. Use the template as a design system: duplicate useful layouts, remove irrelevant slides, and reorder sections when that creates a clearer narrative.

## Structure Planning

Before writing `mapping.json`, choose the deck structure deliberately:

- Keep the original order with `null` only when every template slide has a clear purpose.
- Duplicate an original slide when one layout fits several content sections. Example: use one project/case-study layout for three projects.
- Delete slides that do not match the user's content, audience, or requested story.
- Reorder slides to improve flow, such as cover -> problem -> approach -> evidence -> closing.

After this decision, encode the structure in `slide_order` and write content against the final slide numbers.

## Preferred Input: mapping.json

Create `<output_dir>/mapping.json`. Prefer an explicit `slide_order` for real deck work; use `null` only after deciding the original template structure already fits.

```json
{
  "slide_order": [1, 2, 5, 5, 5, 7, 10],
  "content_map": {
    "1": {
      "TextBox 7": "Portfolio",
      "TextBox 8": "AI Engineering",
      "TextBox 9": {
        "text": "Fixed-size title",
        "font_size": 18
      }
    },
    "3": {
      "TextBox 17": "Project 1",
      "TextBox 18": "Agent routing and observability"
    },
    "4": {
      "TextBox 17": "Project 2",
      "TextBox 18": "RAG knowledge workflow"
    },
    "5": {
      "TextBox 17": "Project 3",
      "TextBox 18": "Recommendation prototype"
    },
    "6": {
      "TextBox 15": "Evidence",
      "TextBox 16": "Measured outcomes and lessons learned"
    },
    "7": {
      "TextBox 7": "Thank You",
      "TextBox 8": "Contact"
    }
  },
  "image_map": {
    "1": {
      "Picture 4": {
        "path": "C:\\path\\to\\image.png",
        "fill": "contain",
        "background": "white"
      }
    }
  }
}
```

With `slide_order: [1, 2, 5, 5, 5, 7, 10]`:

| Final slide | Uses original slide |
|---:|---:|
| 1 | 1 |
| 2 | 2 |
| 3 | 5 |
| 4 | 5 |
| 5 | 5 |
| 6 | 7 |
| 7 | 10 |

Rules:

- `slide_order` is `null` to keep the original order.
- `slide_order` can be a list of 1-based original slide numbers. Repetition duplicates slides; omission deletes slides.
- `content_map` keys are 1-based final slide numbers after reordering.
- When a slide is duplicated, map each duplicated final slide separately.
- If a final slide is omitted from `content_map`, it remains in the deck with its original template text.
- To delete a slide, omit its original slide number from `slide_order`; do not create an empty `content_map` entry.
- Shape keys must be exact names from `placeholders.txt`.
- For duplicated slides, use the shape keys from the original slide that was duplicated.
- A string replaces the first paragraph.
- A list creates multiple paragraphs in the same shape.
- A text object with `text` plus `font_size`, `min_font_size`, or `max_font_size` controls sizing for tight layouts and repeated same-role text.
- `image_map` replaces an existing picture by exact picture name while preserving layout. Use it for required subject or brand imagery instead of manually overwriting `ppt/media` files.

Run:

```powershell
uv run python scripts/orchestrator.py preflight "<output_dir>" "<output_dir>\mapping.json"
```

```powershell
uv run python scripts/orchestrator.py build-json "<output_dir>" "<output_dir>\mapping.json"
```

`build-json` reruns the same preflight before changing the unpacked PPTX, so mapping errors fail before slide XML is modified.
The preflight fully verifies `content_map` and `image_map` names against `placeholders.txt`. `table_map` is slide-range checked and reported as advisory because table names are not currently listed in `placeholders.txt`.

## Fallback Input: inject_content.py

If custom Python logic is required:

1. Copy `scripts/stage2/inject_content_template.py` to `<output_dir>/inject_content.py`.
2. Fill in `SLIDE_ORDER` and `CONTENT_MAP`.
3. Run:

```powershell
uv run python scripts/orchestrator.py build "<output_dir>" "<output_dir>\inject_content.py"
```

## Build Outputs

- `output_draft.pptx`
- `output_draft_text.md`
- `output_draft_text_direct.md`

## Review Checklist

After build:

- Read `output_draft_text.md`.
- If localized text in `output_draft_text.md` appears as mojibake, read `output_draft_text_direct.md`.
- Confirm `mapping_preflight_report.json` passed.
- Confirm `powerpoint_validation_draft.json` passed and contains a readable PowerPoint slide count.
- Confirm every intended content section appears.
- Confirm the slide count and order match the intended story.
- Look for leftover template text, `[Title]`, `{{Name}}`, `Lorem ipsum`, or similar placeholders.
- Note any `[WARN] TextBox ... not found` messages and fix the mapping.
- Treat automatic font scaling messages as a signal to review the slide visually.

## Common Problems

- Wrong slide number: `content_map` uses final slide numbers, not original numbers after `slide_order` is applied.
- Duplicated slide not customized: every duplicated final slide needs its own `content_map` entry, otherwise template text may remain.
- Wrong shape name: use exact names from `placeholders.txt`.
- Overlong text: shorten content before relying on auto-scaling.
- Decorative footer/date/slide-number placeholders: do not map content into them.
