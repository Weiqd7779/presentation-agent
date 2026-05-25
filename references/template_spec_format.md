# Template Spec 格式

這個檔案定義 Stage 1 產生 `<template_name>_spec.md` 時應該遵守的格式。

`placeholders.txt` 是可注入 TextBox 名稱的唯一 source of truth。`spec.md` 的角色是輔助 Stage 2 做 mapping，不是取代 XML 抽取結果。

## 必要結構

```markdown
# Presentation Analysis Specification: [Template Name]

## Overall Structure

- Total slides: [投影片總數]
- Theme: [色彩、字體、視覺風格]
- Template type: [corporate/proposal/training/portfolio/etc.]

## Slide Analysis

### Slide [N]: [簡短用途]

- Layout: [封面、單欄內容、雙欄、數據卡、結尾頁等]
- Editable text boxes:
  - `[Exact TextBox name]`: [目前文字，以及此框適合放什麼內容]
  - `[Exact TextBox name]`: [目前文字，以及此框適合放什麼內容]
- Static visual elements:
  - [背景、照片、logo、裝飾圖形等不應任意改動的元素]
- Mapping notes:
  - [建議哪種內容放入哪個 TextBox]
  - [長度限制、換行風險、可能溢位的地方]
- Risks:
  - [Stage 2 注入時最容易造成的排版問題]

## Mapping Guidance

- Suggested slide order: [建議順序，或 keep original]
- Best reusable layouts:
  - Slide [N]: [適合重複使用的情境]
- Avoid editing:
  - [footer/date/slide-number/decorative shapes 等不建議注入內容的位置]
```

## 規則

- TextBox 名稱必須完全照 `placeholders.txt`，例如 `TextBox 7`。
- 不要發明 placeholder 名稱。
- 要清楚分開「可編輯文字框」與「靜態視覺元素」。
- Mapping notes 要能直接幫 Stage 2 產生 `mapping.json` 或 `CONTENT_MAP`。
- 如果分析中文內容，請補充中文長度建議，例如封面大標適合 4-8 個中文字。
