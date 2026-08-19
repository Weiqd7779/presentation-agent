# Presentation Agent

Presentation Agent 是一套給 Codex 或其他自動化 Agent 使用的 PowerPoint 製作流程。它不從空白畫布任意畫投影片，而是讀取既有 `.pptx` 範本，找出真正可編輯的文字框與圖片，再把內容安全地映射進去。

這個做法適合公司提案、教育訓練與固定品牌簡報：設計師先把版型做好，Agent 負責理解內容、選擇版型、填入文字、檢查結果，而不是每次重新猜字型、色彩與版面。

## 它解決什麼問題

直接修改 PowerPoint XML 很容易遇到幾種失敗：

- Agent 猜錯 `TextBox 7` 之類的 shape 名稱，內容根本沒有寫進去；
- 內容塞進正確位置，但文字溢出或版面重疊；
- 刪除或複製投影片後，關聯檔殘留，PowerPoint 開啟時要求修復；
- 其他工具看起來正常，實際用 PowerPoint 開啟後卻換字型或重新換行；
- 成品沒有明確的品質閘門，流程跑完就被誤認為完成。

Presentation Agent 把這些風險拆成三個階段，並留下可以查驗的 JSON 報告。

## 工作流程

### 1. 分析範本

系統先解開 PPTX、擷取文字框和圖片名稱、記錄幾何位置，並產生加框標註的投影片圖片。`placeholders.txt` 是後續映射的唯一 shape 名稱來源。

主要輸出包括：

- `placeholders.txt`：每張投影片可編輯的文字框與圖片名稱；
- `placeholder_layout.json`：位置、大小與區域；
- `placeholder_roles.json`：視覺角色推測與 deterministic fallback；
- `annotated_slides/`：標出 shape 名稱的投影片圖片。

### 2. 映射內容並建立草稿

使用者或 Agent 在 `mapping.json` 決定：

- 哪些原始投影片要保留、刪除、重排或重複；
- 每個文字框要放什麼內容；
- 哪些既有圖片要替換。

建立前會先執行 preflight，確認投影片編號與 shape 名稱真的存在。建立後會檢查 PPTX 結構、抽出文字，再透過 PowerPoint DOM 實際開啟草稿並讀取投影片數量。

### 3. 視覺 QA 與定稿

草稿會轉成圖片，交給設定的 Vision 模型檢查重疊、溢出、可讀性、表格與殘留 placeholder。只有 QA 通過後才能正常 finalize；定稿還會再跑一次 PowerPoint DOM 驗證。

`run_manifest.json` 會記錄 preflight、草稿驗證、視覺 QA 與定稿驗證的最新狀態。

## 使用前準備

| 需求 | 用途 |
|---|---|
| Python 3.11+ 與 `uv` | 執行 Python 工具與安裝相依套件 |
| Node.js 18+ | 呼叫 Vision provider |
| LibreOffice | 將投影片轉成 PDF／圖片供分析與 QA |
| Windows PowerPoint | 以實際 PowerPoint DOM 驗證檔案；完整流程的必要閘門 |
| 一個 Vision provider | Stage 3 視覺 QA；支援設定檔列出的 OpenAI、Claude、Gemini、Fireworks 或自架 endpoint |

安裝 Python 相依套件：

```powershell
uv sync
```

接著複製 `.vision_llm_config.template.json` 為不受版本控制的 `.vision_llm_config.json`，選擇 provider，並把 API key 放在對應的環境變數。設定檔只保存「要讀哪個環境變數」，不要把真實 key 寫進 repository。

## 完整使用範例

假設你有一份 `proposal-template.pptx`，並要把內容產生到 `work/proposal`：

```powershell
uv run python scripts/orchestrator.py analyze "proposal-template.pptx" "work/proposal"
```

分析完成後，先閱讀 `work/proposal/placeholders.txt` 與加框圖片，再建立 `work/proposal/mapping.json`。不要憑印象猜 shape 名稱。

```powershell
uv run python scripts/orchestrator.py preflight "work/proposal" "work/proposal/mapping.json"
uv run python scripts/orchestrator.py build-json "work/proposal" "work/proposal/mapping.json"
uv run python scripts/orchestrator.py audit "work/proposal/output_draft.pptx" "work/proposal"
uv run python scripts/orchestrator.py finalize "work/proposal"
```

正常完成後，最重要的成品是：

- `output_final.pptx`：定稿簡報；
- `qa/vision_qa_report.json`：視覺檢查結果；
- `powerpoint_validation_final.json`：PowerPoint 實際開啟結果；
- `run_manifest.json`：整次流程的閘門摘要。

如果使用者沒有自己的範本，可以從 `templates/` 的企業簡報、提案或教育訓練範本開始。

## 安全與資料邊界

- `.vision_llm_config.json`、`.env*`、輸出目錄與快取都已排除版本控制。
- 範本投影片在視覺分析與 QA 階段會送到你設定的 Vision provider；含機密內容的簡報不應使用未經核准的外部 provider。
- repository 只包含通用程式、三份可重用範本與設定範例，不包含生成簡報、客戶內容、品牌專用素材、字型包或真實憑證。
- `--force` 只適合使用者明確要求在 QA 未通過時仍封裝檔案；正常流程不應繞過品質閘門。

## 專案結構

- `SKILL.md`：給 Codex Agent 的主操作規則。
- `stages/`：三個階段的詳細執行與驗收說明。
- `scripts/orchestrator.py`：分析、preflight、建立、QA 與定稿入口。
- `scripts/stage1/`：範本解析、渲染與幾何標註。
- `scripts/stage2/`：內容映射、PPTX 打包與結構檢查。
- `scripts/vision_qa.py`：視覺 QA。
- `scripts/validate_powerpoint.py`：PowerPoint DOM 驗證。
- `templates/`：三份通用起始範本。

這個 GitHub 版本刻意不包含執行輸出、實驗 benchmark、特定客戶範本、字型素材、內部開發筆記或本機 Agent 設定。
