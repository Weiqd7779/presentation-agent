# PPTX 幻燈片複製毀損事件調查報告

**日期：** 2026-04-11
**受影響元件：** `pptx_utils.py` → `reorder_slides()`
**嚴重性：** Critical — 產出的 `.pptx` 無法被 PowerPoint 開啟（錯誤碼 `0x80070570 ERROR_FILE_CORRUPT`）

---

## 一、為什麼發生

### 1.1 PPTX 的檔案結構

PPTX 本質是一個 ZIP 壓縮檔。解壓後的目錄結構關鍵如下：

```
ppt/
  presentation.xml          ← 投影片清單（sldIdLst）
  _rels/
    presentation.xml.rels   ← rId → slide 檔案路徑映射
  slides/
    slide1.xml              ← 每張投影片的內容 XML
    _rels/
      slide1.xml.rels       ← 該投影片引用的資源（layout、圖片、圖表…）
  diagrams/
    data1.xml               ← SmartArt / 圖表資料（每份必須唯一）
    drawing1.xml
    colors1.xml
    layout1.xml
    quickStyle1.xml
[Content_Types].xml         ← 每個 Part 的 MIME 類型宣告
```

每張投影片透過 `.rels` 指向它用到的所有資源。**PowerPoint 對每個資源的擁有關係有強烈假設：一個 SmartArt 資料檔只能被一張投影片引用。**

### 1.2 原始 `reorder_slides()` 做了什麼

舊實作只修改 `presentation.xml` 裡的 `<p:sldIdLst>`，僅重排 XML 指針，**磁碟上的 `slide*.xml` 實體檔案完全沒有移動或複製**。

當 `SLIDE_ORDER = [3, 7, 12, 6, 6, 7, 9, 4, 8, 13]`（slide 6 出現兩次）時：

```
sldIdLst:
  <sldId r:id="rId5"/>   → slides/slide6.xml  ← position 4
  <sldId r:id="rId5"/>   → slides/slide6.xml  ← position 5 (同一個 rId！)
```

兩個 `<sldId>` 指向同一個 rId、同一個實體檔案。PowerPoint 內部的 Part 解析器要求每個 Part 只能出現一次，發現重複立刻標記為毀損。

### 1.3 重寫後仍毀損的原因：路徑解析錯誤 + 遺漏關係類型

重寫版本加入了 `_deep_copy_slide_resources()` 準備深複製圖表檔，但犯了兩個獨立的程式錯誤：

**Bug A — 路徑解析方向錯誤**

`.rels` 檔案的 Target 屬性是相對於**投影片 XML 本身**的路徑，例如：

```
slide6.xml 位於: ppt/slides/slide6.xml
Target = "../diagrams/data1.xml"
→ 解析為: ppt/diagrams/data1.xml  ✓
```

程式碼卻從 `.rels` 檔案的位置（`ppt/slides/_rels/`）出發，再把 `../` 字串直接 strip 掉：

```python
# 錯誤
rels_base = tmp_rels_path.parent.parent   # ppt/slides/
diag_path = (rels_base / target.replace("../", "")).resolve()
# 得到: ppt/slides/diagrams/data1.xml  ← 不存在，直接 continue
```

結果 `diag_path.exists()` 永遠是 `False`，複製邏輯完全跳過，兩張投影片依然共享同一份圖表檔。

**Bug B — 遺漏 Microsoft 私有 Namespace**

SmartArt 的 drawing canvas 使用的關係類型是：

```
http://schemas.microsoft.com/office/2007/relationships/diagramDrawing
```

而不是標準 OPC namespace：

```
http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramDrawing
```

`DIAGRAM_REL_TYPES` 集合只收錄了標準 namespace，導致 `drawing1.xml` 沒被複製，slide4 與 slide5 共享同一份 drawing file。

---

## 二、為什麼發生（設計層面）

| # | 問題 | 根本原因 |
|---|------|---------|
| 1 | `reorder_slides` 只改 XML 指針 | 誤以為修改 XML 邏輯結構等於修改實體檔案 |
| 2 | 深複製路徑計算錯誤 | 混淆了 `.rels` 檔案位置與其所屬 Part 位置（兩者差一層目錄）|
| 3 | 缺少 Microsoft namespace | OPC 標準與 Office 私有擴充並存，未全面列舉 |
| 4 | `stage2-build` 可重複執行問題 | `unpacked/` 被 in-place 修改後再次執行 reorder，對象已不是原始模板 |
| 5 | sldId 溢位 | 掃描所有 `id` 屬性包含 `sldMasterId=2147483648`，超過 int32 上限 |

---

## 三、修正摘要

| # | 問題 | 修法 |
|---|------|------|
| 1 | `reorder_slides` 只改 XML 指針 | 完整重寫 8 Phase，Phase 3 用 `shutil.copy2` 給每個位置建獨立實體副本 |
| 2 | SmartArt 圖表檔共享 | `_deep_copy_slide_resources()` 深複製所有 diagram 檔並更新 `.rels` target |
| 3 | 路徑從 `.rels` 位置解析（方向錯誤） | 改為 `(slide_dir / target).resolve()` 讓 Python Path 正確處理 `../` |
| 4 | 缺少 Microsoft 私有 namespace | 加入 `http://schemas.microsoft.com/office/2007/relationships/diagramDrawing` |
| 5 | `stage2-build` 重跑毀損 unpacked | `stage2-extract` 存 `.source_pptx`；`stage2-build` 每次先重新 unpack |
| 6 | sldId 溢位 | Phase 7 改為只掃 `<p:sldId>` 元素，排除 `sldMasterId` |

---

## 四、發生時怎麼做

### 4.1 快速診斷流程

```
PowerPoint 無法開啟 / 顯示修復對話框
        │
        ├─ 是否有 SLIDE_ORDER 重複 index？
        │     ├─ 是 → 檢查 diagrams/ 是否有多份副本
        │     │         ppt/diagrams/ 內 data2.xml、drawing2.xml 等應存在
        │     └─ 否 → 繼續
        │
        ├─ grep 是否有重複的 slide Target？
        │     grep "slides/slide" ppt/_rels/presentation.xml.rels | sort | uniq -d
        │
        ├─ 是否有 broken rels（target 指向不存在的檔案）？
        │     uv run python scripts/stage2/clean_orphans.py <unpacked_dir>
        │
        └─ sldId 是否超過 2147483647？
              grep "sldId id=" ppt/presentation.xml
```

### 4.2 手動修復步驟

1. 重跑 `stage2-extract`（確保 `.source_pptx` 存在且有效）
2. 重跑 `stage2-build`（會自動重置 `unpacked/` 到乾淨狀態）
3. 用 PowerShell COM 驗證結果：

```powershell
$ppt  = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open("C:\path\to\output_draft.pptx")
Write-Host "OK: $($pres.Slides.Count) slides"
$pres.Close(); $ppt.Quit()
```

### 4.3 緊急回退

若不確定根本原因，改用不含重複 index 的 `SLIDE_ORDER` 先產出可開啟的版本，再調查重複問題。

---

## 五、怎麼預防

1. **在正式輸出前一定用 Office COM 驗證**，不要只看 markitdown 或 python-pptx — 它們不驗證 Part 唯一性
2. **有重複 index 的 `SLIDE_ORDER` 要留意含 SmartArt 的投影片**，`pptx_utils.reorder_slides` 已自動處理，但若未來升級要確保深複製邏輯完整
3. **擴充 `DIAGRAM_REL_TYPES` 時** 同時查 OPC 標準與 Microsoft 私有 namespace，兩者都可能出現
4. **每次修改 inject_content.py 後**，透過 `stage2-build` 重跑而不是手動操作 `unpacked/`

---

## 六、Bug 處理機制與禁止事項

### 6.1 Agent 應遵守的 Bug 處理機制

執行 Stage 2 期間若遇到以下症狀，依下列順序處理：

| 症狀 | 優先動作 |
|------|---------|
| PowerPoint 無法開啟（`0x80070570`） | 重跑 `stage2-extract` + `stage2-build`，確認 diagrams/ 有對應的副本檔 |
| PowerPoint 開啟後顯示「修復」對話框 | 執行 `clean_orphans.py`，再檢查 `.rels` 內是否有 broken target |
| `[WARN] Slide index N out of range` | 表示 `stage2-build` 在已 reorder 的目錄上再次執行；重跑 `stage2-extract` 讓 `.source_pptx` 更新 |
| `[WARN] TextBox 'XXX' not found` | 從 `placeholders.txt` 重新確認名稱，**不可自行猜測或縮短** |
| markitdown 輸出顯示原始模板文字殘留 | 補充 CONTENT_MAP 缺失的 mapping，不要刪除投影片來掩蓋問題 |
| CONTENT_MAP 注入到錯誤的投影片 | 確認 key 用的是 **FINAL 順序**（`SLIDE_ORDER` 執行後的位置），而非原始模板編號 |

### 6.2 禁止事項（執行 Agent 必須遵守）

> 以下行為已被確認會導致 PPTX 毀損或結果不一致，**嚴格禁止**：

---

**❌ 禁止直接手動編輯 `unpacked/` 目錄**

`stage2-build` 每次執行都會從 `.source_pptx` 重新解壓縮，覆蓋 `unpacked/`。手動對 `unpacked/` 的任何修改都會在下次 build 時被清除。所有修改必須透過 `inject_content.py` 進行。

---

**❌ 禁止跳過 `stage2-extract` 直接執行 `stage2-build`**

若 `output_dir/.source_pptx` 不存在，`stage2-build` 無法重置 `unpacked/`，會在可能已被污染的狀態上執行，產生不可預測的結果。**務必先跑 `stage2-extract`。**

---

**❌ 禁止用 markitdown / python-pptx 替代 Office COM 作為正確性確認**

markitdown 和 python-pptx 只讀取文字內容，**不驗證 PPTX 的 Part 唯一性與 OPC 結構合規性**。即使它們顯示正常，PowerPoint 仍可能拒絕開啟。最終驗證必須用 Office COM：

```powershell
$pres = (New-Object -ComObject PowerPoint.Application).Presentations.Open($path)
Write-Host "Slides: $($pres.Slides.Count)"
```

---

**❌ 禁止在 CONTENT_MAP 使用原始模板的投影片編號**

`SLIDE_ORDER` 執行後，投影片已被重新編號。CONTENT_MAP 的 key 必須是 **SLIDE_ORDER 執行後的最終位置（1-based）**，而不是原始模板裡的投影片編號。

```python
# 錯誤：SLIDE_ORDER = [3, 7, 6, 6] 後，key 6 不存在最終輸出
CONTENT_MAP = { 6: { "Title 2": "..." } }   # ❌

# 正確：SLIDE_ORDER = [3, 7, 6, 6] → final positions 1,2,3,4
CONTENT_MAP = { 3: { "Title 2": "..." } }   # ✓ 第 3 個位置 = 原始 slide 6 第一份
```

---

**❌ 禁止用 lxml 寫 `.rels` 或 `[Content_Types].xml`**

lxml 在序列化這兩類檔案時會把命名空間前綴改寫為 `ns0:` / `ns1:`，PowerPoint 無法識別，直接拒絕開啟。這兩類檔案**只能用 `defusedxml.minidom` 搭配 `_serialize_xml_bytes()`** 處理。

---

**❌ 禁止假設 SmartArt/圖表檔案可以共享**

當同一張含 SmartArt 的投影片出現在 `SLIDE_ORDER` 兩次以上時，`pptx_utils.reorder_slides` 已自動深複製 diagram 檔案。若繞過此函式直接操作 XML，**必須手動複製所有 diagram 檔並更新 `.rels` target**，否則必定毀損。

需要深複製的關係類型：
```
…/diagramData
…/diagramLayout
…/diagramColors
…/diagramQuickStyle
…/diagramDrawing  （標準）
http://schemas.microsoft.com/office/2007/relationships/diagramDrawing  （MS 私有）
```

---

**❌ 禁止用任何現有 sldId 值（含 sldMasterId）作為新 sldId 的起始點**

`sldMasterId` 的 `id` 屬性值為 `2147483648`（超過 int32 上限）。若掃描所有 `id` 屬性並加 1，會產生 `2147483649` 溢位值，PowerPoint 拒絕載入。新 sldId 只能從 `<p:sldIdLst>` 內的 `<p:sldId>` 子元素取最大值。
