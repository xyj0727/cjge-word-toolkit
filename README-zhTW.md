# CJGE Word Toolkit

[简体中文](README.md) | [English](README-en.md) | 繁體中文

一套 **Word MCP 伺服器 + AI Agent 技能（skill）**，讓 Claude Code、Claude Desktop、Codex 等 AI Agent 讀寫 Word 文件，並把論文**除 MathType 公式以外的全部格式**排成 **《岩土工程学报》（Chinese Journal of Geotechnical Engineering，CJGE）** 的刊出版式：頁面與分欄、題名作者單位、中英文摘要、標題層次、正文、圖表題、三線表、參考文獻、文末聲明、頁首頁碼。

版本 **1.0.0**。由兩部分封裝而成：

- **MCP 伺服器**：fork 自 [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)（MIT）。只改了格式配置（新文件、標題樣式、新表格的預設格式），54 個工具的用法與原版相同。
- **Skill**：`docx-editor-cn`（原作者 Gostyan，MIT），所有格式規則改寫為 CJGE。

MathType 公式的格式（字號、字型、編號、引用）請配合 [xyj0727/mathtype-office-toolkit](https://github.com/xyj0727/mathtype-office-toolkit) 使用，本工具不會改動 MathType 物件。

---

## 目錄

- [適用場景](#適用場景)
- [環境要求](#環境要求)
- [核心能力](#核心能力)
- [預設格式（CJGE）](#預設格式cjge)
- [MCP 工具與 Skill](#mcp-工具與-skill)
- [安裝](#安裝)
- [驗證安裝](#驗證安裝)
- [與 MathType 工具配合](#與-mathtype-工具配合)
- [本 fork 的改動](#本-fork-的改動)
- [倉庫結構](#倉庫結構)
- [致謝與許可](#致謝與許可)

## 適用場景

- 把寫好的論文（.docx）一次性排成《岩土工程学报》格式，再投稿或修改清稿。
- 用 Markdown 或指令碼從零生成符合 CJGE 版式的 Word 稿件。
- 讓 AI 在 Word 裡新建文件、加標題、加段落、加表格時，直接得到 CJGE 格式（宋体/Times New Roman、五號、固定行距 15.6 磅、三線表……）。
- 檢查一篇稿件哪些段落的字型、字號、行距、縮排、表格線、頁邊距不符合 CJGE。

**不適用於**：MathType 公式本身的格式（交給 MathType 工具）；其他期刊的版式（可複製 `cjge.json` 自行改數值）。

## 環境要求

| 專案 | 要求 |
|---|---|
| 作業系統 | Windows、macOS、Linux 均可（MCP 與排版指令碼只讀寫 .docx 檔案，不需要安裝 Word）；`convert_to_pdf` 需要 Word 或 LibreOffice |
| Python | 3.11 及以上；推薦用 [uv](https://docs.astral.sh/uv/) 安裝 |
| Node.js | 18 及以上，僅在用 `new_doc.js` 新建文件時需要（`npm install docx temml fast-xml-parser`） |
| AI 客戶端 | 支援 MCP 的 Agent：Claude Code、Claude Desktop、Codex 等 |

## 核心能力

1. **整篇排版**（`skills/docx-editor-cn/scripts/apply_format.py apply`）：自動識別題名、作者、單位、摘要、關鍵詞、中圖分類號、作者簡介、英文題名/作者/單位/Abstract/Key words、首頁尾註、各級標題、正文、圖表題、分圖題、表格、文末聲明、參考文獻，按 CJGE 設定字型、字號、行距、縮排、對齊；“摘  要：”“Abstract”等標籤單獨設字型。只改格式屬性，不重建文件，域、書籤、嵌入物件位置不變。
2. **頁面與分欄**：A4、CJGE 頁邊距；在 Key words 後插入連續分節符，題名到 Key words 通欄，正文雙欄。
3. **頁首頁碼**：奇偶頁不同、首頁不同；頁碼在頁首，無頁尾。奇數頁頁首預設取“第一作者，等. 题名”。
4. **三線表**：所有非公式表格改為三線表（頂線、底線 0.75 磅，欄目線 0.5 磅，無豎線，表內小五）。
5. **格式檢查**（`apply_format.py audit`）：按 Word 的樣式繼承規則（直接格式 → 樣式 → 文件預設）逐段檢查，列出不符合 CJGE 的字型、字號、加粗、行距、縮排、對齊、表格線、頁邊距和分欄。
6. **新建稿件**（`new_doc.js`）：內建 CJGE 樣式和前置部分、標題、圖表題、三線表、聲明、參考文獻等輔助函式；也可直接把 Markdown 轉成 CJGE 格式的 .docx。
7. **Word MCP 預設格式**：`create_document` 新建的文件自帶 CJGE 頁面和樣式；`add_heading` 的一至三級標題、`add_table` 的表格也直接是 CJGE 格式。
8. **不碰 MathType**：含 MathType 物件的段落、公式表格、編號和引用一律跳過；含行內公式或圖片的段落改用“最小值”行距，避免被固定行距截斷。

## 預設格式（CJGE）

字號對照：二號 22 pt，四號 14 pt，小四 12 pt，五號 10.5 pt，小五 9 pt，六號 7.5 pt。中文用宋体/黑体/仿宋，**西文和數字一律 Times New Roman（TNR）**。

| 元素 | 字型 | 字號 | 段落 |
|---|---|---|---|
| 頁面 | — | — | A4；左 1.65 cm、右 1.53 cm、上下 2.15 cm（版心 17.8 cm × 25.4 cm）；頁首距邊界 1.5 cm |
| 分欄 | — | — | 題名至 Key words 通欄；正文雙欄，欄寬 8.53 cm，欄距 0.74 cm |
| 中文題名 | 黑体 | 二號 | 居中 |
| 作者 | 仿宋，序號上標 | 五號 | 居中 |
| 單位 | 宋体 | 六號 | 居中 |
| 摘  要：/ 關鍵詞： | 標籤黑体，內容宋体 | 小五 | 左右各縮排 0.75 cm，兩端對齊，行距固定值 14 磅 |
| 中圖分類號、作者簡介 | 標籤黑体 | 小五 | 同上，行距固定值 15.6 磅 |
| 英文題名 | TNR 加粗 | 四號 | 居中 |
| 英文作者 / 單位 | TNR | 五號 / 六號 | 居中 |
| Abstract: / Key words: | 標籤 TNR 加粗（冒號不加粗） | 小五 | 左右各縮排 0.75 cm，行距固定值 15.6 磅 |
| 首頁尾註（基金專案、通訊作者、收稿日期） | 中文標籤黑体，英文標籤加粗 | 六號 | 行距固定值 12 磅 |
| 正文 | 宋体 + TNR | 五號 | 兩端對齊，首行縮排 2 字元，**行距固定值 15.6 磅**，段前段後 0 |
| 一級標題“0  引    言” | 黑体，不加粗 | 四號 | 頂格，段前約一行 |
| 二級標題“1.1  试验装置” | 黑体，不加粗 | 五號 | 頂格 |
| 三級標題“（1）级配影响” | 宋体 | 五號 | 首行縮排 2 字元，單獨成行 |
| 四級標題“a）相似性。” | 宋体 + TNR | 五號 | 與正文接排 |
| 圖題“图1 题名”“Fig. 1 Title” | 黑体 / TNR，不加粗 | 小五 | 居中，在圖下方，行距固定值 15.6 磅 |
| 表題“表1 题名”“Table 1 Title” | 黑体 / TNR，不加粗 | 小五 | 居中，在表上方，行距固定值 14 磅 |
| 表格 | 宋体 + TNR | 小五 | 三線表：頂線、底線 0.75 磅，欄目線 0.5 磅，無豎線 |
| 利益衝突聲明、作者貢獻 | 標籤加粗（中文黑体） | 小五 | 頂格，行距固定值 15.6 磅 |
| 參考文獻： | 黑体 | 五號 | 頂格 |
| 參考文獻條目 | 宋体 + TNR | 小五 | 兩端對齊，行距固定值 16 磅，懸掛縮排 2 字元 |
| 頁首 | 宋体 + TNR | 小五 | 奇數頁“第N期 / 第一作者，等. 题名 / 页码”；偶數頁“页码 / 岩 土 工 程 学 报 / 年份”；0.5 磅頁首線；無頁尾 |

全部數值見 [`skills/docx-editor-cn/scripts/format_profiles/cjge.json`](skills/docx-editor-cn/scripts/format_profiles/cjge.json)（整篇排版）和 [`word_document_server/core/format_config.py`](word_document_server/core/format_config.py)（MCP 預設格式）。依據：《岩土工程学报》2026 年第 9 期刊出論文的 PDF 實測值與官網《征稿简则》。

## MCP 工具與 Skill

**MCP 伺服器名**：`word-document-server`，命令 `word_mcp_server`。工具共 54 個，名稱和引數與原版一致，常用的有：

| 類別 | 工具 |
|---|---|
| 文件 | `create_document`、`copy_document`、`get_document_info`、`get_document_text`、`get_document_outline`、`get_document_xml`、`list_available_documents`、`convert_to_pdf` |
| 內容 | `add_heading`、`add_paragraph`、`add_table`、`add_picture`、`add_page_break`、`delete_paragraph`、`search_and_replace`、`insert_header_near_text`、`insert_line_or_paragraph_near_text`、`insert_numbered_list_near_text` |
| 格式 | `format_text`、`create_custom_style`、`format_table`、`format_table_cell_text`、`set_table_*`、`merge_table_cells*`、`highlight_table_header` |
| 腳註 / 批註 / 保護 | `add_footnote_*`、`add_endnote_to_document`、`get_all_comments`、`protect_document`、`unprotect_document` |

完整工具說明見 [docs/upstream-README.md](docs/upstream-README.md)。

**Skill 名**：`docx-editor-cn`（目錄 `skills/docx-editor-cn/`）。Agent 讀取 `SKILL.md`，呼叫其中的指令碼：

| 指令碼 | 作用 |
|---|---|
| `scripts/apply_format.py apply IN.docx OUT.docx` | 整篇排成 CJGE 格式。可選 `--header` 奇數頁頁首、`--year`、`--volume --issue --month`（首頁刊頭）、`--no-journal-header`、`--no-columns` |
| `scripts/apply_format.py audit DOC.docx` | 檢查格式，輸出 JSON 問題清單 |
| `node scripts/new_doc.js [paper.md out.docx]` | 新建 CJGE 稿件，或把 Markdown 轉成 .docx |
| `scripts/table.py` | 在解包的文件裡插入三線表（含中英文表題） |
| `scripts/office/unpack.py`、`pack.py`、`validate.py` | 解包、打包、校驗 .docx |

## 安裝

### Claude Code

以外掛方式安裝（MCP 伺服器和 Skill 一起裝）：

```bash
claude plugin marketplace add xyj0727/cjge-word-toolkit
```

```bash
claude plugin install cjge-word-toolkit@cjge-word-toolkit
```

或者分別安裝。MCP 伺服器：

```bash
claude mcp add word-document-server -s user -- uvx --from git+https://github.com/xyj0727/cjge-word-toolkit word_mcp_server
```

Skill：把 `skills/docx-editor-cn` 資料夾複製到 `~/.claude/skills/`。

### Claude Desktop

在 `claude_desktop_config.json` 中加入（示例見 [mcp-config.json](mcp-config.json)）：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/xyj0727/cjge-word-toolkit", "word_mcp_server"]
    }
  }
}
```

Skill：把 `skills/docx-editor-cn` 壓縮成 zip，在“设置 → 功能 → Skills”上傳。

### Codex

在 `~/.codex/config.toml` 中加入：

```toml
[mcp_servers.word-document-server]
command = "uvx"
args = ["--from", "git+https://github.com/xyj0727/cjge-word-toolkit", "word_mcp_server"]
```

Skill：把 `skills/docx-editor-cn` 複製到 `~/.codex/skills/`。

### 本地安裝（開發用）

```bash
git clone https://github.com/xyj0727/cjge-word-toolkit.git
```

```bash
uv tool install --force --editable ./cjge-word-toolkit
```

安裝後命令 `word_mcp_server` 即指向本地原始碼。新建文件功能還需在 Skill 目錄執行 `npm install`。

## 驗證安裝

1. 在 Agent 裡讓它呼叫 `create_document` 新建一個文件，再用 `get_document_xml` 檢視：頁邊距應為左 935、右 867、上下 1219（單位 twip），`Normal` 樣式為 Times New Roman / 宋体、五號（`w:sz="21"`）、行距 `exact 312`。
2. 對一篇論文執行：

```bash
python skills/docx-editor-cn/scripts/apply_format.py apply 论文.docx 论文-CJGE.docx
```

```bash
python skills/docx-editor-cn/scripts/apply_format.py audit 论文-CJGE.docx
```

`audit` 輸出 `"ok": true`、`"issue_count": 0` 即格式全部符合。

## 與 MathType 工具配合

含 MathType 公式的論文，推薦流程：

1. `apply_format.py apply` 排好正文等格式；
2. 用 [mathtype-office-toolkit](https://github.com/xyj0727/mathtype-office-toolkit) 的 `render_mathtype_word_document` 生成或統一 CJGE 公式格式；
3. 再執行一次 `apply_format.py apply`；
4. 用 `validate_mathtype_word_document` 檢查公式，用本倉庫 MCP 的 `get_document_xml` 檢查其餘格式。

## 本 fork 的改動

**MCP 伺服器（只改格式配置）**

- 新增 [`word_document_server/core/format_config.py`](word_document_server/core/format_config.py)，集中存放 CJGE 預設格式。
- `create_document`：新文件採用 A4 + CJGE 頁邊距，文件預設字型和 Normal、Title、Heading 1–3、Caption 樣式按 CJGE 設定。
- `add_heading`：一至三級標題樣式為 CJGE（黑体四號 / 黑体五號 / 宋体五號，不加粗，黑色）；無樣式時的備用字號也按 CJGE。
- `add_table`：由 `Table Grid` 網格表改為三線表，表內小五。
- 顯式傳入的引數（`font_name`、`font_size`、`bold` 等）仍然優先。

**Skill（docx-editor-cn）**

- 格式規則全部改為 CJGE：`SKILL.md`、`format_profiles/cjge.json`（替換原 `docx_editor_cn.json`）、`new_doc.js`、`table.py`。
- `apply_format.py` 重寫：識別 CJGE 各部分、標籤單獨設字型、通欄/雙欄分節、奇偶頁頁首、首頁尾註、按樣式繼承檢查格式。
- 刪除與本 Skill 無關的示例指令碼 `convert_paper.js`。

**封裝**

- 包名改為 `cjge-word-toolkit` 1.0.0，命令名 `word_mcp_server` 與原版一致，可直接替換。
- 新增 Claude Code 外掛清單（`.claude-plugin/`），原版 README 移至 `docs/upstream-README.md`。

## 倉庫結構

```
cjge-word-toolkit/
├── word_document_server/          ← MCP 伺服器（原版程式碼）
│   └── core/format_config.py      ← CJGE 預設格式（本 fork 新增）
├── skills/docx-editor-cn/         ← Skill（CJGE 版）
│   ├── SKILL.md
│   └── scripts/
│       ├── apply_format.py        ← 整篇排版 / 格式檢查
│       ├── format_profiles/cjge.json
│       ├── new_doc.js             ← 新建稿件 / Markdown 轉換
│       └── table.py …
├── .claude-plugin/                ← Claude Code 外掛清單
├── docs/upstream-README.md        ← 原版 README（全部 MCP 工具說明）
├── pyproject.toml                 ← 包名 cjge-word-toolkit
└── LICENSE
```

## 致謝與許可

- MCP 伺服器：[GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)，MIT。
- Skill：docx-editor-cn，作者 Gostyan，MIT（見 `skills/docx-editor-cn/LICENSE.txt`）。
- 格式依據：《岩土工程学报》刊出論文與官網《征稿简则》。本專案與該刊編輯部無關，投稿以編輯部最新要求為準。

本倉庫以 MIT 許可釋出，見 [LICENSE](LICENSE)。
