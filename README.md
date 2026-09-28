# CJGE Word Toolkit

简体中文 | [English](README-en.md) | [繁體中文](README-zhTW.md)

一套 **Word MCP 服务器 + AI Agent 技能（skill）**，让 Claude Code、Claude Desktop、Codex 等 AI Agent 读写 Word 文档，并把论文**除 MathType 公式以外的全部格式**排成 **《岩土工程学报》（Chinese Journal of Geotechnical Engineering，CJGE）** 的刊出版式：页面与分栏、题名作者单位、中英文摘要、标题层次、正文、图表题、三线表、参考文献、文末声明、页眉页码。

版本 **1.0.0**。由两部分封装而成：

- **MCP 服务器**：fork 自 [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)（MIT）。只改了格式配置（新文档、标题样式、新表格的默认格式），54 个工具的用法与原版相同。
- **Skill**：`docx-editor-cn`（原作者 Gostyan，MIT），所有格式规则改写为 CJGE。

MathType 公式的格式（字号、字体、编号、引用）请配合 [xyj0727/mathtype-office-toolkit](https://github.com/xyj0727/mathtype-office-toolkit) 使用，本工具不会改动 MathType 对象。

---

## 目录

- [适用场景](#适用场景)
- [环境要求](#环境要求)
- [核心能力](#核心能力)
- [默认格式（CJGE）](#默认格式cjge)
- [MCP 工具与 Skill](#mcp-工具与-skill)
- [安装](#安装)
- [验证安装](#验证安装)
- [与 MathType 工具配合](#与-mathtype-工具配合)
- [本 fork 的改动](#本-fork-的改动)
- [仓库结构](#仓库结构)
- [致谢与许可](#致谢与许可)

## 适用场景

- 把写好的论文（.docx）一次性排成《岩土工程学报》格式，再投稿或修改清稿。
- 用 Markdown 或脚本从零生成符合 CJGE 版式的 Word 稿件。
- 让 AI 在 Word 里新建文档、加标题、加段落、加表格时，直接得到 CJGE 格式（宋体/Times New Roman、五号、固定行距 15.6 磅、三线表……）。
- 检查一篇稿件哪些段落的字体、字号、行距、缩进、表格线、页边距不符合 CJGE。

**不适用于**：MathType 公式本身的格式（交给 MathType 工具）；其他期刊的版式（可复制 `cjge.json` 自行改数值）。

## 环境要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows、macOS、Linux 均可（MCP 与排版脚本只读写 .docx 文件，不需要安装 Word）；`convert_to_pdf` 需要 Word 或 LibreOffice |
| Python | 3.11 及以上；推荐用 [uv](https://docs.astral.sh/uv/) 安装 |
| Node.js | 18 及以上，仅在用 `new_doc.js` 新建文档时需要（`npm install docx temml fast-xml-parser`） |
| AI 客户端 | 支持 MCP 的 Agent：Claude Code、Claude Desktop、Codex 等 |

## 核心能力

1. **整篇排版**（`skills/docx-editor-cn/scripts/apply_format.py apply`）：自动识别题名、作者、单位、摘要、关键词、中图分类号、作者简介、英文题名/作者/单位/Abstract/Key words、首页脚注、各级标题、正文、图表题、分图题、表格、文末声明、参考文献，按 CJGE 设置字体、字号、行距、缩进、对齐；“摘  要：”“Abstract”等标签单独设字体。只改格式属性，不重建文档，域、书签、嵌入对象位置不变。
2. **页面与分栏**：A4、CJGE 页边距；在 Key words 后插入连续分节符，题名到 Key words 通栏，正文双栏。
3. **页眉页码**：奇偶页不同、首页不同；页码在页眉，无页脚。奇数页页眉默认取“第一作者，等. 题名”。
4. **三线表**：所有非公式表格改为三线表（顶线、底线 0.75 磅，栏目线 0.5 磅，无竖线，表内小五）。
5. **格式检查**（`apply_format.py audit`）：按 Word 的样式继承规则（直接格式 → 样式 → 文档默认）逐段检查，列出不符合 CJGE 的字体、字号、加粗、行距、缩进、对齐、表格线、页边距和分栏。
6. **新建稿件**（`new_doc.js`）：内置 CJGE 样式和前置部分、标题、图表题、三线表、声明、参考文献等辅助函数；也可直接把 Markdown 转成 CJGE 格式的 .docx。
7. **Word MCP 默认格式**：`create_document` 新建的文档自带 CJGE 页面和样式；`add_heading` 的一至三级标题、`add_table` 的表格也直接是 CJGE 格式。
8. **不碰 MathType**：含 MathType 对象的段落、公式表格、编号和引用一律跳过；含行内公式或图片的段落改用“最小值”行距，避免被固定行距截断。

## 默认格式（CJGE）

字号对照：二号 22 pt，四号 14 pt，小四 12 pt，五号 10.5 pt，小五 9 pt，六号 7.5 pt。中文用宋体/黑体/仿宋，**西文和数字一律 Times New Roman（TNR）**。

| 元素 | 字体 | 字号 | 段落 |
|---|---|---|---|
| 页面 | — | — | A4；左 1.65 cm、右 1.53 cm、上下 2.15 cm（版心 17.8 cm × 25.4 cm）；页眉距边界 1.5 cm |
| 分栏 | — | — | 题名至 Key words 通栏；正文双栏，栏宽 8.53 cm，栏距 0.74 cm |
| 中文题名 | 黑体 | 二号 | 居中 |
| 作者 | 仿宋，序号上标 | 五号 | 居中 |
| 单位 | 宋体 | 六号 | 居中 |
| 摘  要：/ 关键词： | 标签黑体，内容宋体 | 小五 | 左右各缩进 0.75 cm，两端对齐，行距固定值 14 磅 |
| 中图分类号、作者简介 | 标签黑体 | 小五 | 同上，行距固定值 15.6 磅 |
| 英文题名 | TNR 加粗 | 四号 | 居中 |
| 英文作者 / 单位 | TNR | 五号 / 六号 | 居中 |
| Abstract: / Key words: | 标签 TNR 加粗（冒号不加粗） | 小五 | 左右各缩进 0.75 cm，行距固定值 15.6 磅 |
| 首页脚注（基金项目、通信作者、收稿日期） | 中文标签黑体，英文标签加粗 | 六号 | 行距固定值 12 磅 |
| 正文 | 宋体 + TNR | 五号 | 两端对齐，首行缩进 2 字符，**行距固定值 15.6 磅**，段前段后 0 |
| 一级标题“0  引    言” | 黑体，不加粗 | 四号 | 顶格，段前约一行 |
| 二级标题“1.1  试验装置” | 黑体，不加粗 | 五号 | 顶格 |
| 三级标题“（1）级配影响” | 宋体 | 五号 | 首行缩进 2 字符，单独成行 |
| 四级标题“a）相似性。” | 宋体 + TNR | 五号 | 与正文接排 |
| 图题“图1 题名”“Fig. 1 Title” | 黑体 / TNR，不加粗 | 小五 | 居中，在图下方，行距固定值 15.6 磅 |
| 表题“表1 题名”“Table 1 Title” | 黑体 / TNR，不加粗 | 小五 | 居中，在表上方，行距固定值 14 磅 |
| 表格 | 宋体 + TNR | 小五 | 三线表：顶线、底线 0.75 磅，栏目线 0.5 磅，无竖线 |
| 利益冲突声明、作者贡献 | 标签加粗（中文黑体） | 小五 | 顶格，行距固定值 15.6 磅 |
| 参考文献： | 黑体 | 五号 | 顶格 |
| 参考文献条目 | 宋体 + TNR | 小五 | 两端对齐，行距固定值 16 磅，悬挂缩进 2 字符 |
| 页眉 | 宋体 + TNR | 小五 | 奇数页“第N期 / 第一作者，等. 题名 / 页码”；偶数页“页码 / 岩 土 工 程 学 报 / 年份”；0.5 磅页眉线；无页脚 |

全部数值见 [`skills/docx-editor-cn/scripts/format_profiles/cjge.json`](skills/docx-editor-cn/scripts/format_profiles/cjge.json)（整篇排版）和 [`word_document_server/core/format_config.py`](word_document_server/core/format_config.py)（MCP 默认格式）。依据：《岩土工程学报》2026 年第 9 期刊出论文的 PDF 实测值与官网《征稿简则》。

## MCP 工具与 Skill

**MCP 服务器名**：`word-document-server`，命令 `word_mcp_server`。工具共 54 个，名称和参数与原版一致，常用的有：

| 类别 | 工具 |
|---|---|
| 文档 | `create_document`、`copy_document`、`get_document_info`、`get_document_text`、`get_document_outline`、`get_document_xml`、`list_available_documents`、`convert_to_pdf` |
| 内容 | `add_heading`、`add_paragraph`、`add_table`、`add_picture`、`add_page_break`、`delete_paragraph`、`search_and_replace`、`insert_header_near_text`、`insert_line_or_paragraph_near_text`、`insert_numbered_list_near_text` |
| 格式 | `format_text`、`create_custom_style`、`format_table`、`format_table_cell_text`、`set_table_*`、`merge_table_cells*`、`highlight_table_header` |
| 脚注 / 批注 / 保护 | `add_footnote_*`、`add_endnote_to_document`、`get_all_comments`、`protect_document`、`unprotect_document` |

完整工具说明见 [docs/upstream-README.md](docs/upstream-README.md)。

**Skill 名**：`docx-editor-cn`（目录 `skills/docx-editor-cn/`）。Agent 读取 `SKILL.md`，调用其中的脚本：

| 脚本 | 作用 |
|---|---|
| `scripts/apply_format.py apply IN.docx OUT.docx` | 整篇排成 CJGE 格式。可选 `--header` 奇数页页眉、`--year`、`--volume --issue --month`（首页刊头）、`--no-journal-header`、`--no-columns` |
| `scripts/apply_format.py audit DOC.docx` | 检查格式，输出 JSON 问题清单 |
| `node scripts/new_doc.js [paper.md out.docx]` | 新建 CJGE 稿件，或把 Markdown 转成 .docx |
| `scripts/table.py` | 在解包的文档里插入三线表（含中英文表题） |
| `scripts/office/unpack.py`、`pack.py`、`validate.py` | 解包、打包、校验 .docx |

## 安装

### Claude Code

以插件方式安装（MCP 服务器和 Skill 一起装）：

```bash
claude plugin marketplace add xyj0727/cjge-word-toolkit
```

```bash
claude plugin install cjge-word-toolkit@cjge-word-toolkit
```

或者分别安装。MCP 服务器：

```bash
claude mcp add word-document-server -s user -- uvx --from git+https://github.com/xyj0727/cjge-word-toolkit word_mcp_server
```

Skill：把 `skills/docx-editor-cn` 文件夹复制到 `~/.claude/skills/`。

### Claude Desktop

在 `claude_desktop_config.json` 中加入（示例见 [mcp-config.json](mcp-config.json)）：

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

Skill：把 `skills/docx-editor-cn` 压缩成 zip，在“设置 → 功能 → Skills”上传。

### Codex

在 `~/.codex/config.toml` 中加入：

```toml
[mcp_servers.word-document-server]
command = "uvx"
args = ["--from", "git+https://github.com/xyj0727/cjge-word-toolkit", "word_mcp_server"]
```

Skill：把 `skills/docx-editor-cn` 复制到 `~/.codex/skills/`。

### 本地安装（开发用）

```bash
git clone https://github.com/xyj0727/cjge-word-toolkit.git
```

```bash
uv tool install --force --editable ./cjge-word-toolkit
```

安装后命令 `word_mcp_server` 即指向本地源码。新建文档功能还需在 Skill 目录执行 `npm install`。

## 验证安装

1. 在 Agent 里让它调用 `create_document` 新建一个文档，再用 `get_document_xml` 查看：页边距应为左 935、右 867、上下 1219（单位 twip），`Normal` 样式为 Times New Roman / 宋体、五号（`w:sz="21"`）、行距 `exact 312`。
2. 对一篇论文运行：

```bash
python skills/docx-editor-cn/scripts/apply_format.py apply 论文.docx 论文-CJGE.docx
```

```bash
python skills/docx-editor-cn/scripts/apply_format.py audit 论文-CJGE.docx
```

`audit` 输出 `"ok": true`、`"issue_count": 0` 即格式全部符合。

## 与 MathType 工具配合

含 MathType 公式的论文，推荐流程：

1. `apply_format.py apply` 排好正文等格式；
2. 用 [mathtype-office-toolkit](https://github.com/xyj0727/mathtype-office-toolkit) 的 `render_mathtype_word_document` 生成或统一 CJGE 公式格式；
3. 再运行一次 `apply_format.py apply`；
4. 用 `validate_mathtype_word_document` 检查公式，用本仓库 MCP 的 `get_document_xml` 检查其余格式。

## 本 fork 的改动

**MCP 服务器（只改格式配置）**

- 新增 [`word_document_server/core/format_config.py`](word_document_server/core/format_config.py)，集中存放 CJGE 默认格式。
- `create_document`：新文档采用 A4 + CJGE 页边距，文档默认字体和 Normal、Title、Heading 1–3、Caption 样式按 CJGE 设置。
- `add_heading`：一至三级标题样式为 CJGE（黑体四号 / 黑体五号 / 宋体五号，不加粗，黑色）；无样式时的备用字号也按 CJGE。
- `add_table`：由 `Table Grid` 网格表改为三线表，表内小五。
- 显式传入的参数（`font_name`、`font_size`、`bold` 等）仍然优先。

**Skill（docx-editor-cn）**

- 格式规则全部改为 CJGE：`SKILL.md`、`format_profiles/cjge.json`（替换原 `docx_editor_cn.json`）、`new_doc.js`、`table.py`。
- `apply_format.py` 重写：识别 CJGE 各部分、标签单独设字体、通栏/双栏分节、奇偶页页眉、首页脚注、按样式继承检查格式。
- 删除与本 Skill 无关的示例脚本 `convert_paper.js`。

**封装**

- 包名改为 `cjge-word-toolkit` 1.0.0，命令名 `word_mcp_server` 与原版一致，可直接替换。
- 新增 Claude Code 插件清单（`.claude-plugin/`），原版 README 移至 `docs/upstream-README.md`。

## 仓库结构

```
cjge-word-toolkit/
├── word_document_server/          ← MCP 服务器（原版代码）
│   └── core/format_config.py      ← CJGE 默认格式（本 fork 新增）
├── skills/docx-editor-cn/         ← Skill（CJGE 版）
│   ├── SKILL.md
│   └── scripts/
│       ├── apply_format.py        ← 整篇排版 / 格式检查
│       ├── format_profiles/cjge.json
│       ├── new_doc.js             ← 新建稿件 / Markdown 转换
│       └── table.py …
├── .claude-plugin/                ← Claude Code 插件清单
├── docs/upstream-README.md        ← 原版 README（全部 MCP 工具说明）
├── pyproject.toml                 ← 包名 cjge-word-toolkit
└── LICENSE
```

## 致谢与许可

- MCP 服务器：[GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)，MIT。
- Skill：docx-editor-cn，作者 Gostyan，MIT（见 `skills/docx-editor-cn/LICENSE.txt`）。
- 格式依据：《岩土工程学报》刊出论文与官网《征稿简则》。本项目与该刊编辑部无关，投稿以编辑部最新要求为准。

本仓库以 MIT 许可发布，见 [LICENSE](LICENSE)。
