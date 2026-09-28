# CJGE Word Toolkit

[简体中文](README.md) | English | [繁體中文](README-zhTW.md)

A **Word MCP server + AI-agent skill** that lets Claude Code, Claude Desktop, Codex and other MCP agents read and write Word documents, and lays out **everything except MathType equations** in the published format of the **Chinese Journal of Geotechnical Engineering (《岩土工程学报》, CJGE)**: page and columns, title/authors/affiliations, Chinese and English abstracts, heading levels, body text, figure and table captions, three-line tables, references, end statements and running headers.

Version **1.0.0**. Two parts, packaged together:

- **MCP server** — a fork of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) (MIT). Only the format configuration changed (defaults for new documents, heading styles and new tables); all 54 tools work as in the original.
- **Skill** — `docx-editor-cn` (by Gostyan, MIT) with every format rule rewritten for CJGE.

For the MathType equation format (size, fonts, numbering, references) use [xyj0727/mathtype-office-toolkit](https://github.com/xyj0727/mathtype-office-toolkit); this toolkit never modifies MathType objects.

---

## Contents

- [When to use](#when-to-use)
- [Requirements](#requirements)
- [Core capabilities](#core-capabilities)
- [Default format (CJGE)](#default-format-cjge)
- [MCP tools and skill](#mcp-tools-and-skill)
- [Installation](#installation)
- [Verify the installation](#verify-the-installation)
- [Working with the MathType toolkit](#working-with-the-mathtype-toolkit)
- [Changes in this fork](#changes-in-this-fork)
- [Repository layout](#repository-layout)
- [Credits and license](#credits-and-license)

## When to use

- Reformat a finished paper (.docx) to the CJGE layout before submission or for the revised manuscript.
- Generate a CJGE-formatted Word manuscript from Markdown or a script.
- Have an AI agent create documents, headings, paragraphs and tables in Word that come out in CJGE format directly (SimSun/Times New Roman, 10.5 pt, fixed 15.6 pt line spacing, three-line tables …).
- Find every paragraph whose font, size, line spacing, indent, table rules, margins or columns do not match CJGE.

**Not for**: the format of MathType equations themselves (use the MathType toolkit); other journals' layouts (copy `cjge.json` and change the values).

## Requirements

| Item | Requirement |
|---|---|
| OS | Windows, macOS or Linux (the server and scripts only read and write .docx files; Word is not needed). `convert_to_pdf` needs Word or LibreOffice |
| Python | 3.11+; installing with [uv](https://docs.astral.sh/uv/) is recommended |
| Node.js | 18+, only for creating documents with `new_doc.js` (`npm install docx temml fast-xml-parser`) |
| AI client | Any MCP agent: Claude Code, Claude Desktop, Codex, … |

## Core capabilities

1. **Whole-paper formatting** (`skills/docx-editor-cn/scripts/apply_format.py apply`): recognises the title, authors, affiliations, abstract, key words, CLC number line, author biography, the English title/authors/affiliations/Abstract/Key words, first-page notes, heading levels, body text, captions, sub-figure labels, tables, end statements and references, and sets fonts, sizes, line spacing, indents and alignment per CJGE. Labels such as “摘  要：” and “Abstract” get their own fonts. Only formatting properties change; the document is never rebuilt, so fields, bookmarks and embedded objects stay in place.
2. **Page and columns**: A4 with CJGE margins; a continuous section break after the Key words puts the front matter in one column and the body in two.
3. **Running headers**: different odd/even and first pages; the page number is in the header, there is no footer. The odd-page header defaults to “第一作者，等. 题名” (first author et al. Title).
4. **Three-line tables**: every non-equation table becomes a three-line table (top and bottom rules 0.75 pt, header rule 0.5 pt, no vertical rules, 9 pt text).
5. **Format audit** (`apply_format.py audit`): checks every paragraph with Word's inheritance rules (direct formatting → styles → document defaults) and lists fonts, sizes, bold, line spacing, indents, alignment, table rules, margins and columns that differ from CJGE.
6. **New manuscripts** (`new_doc.js`): CJGE styles plus helpers for front matter, headings, captions, three-line tables, statements and references; also converts Markdown straight to a CJGE .docx.
7. **Word MCP defaults**: documents created with `create_document` carry the CJGE page setup and styles; level 1–3 headings from `add_heading` and tables from `add_table` are CJGE-formatted as well.
8. **MathType is left alone**: paragraphs holding MathType objects, equation tables, numbers and references are skipped; paragraphs with inline equations or pictures get “at least” line spacing so fixed spacing never clips them.

## Default format (CJGE)

Chinese font sizes: 二号 22 pt, 四号 14 pt, 小四 12 pt, 五号 10.5 pt, 小五 9 pt, 六号 7.5 pt. Chinese text uses SimSun (宋体), SimHei (黑体) or FangSong (仿宋); **all Latin letters and digits use Times New Roman (TNR)**.

| Element | Font | Size | Paragraph |
|---|---|---|---|
| Page | — | — | A4; margins left 1.65 cm, right 1.53 cm, top/bottom 2.15 cm (type area 17.8 × 25.4 cm); header 1.5 cm from the edge |
| Columns | — | — | Title … Key words in one column; body in two columns of 8.53 cm with a 0.74 cm gap |
| Chinese title | SimHei | 22 pt | centred |
| Authors | FangSong, superscript affiliation numbers | 10.5 pt | centred |
| Affiliations | SimSun | 7.5 pt | centred |
| 摘  要： / 关键词： | label SimHei, text SimSun | 9 pt | 0.75 cm left and right indent, justified, fixed 14 pt |
| CLC line, author biography | label SimHei | 9 pt | same indent, fixed 15.6 pt |
| English title | TNR bold | 14 pt | centred |
| English authors / affiliations | TNR | 10.5 / 7.5 pt | centred |
| Abstract: / Key words: | label TNR bold (colon not bold) | 9 pt | 0.75 cm indent, fixed 15.6 pt |
| First-page notes (fund, corresponding author, dates) | Chinese labels SimHei, English labels bold | 7.5 pt | fixed 12 pt |
| Body | SimSun + TNR | 10.5 pt | justified, 2-character first-line indent, **fixed 15.6 pt**, no space before/after |
| Level 1 “0  引    言” | SimHei, not bold | 14 pt | flush left, about one line before |
| Level 2 “1.1  试验装置” | SimHei, not bold | 10.5 pt | flush left |
| Level 3 “（1）级配影响” | SimSun | 10.5 pt | 2-character indent, own line |
| Level 4 “a）相似性。” | SimSun + TNR | 10.5 pt | runs in with the text |
| Figure caption “图1 …” + “Fig. 1 …” | SimHei / TNR, not bold | 9 pt | centred below the figure, fixed 15.6 pt |
| Table caption “表1 …” + “Table 1 …” | SimHei / TNR, not bold | 9 pt | centred above the table, fixed 14 pt |
| Tables | SimSun + TNR | 9 pt | three-line: top/bottom 0.75 pt, header rule 0.5 pt, no vertical rules |
| Conflict of interests, author contributions | label bold (Chinese part SimHei) | 9 pt | flush left, fixed 15.6 pt |
| 参考文献： | SimHei | 10.5 pt | flush left |
| Reference entries | SimSun + TNR | 9 pt | justified, fixed 16 pt, 2-character hanging indent |
| Header | SimSun + TNR | 9 pt | odd: issue / first author et al. title / page; even: page / 岩 土 工 程 学 报 / year; 0.5 pt rule; no footer |

All values: [`skills/docx-editor-cn/scripts/format_profiles/cjge.json`](skills/docx-editor-cn/scripts/format_profiles/cjge.json) (whole-paper formatting) and [`word_document_server/core/format_config.py`](word_document_server/core/format_config.py) (MCP defaults). Source: measurements of a paper published in CJGE 2026 issue 9 and the journal's author guidelines.

## MCP tools and skill

**MCP server name**: `word-document-server`, command `word_mcp_server`. 54 tools with the same names and parameters as the original, for example:

| Group | Tools |
|---|---|
| Documents | `create_document`, `copy_document`, `get_document_info`, `get_document_text`, `get_document_outline`, `get_document_xml`, `list_available_documents`, `convert_to_pdf` |
| Content | `add_heading`, `add_paragraph`, `add_table`, `add_picture`, `add_page_break`, `delete_paragraph`, `search_and_replace`, `insert_header_near_text`, `insert_line_or_paragraph_near_text`, `insert_numbered_list_near_text` |
| Formatting | `format_text`, `create_custom_style`, `format_table`, `format_table_cell_text`, `set_table_*`, `merge_table_cells*`, `highlight_table_header` |
| Footnotes / comments / protection | `add_footnote_*`, `add_endnote_to_document`, `get_all_comments`, `protect_document`, `unprotect_document` |

Full tool reference: [docs/upstream-README.md](docs/upstream-README.md).

**Skill name**: `docx-editor-cn` (folder `skills/docx-editor-cn/`). The agent reads `SKILL.md` and runs its scripts:

| Script | Purpose |
|---|---|
| `scripts/apply_format.py apply IN.docx OUT.docx` | Format a whole paper. Options: `--header` (odd-page header), `--year`, `--volume --issue --month` (first-page masthead), `--no-journal-header`, `--no-columns` |
| `scripts/apply_format.py audit DOC.docx` | Check the format; prints a JSON list of issues |
| `node scripts/new_doc.js [paper.md out.docx]` | Create a CJGE manuscript or convert Markdown |
| `scripts/table.py` | Insert a three-line table (with Chinese and English captions) into an unpacked document |
| `scripts/office/unpack.py`, `pack.py`, `validate.py` | Unpack, pack and validate .docx files |

## Installation

### Claude Code

As a plugin (server and skill together):

```bash
claude plugin marketplace add xyj0727/cjge-word-toolkit
```

```bash
claude plugin install cjge-word-toolkit@cjge-word-toolkit
```

Or separately. MCP server:

```bash
claude mcp add word-document-server -s user -- uvx --from git+https://github.com/xyj0727/cjge-word-toolkit word_mcp_server
```

Skill: copy `skills/docx-editor-cn` to `~/.claude/skills/`.

### Claude Desktop

Add to `claude_desktop_config.json` (see [mcp-config.json](mcp-config.json)):

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

Skill: zip `skills/docx-editor-cn` and upload it under Settings → Capabilities → Skills.

### Codex

Add to `~/.codex/config.toml`:

```toml
[mcp_servers.word-document-server]
command = "uvx"
args = ["--from", "git+https://github.com/xyj0727/cjge-word-toolkit", "word_mcp_server"]
```

Skill: copy `skills/docx-editor-cn` to `~/.codex/skills/`.

### Local install (development)

```bash
git clone https://github.com/xyj0727/cjge-word-toolkit.git
```

```bash
uv tool install --force --editable ./cjge-word-toolkit
```

The `word_mcp_server` command then runs the local source. Run `npm install` in the skill folder to create documents with `new_doc.js`.

## Verify the installation

1. Ask the agent to call `create_document`, then `get_document_xml`: margins should be left 935, right 867, top/bottom 1219 (twips), and the `Normal` style Times New Roman / SimSun, 10.5 pt (`w:sz="21"`), line spacing `exact 312`.
2. Run on a paper:

```bash
python skills/docx-editor-cn/scripts/apply_format.py apply paper.docx paper-cjge.docx
```

```bash
python skills/docx-editor-cn/scripts/apply_format.py audit paper-cjge.docx
```

`"ok": true` and `"issue_count": 0` mean the whole format matches.

## Working with the MathType toolkit

For papers with MathType equations:

1. `apply_format.py apply` for the text formatting;
2. `render_mathtype_word_document` from [mathtype-office-toolkit](https://github.com/xyj0727/mathtype-office-toolkit) to create or restyle the equations in the CJGE equation format;
3. `apply_format.py apply` once more;
4. check equations with `validate_mathtype_word_document` and everything else with this server's `get_document_xml`.

## Changes in this fork

**MCP server (format configuration only)**

- New [`word_document_server/core/format_config.py`](word_document_server/core/format_config.py) holding the CJGE defaults.
- `create_document`: A4 with CJGE margins; document defaults and the Normal, Title, Heading 1–3 and Caption styles set to CJGE.
- `add_heading`: level 1–3 heading styles in CJGE format (SimHei 14 pt / SimHei 10.5 pt / SimSun 10.5 pt, not bold, black); the fallback sizes follow CJGE too.
- `add_table`: a three-line table with 9 pt text instead of the `Table Grid` style.
- Explicit arguments (`font_name`, `font_size`, `bold`, …) still take precedence.

**Skill (docx-editor-cn)**

- All format rules rewritten for CJGE: `SKILL.md`, `format_profiles/cjge.json` (replaces `docx_editor_cn.json`), `new_doc.js`, `table.py`.
- `apply_format.py` rewritten: recognises the CJGE parts, formats labels separately, splits one-column front matter from the two-column body, builds odd/even headers, formats first-page notes and audits through style inheritance.
- Removed `convert_paper.js`, a sample unrelated to the skill.

**Packaging**

- Package renamed to `cjge-word-toolkit` 1.0.0; the command is still `word_mcp_server`, so it is a drop-in replacement.
- Claude Code plugin manifests (`.claude-plugin/`); the original README moved to `docs/upstream-README.md`.

## Repository layout

```
cjge-word-toolkit/
├── word_document_server/          ← MCP server (original code)
│   └── core/format_config.py      ← CJGE defaults (added in this fork)
├── skills/docx-editor-cn/         ← skill (CJGE edition)
│   ├── SKILL.md
│   └── scripts/
│       ├── apply_format.py        ← whole-paper formatting / audit
│       ├── format_profiles/cjge.json
│       ├── new_doc.js             ← new manuscripts / Markdown conversion
│       └── table.py …
├── .claude-plugin/                ← Claude Code plugin manifests
├── docs/upstream-README.md        ← original README (full MCP tool reference)
├── pyproject.toml                 ← package cjge-word-toolkit
└── LICENSE
```

## Credits and license

- MCP server: [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server), MIT.
- Skill: docx-editor-cn by Gostyan, MIT (see `skills/docx-editor-cn/LICENSE.txt`).
- Format source: papers published in CJGE and the journal's author guidelines. This project is not affiliated with the journal; always follow the editorial office's current requirements.

Released under the MIT license, see [LICENSE](LICENSE).
