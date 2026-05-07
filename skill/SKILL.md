---
name: docx-master
description: "Standardize, edit, or audit a Word (.docx) document via direct OOXML mutation. Three commands: standardize (role-based whole-doc reshape — paragraph classification, named styles, multi-level auto-numbering, template import), edit (location-based surgical changes — replace/insert/delete paragraphs, table cells, image embedding, optional Word tracked changes), audit (read-only conformance check, no file written). Use whenever the user wants to format / restyle / normalize / edit / audit a Word document. Illustrative phrasings: '排版 / 套模板 / 按学校格式', '统一样式 / 标题字号不对', '把第N段改成… / 插一段 / 改这个单元格', '看看合不合规范'. Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx) OOXML directly: classify paragraph roles, inject named styles, migrate manual numbering, import template styles, surgically edit content, audit conformance. Output is a new docx plus a change report; the original file is never touched.

## Core Principle

**You are the analyst. The tools are your instruments.**

Tools only present facts — computed styles, element positions, document structure. They never classify or judge. All semantic reasoning is yours: deciding what role a paragraph plays, what style to name it, whether two similar formats should merge or stay separate, how to handle edge cases.

## Commands

Pick the command that matches **what the user wants the document to become after the operation** — not what words they used. Same surface phrasing can land in different commands depending on intent; the examples illustrate, they're not triggers.

| Command | Use when… | Reference |
|---|---|---|
| `standardize` | The user wants style / numbering / structural changes applied across the doc, by **role** (heading, body, caption …). Covers full-doc reshape, narrow style-system edits, and the manual-XML escape hatch. *Illustrative phrasings: "帮我排一下版", "套学校格式", "加个 X 样式 / 其他不动", "Heading2 字号改小一号".* | [references/standardize.md](references/standardize.md) |
| `edit` | The user wants surgical changes at **specific locations** — replace/insert/delete a paragraph, change a table cell, embed an image, restyle one paragraph or range. Optional Word tracked-changes mode. *Illustrative phrasings: "把第 3 段改成 ...", "在第 X 章后面插一段", "这个表格第 2 行换成 ...", "给第 5 段加粗", "在结论后插张图".* | [references/edit.md](references/edit.md) |
| `audit` | The user wants a read-only conformance check against a typography spec; no file output. *Illustrative phrasings: "看看这份合不合学校规范", "对照这个标准检查一下".* | [references/audit.md](references/audit.md) |

When intent is genuinely ambiguous (mixed scope, broad request that could be standardize or edit), prefer asking one focused question over guessing.

The split between `standardize` and `edit`: standardize works on **roles** (every H2 in the doc, every body paragraph, every figure caption); edit works on **locations** (paragraph #17, the cell at row 3 col 2 of table 1). If the user names specific paragraph indices or table cells, that's edit. If they describe a class of paragraphs, that's standardize.

### Composing scopes

The commands above are **scopes of mutation, not exclusive paths**. Real tasks often span more than one — compose by intent. Three concrete shapes:

- **Messy template + content to fill**: `standardize` first to install a clean style system, then `edit` to insert content. Filling a dirty template directly with `edit` propagates the mess (Match-Destination-Formatting inherits whatever's there, including bad styles).
- **Audit then fix**: `audit` (read-only) produces a violation list; `standardize` applies the fixes. Don't auto-fix without permission.
- **Standardize then touch up**: bulk role-based reshape via `standardize`, then a few `edit` ops for paragraphs the rules missed or for content insertion the spec implied.

Markdown content as input has no adapter yet — agent translates MD → Block JSON manually. Tables, footnotes, math, and cross-references in MD have no clean docx mapping in Phase 1; degrade or surface to the user.

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | First call on Full Standardization or Audit. Returns metadata, page setup (mm), theme, style definitions, numbering schemes (clustered by pattern), visual style statistics, and document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump when a paragraph has run-level mixed formatting (bold lead phrase, colored numbering prefix, inline emphasis). Tells you which properties are uniform vs. mixed across runs — critical for predicting what `apply_styles` will preserve vs. strip. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph: nearby images, tables, equations, page breaks, sibling paragraphs. Default radius 4. **First choice for figure-caption / table-caption / first-after-heading classification.** |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | When you see a fingerprint in the overview and need to understand what role it plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` — what they currently are and what `basedOn` chain produces them. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections (headers, footers, page numbering). |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Lists top-level tables with cell text snippets at `[row,col]` coordinates. Use before composing a `cell` locator on the `edit` path. Includes data/form tables (which paragraph indices skip). |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Lists paragraphs that `apply_edits` will refuse to touch — existing tracked changes, complex field regions (TOC / STYLEREF / cross-references), SDT content controls. Use before composing edit ops to plan around them. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. Use to discover content-defined roles (figure/table captions, references, keywords) and to validate `pattern_rules` regex coverage before applying. **First choice on the Targeted Edit path** — it locates exactly the paragraphs you want to change, without requiring a full overview. |
| `apply_styles --dry-run` | `node scripts/apply_styles.js --dry-run <config.json>` | Iterate on a config without writing the output file. Returns the full change report including sample affected paragraphs and the Style Resolution annotation block. **Use between every config edit** — it's seconds per cycle. |
| `apply_styles` | `node scripts/apply_styles.js <config.json>` | Combined orchestrator. Accepts styles[] + numbering + template in one config and runs them in the correct order. Use on the Full Standardization path. |
| `restyle` | `node scripts/restyle.js [--dry-run] <config.json>` | Narrow entry: paragraph restyle only — rejects `template` and `numbering`. Same config schema as `apply_styles` minus those blocks. **First choice on the Targeted Edit path** when the change is purely a style assignment. |
| `migrate_numbering` | `node scripts/migrate_numbering.js [--dry-run] <config.json>` | Narrow entry: install / replace a numbering scheme. `styles[]` is optional — numbering levels can target heading styles already defined in the doc's `styles.xml` without redeclaration. Use to add or change heading numbering on a doc that's otherwise correctly styled. |
| `import_template` | `node scripts/import_template.js [--dry-run] <config.json>` | Narrow entry: import named styles from a template doc. `styles[]` is optional. Use when you want to pull in a template's style system without simultaneously restyling — typical first step before chaining `restyle`. |
| `apply_edits` | `node scripts/apply_edits.js <config.json>` | `edit` command entry. Applies an ordered list of `replace` / `insert-before` / `insert-after` / `delete` / `format` ops at specific locators (paragraph index, range, cell, heading, whole-body). Optional `trackChanges: true` emits Word revision markup. See [references/edit.md](references/edit.md). |

## Cross-command invariants

These hold regardless of which command you're running:

- **The original file is never modified.** Every applying CLI writes a fresh copy and validates it before keeping it; on validation failure the output is discarded and the original is untouched. Don't silently retry on validation errors — surface them.
- **Section properties (page size, margins, headers, footers, columns) are never modified.**
- **Paragraph indexing is 1-based**, matching `#NNN` labels in the skeleton. Layout-table paragraphs are indexed; data/form tables are not (reachable on the `edit` path via cell locator).
- **Edit blockers**: the `edit` command refuses to touch paragraphs inside existing tracked changes / complex field regions (TOC/STYLEREF/REF) / SDT content controls. Run `inspect_blockers` to see which paragraphs are off-limits before composing edits.
- **Paths resolve against current working directory.** `source` and `output` are passed through `path.resolve()` against cwd; absolute paths are passed through unchanged. If you may have changed directories during the session, use absolute paths to avoid surprises.
- **Restyle behavior:** when a paragraph is restyled, run-level direct formatting that is *uniform across all runs* gets stripped (so the new style's defaults take effect). Direct formatting that *differs between runs* is preserved as intentional inline emphasis — bold lead phrase + non-bold body, colored numbering prefix + black title, etc. survive automatically.
- **Field codes** (`STYLEREF`, `TOC`, `REF`, `DATE`, …) are preserved as-is; this skill does not edit content inside fields.
- **TOC content is not regenerated.** Heading `outlineLevel` is set correctly, but the user must right-click → "Update Field" in Word after opening to refresh the TOC body.

Iterating with `--dry-run` between edits is the supported workflow, not a fallback. Configs can be sparse: declare only what you're touching.
