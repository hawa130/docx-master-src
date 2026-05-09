---
name: docx-master
description: "Standardize, edit, or audit a Word (.docx) document via direct OOXML mutation. Three commands: standardize (role-based whole-doc reshape — paragraph classification, named styles, multi-level auto-numbering, template import), edit (location-based surgical changes — replace/insert/delete paragraphs, table cells, image embedding, optional Word tracked changes), audit (read-only conformance check, no file written). Use whenever the user wants to format / restyle / normalize / edit / audit a Word document. Illustrative phrasings: '排版 / 套模板 / 按学校格式', '统一样式 / 标题字号不对', '把第N段改成… / 插一段 / 改这个单元格', '看看合不合规范'. Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx) OOXML directly: classify paragraph roles, inject named styles, migrate manual numbering, import template styles, surgically edit content, audit conformance. Output is a new docx plus a change report; the original file is never touched.

## Core Principle

**You are the analyst. The tools are your instruments.**

Tools present visible facts — computed styles, element positions, document structure. They don't classify or judge. Semantic reasoning is yours.

Two workflows compose every real task:

- `standardize` — reshape the doc's style system + numbering + role assignments. Operates by **pattern**: you describe categories of paragraphs (matching this regex / matching this fingerprint / sitting in this style); the engine applies uniformly.
- `edit` — surgical changes at **specific locations** (paragraph index, table cell, range). Used for content insertion and one-off corrections.

Compose by intent. A real task often needs both: standardize to install or fix the style system, then edit to insert or adjust content. Each command's reference doc covers its own scope. There is no separate "fill" workflow — when a user asks to "fill this template with X content", that's `standardize` (if structure needs work) then `edit`.

## Target state: structure-driven, not text-driven

A well-formed Word document expresses structural decisions through styles + numbering + sections, not through typed text that mimics structure. (Industry consensus: Microsoft, WebAIM, ECMA-376.) When the user has not pinned a contradicting choice, every standardize / edit pass produces a document with:

- Every paragraph carrying a semantic styleId. Direct paragraph format only as one-off exceptions.
- Structural hierarchy in **one unified multi-level numbering scheme** bound to Heading styles. Every installed Heading level binds to its corresponding scheme level (Heading1 → numLevel 0 + outlineLevel 0, …, HeadingN → numLevel N-1 + outlineLevel N-1). Manually-typed structural prefixes inside heading text — decimal hierarchy, locale numerals, parenthesized markers, chapter sentinels — converted via `stripPrefixPatterns`.
- Body lists bound to list-bound styles + a separate single-level numbering scheme. List markers never written as typed text.
- Heading levels nesting without skipping.

This applies to pre-existing chrome the template designer typed by hand and to source content the agent transcribes in. Manual structural prefixes are conversion targets, not chrome to preserve. Visual rendering is preserved (auto-numbering produces the same glyphs); logical structure is gained.

**Source-content → styleId mapping** (fixed; install missing pieces, don't substitute):

| Source content shape | Word styleId |
|---|---|
| Top-level heading | `Heading1` |
| Second-level heading | `Heading2` |
| Third-level heading | `Heading3` |
| Fourth-level heading | `Heading4` |
| Ordered list item | `ListNumber` |
| Bulleted list item | `ListBullet` |
| Code block | `Code` |
| Body paragraph | existing body style (`BodyText` / `Normal`-equivalent) |

The destination document type — form, report, thesis, contract, business memo — doesn't change this mapping. Specific surface markers (`#` / `1.1` / `一、` / `Chapter N` / Roman numerals) are variations on the same hierarchy.

**Locale defaults** (apply when the doc is in the named locale; pin otherwise via user instruction):

- **Chinese (CN) body text**: 2-character first-line indent (`firstLineIndent: "2char"`). Apply to body-class styles unless the user pins otherwise.
- **CJK ↔ Latin spacing**: strip literal spaces between CJK characters and Western characters / digits before emitting `text`. Source content (markdown, copied prose) often carries them for source readability; transcribing verbatim produces double gaps because Word's `autoSpaceDE` / `autoSpaceDN` add their own. Latin-only spans keep their internal spaces.

**Out of Phase 1 scope** (leave alone, surface to user if the limit blocks the task):

- Layout-**table** structure — restructuring or removing the layout table itself. Paragraphs *inside* layout-table cells, including any chrome the template designer typed by hand, are normal indexed paragraphs and fully restyleable; the table holding them stays a table.
- TOC body content — Word regenerates the field on open after `outlineLevel` is set.
- Cross-references, footnotes, comments, headers / footers — separate XML parts.

## How to standardize: describe by pattern, not by enumeration

The single most leveraged decision in `standardize`: **describe the change by category, not by listing paragraphs**.

`apply_styles` accepts four levels of paragraph targeting, in increasing specificity:

1. **`pattern_rules`** — regex matches paragraph text from the start. Engine applies to all matches uniformly. Best for chrome conversion (`^[一二三...]+、` / `^（[一二三...]）` / `^\d+\.\d+\s` / `^Chapter \d+` / etc.) and content-shape rules (`^Figure \d+` → `Caption`).
2. **`bulk_rules`** — match by visual fingerprint (font + size + spacing). Best when chrome shares a styleId / direct format but no consistent text pattern.
3. **`assignments`** — per-paragraph index. Reserved for **outliers** — single paragraphs that don't fit any pattern, or false-positive corrections.
4. **`exclude`** — paragraph indices to skip (after pattern + bulk match). Reserved for **false positives** the regex caught wrongly.

Default workflow for any standardize pass:

1. **Survey** via `overview`. Read the visual fingerprint summary, style definitions, numbering schemes, and skeleton. Identify what kinds of chrome and structural prefixes exist.
2. **Design ONE config** with `styles[]` (every Heading level the doc and content need), `numbering` as an array (multi-level heading scheme + single-level list scheme), `pattern_rules` (one rule per chrome shape, with `stripMatch: true`), and `bulk_rules` for fingerprint-keyed roles.
3. **Dry-run**, read the change report, refine. The report names the affected paragraphs per rule — you'll see false positives and gaps.
4. **Apply.** If filling content follows, hand off to `apply_edits`.

Per-paragraph enumeration via `assignments` is the **last resort**, not the default. When agents enumerate by index, every paragraph becomes an opportunity to silently skip; pattern-based rules don't have that failure mode.

See [references/standardize.md](references/standardize.md) for worked recipes and the full path catalogue.

## How to edit: insert content with semantic bindings

`apply_edits` operates at specific locators (paragraph index, range, cell, heading, whole-body). When inserting content:

- **Express structure semantically.** List items use `numbering: { numId, level }`; sub-headings use `styleId: "Heading3"` / etc. Don't type list markers or heading numbers as `text`.
- **Match-Destination-Formatting** picks up the anchor paragraph's pPr by default. Inspect the anchor's rPr first; if it carries unwanted bold (typically inherited from a label paragraph mark), override with `runFormat: { bold: false }` per Block.
- The styleId / numId you reference must exist. If they don't, that's a `standardize` task before this `edit` runs.

See [references/edit.md](references/edit.md) for locators, ops, blocks, and track-changes mode.

## Ask, don't decide (genuinely hard cases only)

Most strategy choices have a default — Target state pins them. Apply the default; don't ask just because a choice technically exists. Ask only when even the right semantic mapping is unclear:

- A typed chapter / section sentinel inside body prose — structural heading, or rhetorical citation? (`第N章` / `Chapter N` mentioned mid-paragraph vs. as a paragraph start.)
- A bold paragraph that could be a sub-heading or in-paragraph emphasis.
- Source content lacks coverage for some template slots — leave empty, generate, or surface to user?
- Content has tables / footnotes / math / cross-references with no clean Phase 1 mapping.

For these, send one focused message naming the choice + your default, and yield. (Subagents producing one final output: the output IS the question — return without executing.)

For everything else Target state covers, the default is the answer. Don't surface "I'll convert manual prefixes to auto-numbering" as a strategy question; just do it.

## Commands

| Command | Use when… | Reference |
|---|---|---|
| `standardize` | Style / numbering / structural changes applied across the doc, by **role** or **pattern**. *Illustrative phrasings: "帮我排一下版", "套学校格式", "Heading2 字号改小一号", "加个图注样式".* | [references/standardize.md](references/standardize.md) |
| `edit` | Surgical changes at **specific locations** — replace/insert/delete a paragraph, change a table cell, embed an image. *Illustrative phrasings: "把第 3 段改成 ...", "在 X 章后面插一段", "改这个单元格".* | [references/edit.md](references/edit.md) |
| `audit` | Read-only conformance check; no file output. *Illustrative phrasings: "看看这份合不合规范", "对照标准检查一下".* | [references/audit.md](references/audit.md) |

Real tasks often span both writing commands. Typical compositions: `standardize` first when the template's style system needs work, then `edit` to fill content; `audit` first to identify violations, then `standardize` to apply fixes.

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | First call on any task. Returns metadata, page setup (mm), theme, style definitions, numbering schemes (clustered by pattern), visual style statistics, and document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump for paragraphs with run-level mixed formatting. Shows which properties are uniform vs. mixed across runs. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph: nearby images, tables, equations, page breaks. **First choice for figure-caption / table-caption / first-after-heading classification.** |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | What role a fingerprint plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` and their `basedOn` chain. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections. |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Top-level tables with cell text snippets at `[row,col]`. Use before composing a `cell` locator on the `edit` path. |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Paragraphs `apply_edits` will refuse — existing tracked changes, complex fields, SDT controls. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. Use to **validate `pattern_rules` regex coverage before applying** — see exactly which paragraphs your regex catches. Also useful for content-defined role discovery. |
| `apply_styles --dry-run` | `node scripts/apply_styles.js --dry-run <config.json>` | Iterate on a config without writing the output file. Returns the full change report. **Use between every config edit.** |
| `apply_styles` | `node scripts/apply_styles.js <config.json>` | Combined orchestrator. Accepts styles[] + numbering (single or array of schemes) + template + pattern_rules + bulk_rules + assignments in one config. |
| `restyle` | `node scripts/restyle.js [--dry-run] <config.json>` | Narrow entry: paragraph restyle only — rejects `template` and `numbering`. |
| `migrate_numbering` | `node scripts/migrate_numbering.js [--dry-run] <config.json>` | Narrow entry: install / replace numbering. `styles[]` is optional. |
| `import_template` | `node scripts/import_template.js [--dry-run] <config.json>` | Narrow entry: import named styles from a template doc. |
| `apply_edits` | `node scripts/apply_edits.js <config.json>` | `edit` command entry. See [references/edit.md](references/edit.md). |

## Cross-command invariants

- **The original file is never modified.** Every applying CLI writes a fresh copy and validates it before keeping it; on validation failure the output is discarded and the original is untouched. Don't silently retry on validation errors — surface them.
- **Section properties (page size, margins, headers, footers, columns) are never modified.**
- **Paragraph indexing is 1-based**, matching `#NNN` labels in the skeleton. Layout-table paragraphs are indexed; data/form-table paragraphs are not (reachable on the `edit` path via cell locator).
- **Paths resolve against current working directory.** Use absolute paths if you may have changed directories during the session.
- **Restyle behavior:** run-level direct formatting that is *uniform across all runs* gets stripped on restyle; formatting that *differs between runs* is preserved as intentional inline emphasis.
- **Field codes** (`STYLEREF`, `TOC`, `REF`, `DATE`, …) are preserved as-is; this skill does not edit content inside fields.
- **TOC content is not regenerated.** Heading `outlineLevel` is set; the user must right-click → "Update Field" in Word after opening.
- **Edit blockers**: the `edit` command refuses to touch paragraphs inside existing tracked changes / complex field regions / SDT controls. Run `inspect_blockers` first.

Iterate with `--dry-run` between config edits — it's seconds per cycle and shows you the change report's effect.
