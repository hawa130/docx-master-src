---
name: docx-master
description: "Standardize, edit, or audit a Word (.docx) document via direct OOXML mutation. Three commands: standardize (role-based whole-doc reshape — paragraph classification, named styles, multi-level auto-numbering, template import, content insertion), edit (location-based surgical changes — replace/insert/delete paragraphs, table cells, image embedding, optional Word tracked changes), audit (read-only conformance check, no file written). Use whenever the user wants to format / restyle / normalize / edit / audit a Word document. Illustrative phrasings: '排版 / 套模板 / 按学校格式', '统一样式 / 标题字号不对', '把第N段改成… / 插一段 / 改这个单元格', '看看合不合规范'. Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx) OOXML directly: produce well-formed documents from messy templates, install styles + numbering, convert hand-typed prefixes to auto-numbering, insert content, audit conformance. Output is a new docx; the original file is never touched.

## Target state — the highest principle

A well-formed Word document expresses structural decisions through **styles + numbering + sections**, not through typed text that mimics structure. (Industry consensus: Microsoft, WebAIM, ECMA-376.) **This is the destination — every write operation in this skill exists to move the document toward this state.**

The required output shape:

- Every paragraph carries a semantic styleId (Heading1 / Heading2 / … / BodyText / ListNumber / Caption / etc.). Direct paragraph format only as one-off exceptions.
- Structural hierarchy in **one unified multi-level numbering scheme** bound to Heading styles. Every installed Heading level binds to its scheme level (Heading1 → numLevel 0 + outlineLevel 0, …, HeadingN → numLevel N-1 + outlineLevel N-1).
- Body lists bound to list-bound styles (`ListNumber` / `ListBullet`) + a separate single-level numbering scheme.
- **All hierarchy and list markers come from auto-numbering — never typed text, in either direction.** Pre-existing typed chrome is converted via `stripPrefixPatterns`. Newly inserted heading or list content omits the prefix from its `text` — Word renders the marker from the bound numbering level. The display format follows the `lvlText` pattern (`%1` references the level's own counter; composite forms like `%1.%2` reference multiple levels). See `references/numbering-formats.md` for level-shape recipes covering common patterns.
- Heading levels nest without skipping.
- Locale typography defaults applied (CN body: 2-character first-line indent; CJK ↔ Latin literal spaces stripped — Word's autoSpace handles the gap). For Chinese font-size names (初号 / 一号 / … / 小六) see [`references/chinese-font-sizes.md`](references/chinese-font-sizes.md).

## Tools you are the analyst

Tools present visible facts — computed styles, element positions, document structure. They don't classify or judge. Semantic reasoning is yours.

## Commands

| Command | When | Reference |
|---|---|---|
| **`apply`** | The unified writer. Install / override styles, install numbering, set theme, import templates, restyle by pattern / fingerprint, insert content via edits. Single config, single call — even when task spans installing structure + filling content. Pure-edit tasks (locator-based, no style install) also use `apply` with just `edits[]` in the config. | [references/standardize.md](references/standardize.md) + [references/edit.md](references/edit.md) |
| `audit` | Read-only conformance check; no file output. | [references/audit.md](references/audit.md) |

`apply` runs config blocks in this internal order:

```
install styles + numbering + theme + template
  → apply edits (insert content, referencing the now-installed styles)
  → re-fingerprint paragraphs
  → apply rules (pattern_rules / bulk_rules / assignments / exclude — these
    match BOTH pre-existing chrome paragraphs AND any agent-inserted content
    uniformly; one regex matches both kinds of typed prefixes)
  → validate, write
```

**Sparse by design** — only declared blocks apply. Untouched styles, numbering, paragraphs, theme stay as they are.

Pure content-only edits (no style install) still go through `apply` with `edits[]` only; there is no separate edit-only CLI. The narrower CLIs (`restyle` / `migrate_numbering` / `import_template`) are filtered views of the same engine — see Tool Reference below.

## How to apply (the default workflow)

For any task that produces output:

1. **Survey** via `overview` — read style definitions, numbering schemes, fingerprint summary, document skeleton. Identify:
   - What Heading levels and other styles already exist.
   - What numbering schemes are installed and how many levels they cover.
   - Where typed structural prefixes appear (chrome `一、` / `（一）` / `第N章` / `1.1`, etc.).
   - What the source content carries (when filling).

2. **Design ONE config** combining everything the pass needs. When filling, sketch the content's structural outline first — `styles[]` follows that outline, not the other way around. Reactive style additions mid-write accrete debt that later edits have to re-untangle.
   - `styles[]` — every Heading level the doc + content combined need; List/Caption/Body styles as required. Adjacent levels need a visible typographic gradient (size / weight / spacing); identical styling defeats the structural signal.
   - `numbering` — array of schemes: multi-level for headings, single-level per list-bound style.
   - `pattern_rules` — one regex per chrome shape with `stripMatch: true`. Engine applies uniformly to every match — you can't selectively skip matched paragraphs.
   - `edits[]` — content insertion if needed (when filling a template).
   - `bulk_rules` — fingerprint-keyed assignments for paragraphs without a clean text pattern.
   - `exclude` — false-positive corrections.
   - `assignments` — per-paragraph corrections, **last resort**, used only for outliers.

3. **Dry-run** with `apply --dry-run`. Read the change report:
   - Style Resolution: does each installed style match user spec / extracted source?
   - Sample Affected Paragraphs: did `pattern_rules` hit the right targets?
   - Implicit-keep: any non-empty fingerprint not yet routed?

4. **Apply.** Output is a fresh docx; the original is never modified.

## Combining standardize + edits in one call

A template-fill task usually needs both shapes: install the style system AND insert content. Compose them into one `apply` config:

```jsonc
{
  "source": "...", "output": "...",
  "styles":        [ /* Heading1..N + ListNumber + ... */ ],
  "numbering":     [ /* multi-level heading scheme + single-level list scheme */ ],
  "pattern_rules": [ /* one entry per typed chrome shape */ ],
  "edits":         [ /* content insertion ops */ ]
}
```

Pipeline order: install styles + numbering → run edits (referencing the just-installed styleIds) → re-fingerprint → run rules (cleans BOTH pre-existing chrome AND any typed prefixes the agent emitted in content) → validate, write.

Block-level details: [`references/standardize.md`](references/standardize.md) for `styles[]` / `numbering` / `pattern_rules` / `bulk_rules` / `assignments` / `exclude`. [`references/edit.md`](references/edit.md) for `edits[]` (locators, ops, MDF, track-changes). [`references/numbering-formats.md`](references/numbering-formats.md) for level-shape recipes. [`references/config-schema.md`](references/config-schema.md) for full field reference.

## Out of Phase 1 scope

These are accessibility / portability anti-patterns in the abstract, but Phase 1 has no operation to fix them. Surface to the user if a limit blocks the task:

- Layout-**table** structure — restructuring or removing the layout table itself. Paragraphs *inside* layout-table cells, including any chrome the template designer typed by hand, are normal indexed paragraphs and fully restyleable; the table holding them stays a table.
- TOC body content — Word regenerates the field on open after `outlineLevel` is set on Heading styles.
- Cross-references, footnotes, comments, headers / footers — separate XML parts not addressable by Phase 1 ops.

## Ask only when truly uncertain

Most strategy choices have a default — Target state pins them. Apply the default. Ask only when even the right semantic mapping is genuinely unclear:

- A typed chapter / section sentinel inside body prose — structural heading, or rhetorical citation? (e.g. `第N章` mentioned mid-paragraph vs. as a paragraph start.)
- A bold paragraph that could be a sub-heading or in-paragraph emphasis.
- Source content lacks coverage for some template slots — leave empty, generate, or surface to user?
- Content has tables / footnotes / math / cross-references with no clean Phase 1 mapping.

For these, send one focused message naming the choice + your default, and yield. (Subagents producing one final output: the output IS the question — return without executing.)

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | First call on any task. Returns metadata, page setup (mm), theme, style definitions, numbering schemes (clustered by pattern), visual style statistics, and document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump for paragraphs with run-level mixed formatting. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph. **First choice for figure-caption / table-caption / first-after-heading classification.** |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | What role a fingerprint plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` and their `basedOn` chain. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections. |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Top-level tables with cell text snippets at `[row,col]`. Use before composing a `cell` locator on the `edit` path. |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Paragraphs `apply`'s edit phase will refuse — existing tracked changes, complex fields, SDT controls. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. **Use to validate `pattern_rules` regex coverage before applying** — see exactly which paragraphs your regex catches. |
| `validate` | `node scripts/validate.js <file>` | Schema-aware OOXML check. Surfaces element-ordering / required-element issues that make Word prompt to "repair" the file. `apply` runs this automatically on every write; use the standalone CLI to spot-check arbitrary .docx files. |
| `apply --dry-run` | `node scripts/apply.js --dry-run <config.json>` | Iterate on a config without writing the output file. **Use between every config edit.** |
| `apply` | `node scripts/apply.js <config.json>` | Unified writer. Accepts styles + numbering + template + theme + rules + edits in one config. **Default for any write task** — including pure-edit (just `edits[]`) and pure-restyle. |
| `restyle` | `node scripts/restyle.js [--dry-run] <config.json>` | Narrow: paragraph restyle only. |
| `migrate_numbering` | `node scripts/migrate_numbering.js [--dry-run] <config.json>` | Narrow: numbering install only. |
| `import_template` | `node scripts/import_template.js [--dry-run] <config.json>` | Narrow: template style import only. |

## Cross-command invariants

- **The original file is never modified.** Every applying CLI writes a fresh copy and validates it before keeping it; on validation failure the output is discarded. Don't silently retry on validation errors — surface them.
- **Section properties (page size, margins, headers, footers, columns) are never modified.**
- **Paragraph indexing is 1-based**, matching `#NNN` labels in the skeleton. Layout-table paragraphs are indexed; data/form-table paragraphs are not (reachable on the `edit` path via cell locator).
- **Paths resolve against current working directory.** Use absolute paths if you may have changed directories during the session.
- **Restyle behavior:** run-level direct formatting that is *uniform across all runs* gets stripped on restyle; formatting that *differs between runs* is preserved as intentional inline emphasis.
- **Field codes** (`STYLEREF`, `TOC`, `REF`, `DATE`, …) are preserved as-is; this skill does not edit content inside fields.
- **TOC content is not regenerated.** Heading `outlineLevel` is set; the user must right-click → "Update Field" in Word after opening.
- **Edit blockers**: the `edit` command refuses to touch paragraphs inside existing tracked changes / complex field regions / SDT controls. Run `inspect_blockers` first.

Iterate with `--dry-run` between config edits — it's seconds per cycle and shows the change report's effect.
