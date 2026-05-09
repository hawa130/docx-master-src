---
name: docx-master
description: "Standardize, edit, or audit a Word (.docx) document via direct OOXML mutation. Three commands: standardize (role-based whole-doc reshape — paragraph classification, named styles, multi-level auto-numbering, template import, content insertion), edit (location-based surgical changes — replace/insert/delete paragraphs, table cells, image embedding, optional Word tracked changes), audit (read-only conformance check, no file written). Use whenever the user wants to format / restyle / normalize / edit / audit a Word document. Illustrative phrasings: '排版 / 套模板 / 按学校格式', '统一样式 / 标题字号不对', '把第N段改成… / 插一段 / 改这个单元格', '看看合不合规范'. Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx) OOXML directly: produce well-formed documents from messy templates, install styles + numbering, convert hand-typed prefixes to auto-numbering, insert content, audit conformance. Output is a new docx; the original is never touched.

## Target state

A well-formed Word document expresses structure through **styles + numbering + sections**, not typed text mimicking structure (Microsoft / WebAIM / ECMA-376 consensus). Every write moves the document toward this destination:

- Every paragraph carries a semantic styleId (Heading1..N / BodyText / ListNumber / Caption / etc.); direct paragraph format only as one-off exceptions.
- One unified multi-level numbering scheme bound to all heading styles; one separate single-level scheme per list-bound style.
- Hierarchy and list markers come from auto-numbering — **never typed text** in either direction. Existing chrome stripped via `pattern_rules`; inserted content omits the prefix in `text`. Display follows the level's `lvlText` pattern (`%1` = own counter; `%1.%2` references multiple levels). See [`references/numbering-formats.md`](references/numbering-formats.md).
- Heading levels nest without skipping.
- **Match content shape to slot shape.** Prose slots (multi-line body, sections like abstract / introduction) take prose typography. Inline-value slots (single short phrase filling a labeled cell or single-line slot) inherit the slot's existing format — don't override with BodyText / first-line indent. List items default to body weight; bold only when the source explicitly emphasizes.
- **Locale-specific typography.** CJK docs: prose body and list items typically take a 2-char first-line indent (so the list marker aligns with body's first character — flush-left looks foreign in CN academic conventions); literal whitespace between CJK and Latin runs is stripped (Word's autoSpace handles the gap). Chinese font-size names: [`references/chinese-font-sizes.md`](references/chinese-font-sizes.md).

## Tools you are the analyst

Tools surface visible facts; classification and judgment are yours.

## Commands

| Command | When | Reference |
|---|---|---|
| **`apply`** | The unified writer. Install styles + numbering + theme + template, restyle by pattern / fingerprint, insert content via edits — single config, single call. Pure-edit tasks use `apply` with just `edits[]`. | [standardize.md](references/standardize.md) + [edit.md](references/edit.md) |
| `audit` | Read-only conformance check; no file output. | [audit.md](references/audit.md) |

`apply` pipeline order:

```
install styles + numbering + theme + template
  → run edits (referencing just-installed styleIds)
  → re-fingerprint
  → run rules (pattern_rules / bulk_rules / assignments / exclude —
    match BOTH pre-existing chrome AND agent-inserted content uniformly)
  → validate, write
```

Sparse by design — only declared blocks apply. Untouched styles / numbering / paragraphs / theme stay as they are. **Declare only what's wrong, missing, or the user explicitly asked to change** — existing chrome that already meets the target state (correct styles / numbering / typography) is not your business to normalize. Match existing conventions; fill gaps; route unrouted body fingerprints. Don't rewrite what already works.

## Workflow

### 1. Survey

`overview` first. From the output, note:

- Existing Heading levels, numbering schemes, fingerprints, document skeleton
- **Typography requirements baked into the document** — paragraphs that prescribe formatting (fonts, sizes, line spacing, alignment). Often live in front matter, instruction blocks, footnotes, or "filling guidelines" sections. These override any defaults you would invent. Skim before designing styles.
- Typed structural prefixes in chrome — chapter / section markers like `Chapter N.`, `第N章`, `一、`, `1.1`, `（一）`, etc. List the patterns you see; one regex per shape will go into `pattern_rules`.
- **Form-fill paragraphs** — paragraphs whose visible text looks like `label + long whitespace gap`, `label + ____ underscore placeholder`, or several `label / gap / label` pairs in one paragraph. The blank is usually a separate run carrying `<w:u/>`. Run `inspect_runs` on these before deciding the edit shape; whole-paragraph `replace` destroys the placeholder structure.
- (Filling) What content the source carries

### 2. Design ONE config

Sketch the content's structural outline first; `styles[]` follows that outline, not the other way around. Reactive style additions accrete debt later edits have to re-untangle.

- `styles[]` — every Heading level the combined doc + content needs; `BodyText`; **`ListNumber` + single-level numbering scheme for any body list content**. Do NOT type list markers (`(1)(2)(3)`, `①②③`, `1. 2.`, `- `) in inserted text — Word renders the literal characters, not a list. **Sizes follow the document's existing chrome** (visual style summary in `overview`): if heading-shaped paragraphs share a uniform size with hierarchy signalled by weight / spacing / indent, preserve that — differentiate levels on the same axes the chrome uses. Invent sizes only when the document has no chrome convention.
- `numbering` — one multi-level scheme bound to Heading1..N; one single-level per list-bound style.
- `pattern_rules` — one regex per chrome shape with `stripMatch: true`. Applies uniformly to every match.
- `edits[]` — content insertion. For **form-fill paragraphs** identified in Step 1, use cell-content insert or run-level surgery, not whole-paragraph `replace`.
- `bulk_rules` — fingerprint-keyed routing. **Every body-content fingerprint not matched by `pattern_rules` needs a `bulk_rules` entry** — otherwise existing direct-formatted paragraphs won't pick up the new style cascade and your `firstLineIndent` / `fontCJK` etc. won't apply to them.
- `exclude` — false-positive corrections.
- `assignments` — per-paragraph corrections, **last resort**, for outliers only.

### 3. Dry-run

`apply --dry-run`. Read the change report:

- **Style Resolution** — each installed style matches user spec / extracted source / template-prescribed values
- **Sample Affected Paragraphs** — `pattern_rules` hit the right targets (`find_paragraphs --regex` validates coverage before apply)
- **Implicit-keep is a FAILURE signal** — non-empty fingerprints not routed by any rule. Add `bulk_rules` / `assignments` until implicit-keep is empty, OR until the remaining ones are intentional non-content (true spacers, blank-line slots).

### 4. Apply

Output is a fresh docx; the original is never modified.

## Asking the user

Most strategy choices have a default — Target state pins them. Apply the default. Ask only when the right semantic mapping is genuinely unclear: typed sentinel mid-prose vs. heading; bold paragraph as sub-heading vs. emphasis; missing source content for template slots; unsupported structures (footnotes, math, cross-references). Send one focused message naming the choice + your default, then yield. Subagent producing one final output: the output IS the question — return without executing.

## Out of Phase 1 scope

Surface to the user when blocked by: layout-**table** restructuring, TOC body content, footnotes / comments / headers / footers content. Paragraphs *inside* layout-table cells ARE indexed and fully editable; only the table holding them is off-limits.

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | First call on any task. Metadata, page setup (mm), theme, style defs, numbering schemes (clustered by pattern), visual style statistics, document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump. Use for paragraphs with mixed run-level formatting OR **form-fill segments** (label + underscore-blank pattern) — see how the blank is structured before deciding the edit shape. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph. First choice for figure-caption / table-caption / first-after-heading classification. |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | What role a fingerprint plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` and their `basedOn` chain. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections. |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Top-level tables with cell text snippets at `[row,col]`. Use before composing a `cell` locator. |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Paragraphs `apply`'s edit phase will refuse — tracked changes, complex fields, SDT controls. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. Validate `pattern_rules` regex coverage before applying. |
| `validate` | `node scripts/validate.js <file>` | Schema-aware OOXML check. `apply` runs this automatically; standalone for spot-checking arbitrary .docx files. |
| `apply --dry-run` | `node scripts/apply.js --dry-run <config.json>` | Iterate on a config without writing. **Use between every config edit.** |
| `apply` | `node scripts/apply.js <config.json>` | Unified writer. Default for any write task. |
| `restyle` | `node scripts/restyle.js [--dry-run] <config.json>` | Narrow: paragraph restyle only. |
| `migrate_numbering` | `node scripts/migrate_numbering.js [--dry-run] <config.json>` | Narrow: numbering install only. |
| `import_template` | `node scripts/import_template.js [--dry-run] <config.json>` | Narrow: template style import only. |

## Cross-command invariants

- Original file is never modified; every applying CLI writes a fresh copy + validates before keeping. Validation failure → discard, surface to user — don't silently retry.
- Section properties (page size, margins, headers, footers, columns) never modified.
- Paragraph indexing is 1-based, matching `#NNN` in skeleton. Layout-table paragraphs are indexed; data / form-table paragraphs aren't (reachable via cell locator on edit path).
- Paths resolve against CWD; use absolute paths if you may have changed directories.
- Restyle: run-level direct formatting uniform across all runs gets stripped; per-run differences preserved as intentional inline emphasis.
- Field codes (`STYLEREF` / `TOC` / `REF` / `DATE` / …) preserved as-is; not editable inside.
- TOC content is not regenerated; user must right-click → "Update Field" in Word after opening.
- Edit blockers: the edit phase refuses paragraphs inside existing tracked changes / complex fields / SDT controls. Run `inspect_blockers` first.

Block-level config details: [`references/standardize.md`](references/standardize.md) (styles / numbering / pattern_rules / bulk_rules / assignments / exclude), [`references/edit.md`](references/edit.md) (locators / ops / MDF / track-changes), [`references/config-schema.md`](references/config-schema.md) (full field reference).
