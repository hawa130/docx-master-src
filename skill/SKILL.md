---
name: docx-master
description: "Standardize, edit, or audit a Word (.docx) document via direct OOXML mutation. Two CLIs: `apply` (the unified writer — install styles + numbering + theme + template, restyle by pattern / fingerprint, insert content via edits, all in one config), `audit` (read-only conformance check, no file written). Common config shapes inside `apply`: standardize (role-based whole-doc reshape — paragraph classification, named styles, multi-level auto-numbering, template import), edit (location-based surgical changes — replace/insert/delete paragraphs, table cells, image embedding, optional tracked changes). Use whenever the user wants to format / restyle / normalize / edit / audit a Word document. Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# docx-master

Mutates Word (.docx) OOXML directly: produce well-formed documents from messy templates, install styles + numbering, convert hand-typed prefixes to auto-numbering, insert content, audit conformance. Output is a new docx; the original is never touched.

## How to think about formatting

User prompt overrides everything below.

Otherwise, treat a document as two kinds of paragraphs:

- **Content chrome** — paragraphs with typed text already in the document (chapter markers, instruction notes, fixed labels). Three distinct operations apply:
  - **Preserve typography** (font / size / spacing / indent — the direct `pPr` / `rPr`). Don't declare those attributes on the styles that re-tag these paragraphs; engine's selective-strip keeps them in place as long as the style doesn't override.
  - **Retag structure** (bind `pStyle`, strip typed structural prefixes via `pattern_rules`, install matching `numbering`). Chapter / heading / list markers are structure, not typography — they belong to auto-numbering. Existing typed prefixes like `一、` / `（一）` / `1.1` are the **input** to retagging, not output to preserve.
  - **Don't replace, don't duplicate.** A chrome paragraph stays as itself — never the target of `replace`, never quoted into an inserted paragraph's `text`. Content insertion lands in the empty placeholder slot beside the chrome (`set-run` on inline blanks; `replace` against the empty slot for separate-paragraph forms). Pulling a chrome label into an inserted paragraph's text strips the label's typography along with the original paragraph.
- **Empty slots** — blank paragraphs pre-allocated for content insertion. **Bind inserted content to a semantic style** (`BodyText` for prose, `ListNumber` / `ListBullet` for block enumerations, `Heading1..N` for sub-sections); install the style in `styles[]` if the doc doesn't already have one that fits. `Normal` is the fallback style, not a body style — don't bind body prose to it.

The practical rule for `styles[]`: one style per semantic role — reuse the source's existing styleId when one fits, install fresh only when source has no fit (or chaotic source bundles many roles on one styleId). **For every entry whose role has a representative paragraph in the source, use `fromParagraph: N`** — the engine extracts the template's typography automatically; add only `outlineLevel` / `basedOn` and user-spec overrides via `overrides: { ... }`. Top-level explicit `size` / `bold` / `alignment` / `spaceBefore` / `lineSpacing` / etc. are reserved for empty-slot roles (no source paragraph plays the role) or fresh styles the user explicitly speced — declaring them on a represented role silently overrides the template.

Tools surface visible facts; classification and judgment are yours.

## Target state

A well-formed Word document expresses structure through **styles + numbering + sections**, not typed text mimicking structure (Microsoft / WebAIM / ECMA-376 consensus). When filling empty slots, move toward:

- Every paragraph carries a semantic styleId (Heading1..N / BodyText / ListNumber / Caption / etc.); direct paragraph format only as one-off exceptions.
- One unified multi-level numbering scheme bound to all heading styles; one separate single-level scheme per list-bound style.
- Every paragraph-leading counter comes from auto-numbering — heading prefix, **in-prose enumeration marker**, caption number, reference number. **Never typed in `text`**, regardless of shape (CJK ordinal, parenthesized digit, full-width period, circled, alphabetical, etc.). Existing chrome stripped via `pattern_rules`; inserted content omits the prefix and binds via `styleId` to a numbering scheme. See [`references/numbering-formats.md`](references/numbering-formats.md).
- Heading levels nest without skipping.
- **Match content shape to slot shape.** List items default to body weight; bold only when the source explicitly emphasizes:
  - **Prose** (multi-paragraph body) → prose typography.
  - **Inline-value** (short phrase filling a labeled cell) → inherit the slot's existing format.
  - **Block enumeration** → `ListNumber` + single-level scheme; markers from the scheme.
  - **Figure** → `FigureImage` (image paragraph; line height stays flexible so the drawing isn't clipped) + `CaptionBlock` with `captionId: "Figure"` **below** — declare the identifier in the top-level [`captions`](references/captions.md) table.
  - **Table** → `{ "type": "table", ... }` block + `CaptionBlock` with `captionId: "Table"` **above** — see [references/tables.md](references/tables.md).
  - **Equation** → `{ "type": "equation", "latex": "...", "captionId": "Equation" }` for a numbered equation (engine emits the 3-column borderless layout automatically); omit `captionId` for unnumbered. Inline math uses `{ "math": "..." }` in Paragraph.text — see [references/equations.md](references/equations.md).
- **Caption mechanism.** Caption-class numbering (figure / table / equation / theorem / ...) goes through the top-level `captions` config — declare an entry per identifier with prefix / suffix / chapter prefix / styleId, then reference `captionId` on `CaptionBlock` / `EquationBlock`. Word emits SEQ + STYLEREF fields; cross-references use REF `\h`. Do NOT bind caption-class styleIds to `numbering[]` — see [references/captions.md](references/captions.md).
- **Locale-specific typography.** CJK prose body and list items get a 2-char first-line indent — declare `firstLineIndent: "2char"` on `BodyText` / `ListNumber`. Chinese font-size names: [`references/chinese-font-sizes.md`](references/chinese-font-sizes.md).
- **No spaces at CJK ↔ Latin / digit boundaries.** Applies at every boundary type — inside one `text` run, between adjacent runs, around `InlineRef`, around `{ math }`. Word's autoSpace inserts the gap automatically; a typed space stacks on top of it and renders as a visibly wide gap.

## Commands

| Command | When | Reference |
|---|---|---|
| **`apply`** | The unified writer. Install styles + numbering + theme + template, restyle by pattern / fingerprint, insert content via edits — single config, single call. | [standardize.md](references/standardize.md) + [edit.md](references/edit.md) |
| `audit` | Read-only conformance check workflow. No CLI; uses the inspect tools to produce a violation report. | [audit.md](references/audit.md) |

`apply` pipeline order:

```
install styles + numbering + theme + template
  → run edits (referencing just-installed styleIds)
  → re-fingerprint
  → run rules (pattern_rules / bulk_rules / assignments / exclude —
    match BOTH pre-existing chrome AND agent-inserted content uniformly)
  → validate, write
```

Sparse by design — only declared blocks apply; untouched styles / numbering / paragraphs / theme stay as they are. **Declare only what's wrong, missing, or what the user explicitly asks to change.**

## Workflow

### 1. Survey

`overview` first. From the output, note:

- **Existing structure** — Heading levels, numbering schemes, fingerprints, document skeleton.
- **Content chrome formatting** — what direct `pPr` / `rPr` content paragraphs carry (font / size / line spacing / alignment / indent). This IS the document's typographic convention; you'll preserve it by not declaring those attributes on the styles that re-tag these paragraphs.
- **Typed structural prefixes** — chapter / section markers like `Chapter N.`, `第N章`, `一、`, `1.1`, `（一）`. List the shapes you see; one regex per shape feeds `pattern_rules`.
- **Form-fill paragraphs** — text shaped like `label + long whitespace gap`, `label + ____ underscore placeholder`, or several `label / gap / label` pairs in one paragraph. The blank is usually a separate run with `<w:u/>`. Note indices for Step 2.
- **Source content** (fill tasks) — what the user-provided content actually carries.

### 2. Design ONE config

Plan styles first based on the content's structural outline. Then route in one config: chrome → apply those styles via `pattern_rules` / `assignments` / `bulk_rules`; inserts → bind via `styleId` on `edits[]` Blocks. Reactive style additions later accrete debt.

- `styles[]` — install one style per semantic role the doc + content combined contain (Heading1..N, `BodyText`, `ListNumber`, `FigureImage` + `FigureCaption`, `TableCaption`, `EquationNumber`, `Reference`, etc.). Use `fromParagraph` for represented roles; Mode B with declared fields only for empty-slot or user-spec-named — see "How to think about formatting" above.
- `numbering` — one multi-level scheme bound to Heading1..N; one single-level per list-bound style. The scheme's `lvlText` chooses the marker shape (decimal / parenthesized / CJK 序号 / bullet / ...) — pick to match document convention or user request.
- `pattern_rules` — one regex per chrome shape with `stripMatch: true`. Applies uniformly to every match.
- `edits[]` — content insertion. New `paragraph` Blocks bind to the styles you installed via `styleId`; per-op `paraFormat` / `runFormat` are one-off overrides, not the default. For **form-fill paragraphs** identified in Step 1, use the `set-run` op with a `run` locator (`blank: K` for Kth blank placeholder) — preserves the placeholder run's rPr automatically; whole-paragraph `replace` is the wrong tool there.
- `bulk_rules` — fingerprint-keyed routing for body paragraphs without a clean text pattern. Whether to route a given fingerprint is a Step 3 question — `dry-run` flags unrouted ones via implicit-keep, and you decide there.
- `exclude` — false-positive corrections.
- `assignments` — per-paragraph corrections, **last resort**, for outliers only.

### 3. Dry-run

`apply --dry-run`. Read the change report:

- **Style Resolution** — each installed style matches user spec / extracted source / template-prescribed values
- **Sample Affected Paragraphs** — `pattern_rules` hit the right targets (`find_paragraphs --regex` validates coverage before apply)
- **Implicit-keep is a FAILURE signal** — non-empty fingerprints not routed by any rule. Add `bulk_rules` / `assignments` until implicit-keep is empty, OR until the remaining ones are intentional non-content the user wants preserved.

### 4. Apply

Output is a fresh docx; the original is never modified.

## Asking the user

Default first. Don't ask unless one of the cases below applies:
- typed sentinel mid-prose vs. heading
- bold paragraph as sub-heading vs. emphasis
- missing source content for template slots
- unsupported structures (footnotes, page-number cites)
- font / spacing prompt without explicit scope (theme layer vs `Normal` cascade vs per-role `styles[]` — see standardize.md "Three layers for setting fonts")

## Out of scope

Surface to the user when blocked by: layout-**table** restructuring, TOC body content, footnotes / comments / headers / footers content. Paragraphs *inside* layout-table cells ARE indexed and fully editable; only the table holding them is off-limits.

**Cross-references are mandatory for any cite to a numbered target.** Figures, tables, sections, chapters, equations, theorem numbers, reference-list entries — any time body text mentions an auto-numbered target, the cite goes through `edits[].text[]`'s `InlineRef` node, never as typed counter. Literal counters silently desync when a figure is inserted or a reference reordered; REF fields stay correct. See [`references/cross-references.md`](references/cross-references.md) for the InlineRef schema, target requirements, and display options. Page-number cites remain out of scope — surface to the user.

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file> [--paras=A..B\|none] [--include-unused]` | First call on any task. Metadata, page setup (mm), theme, style defs, numbering schemes (clustered by pattern), visual style statistics, **direct-format summary per fingerprint** (which pPr / run-level rPr attributes content chrome carries — drives the "don't redeclare" rule), document skeleton. `--paras=A..B` slices the skeleton to a range (re-consult after first survey); `--paras=none` drops the skeleton (when only the style / numbering dictionary is needed). Unused styles (usage=0) are hidden by default; `--include-unused` shows them. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump. Use for paragraphs with mixed run-level formatting OR **form-fill segments** (label + underscore-blank pattern) — see how the blank is structured before deciding the edit shape. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph. First choice for figure-caption / table-caption / first-after-heading classification. |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | What role a fingerprint plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` and their `basedOn` chain. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections. |
| `inspect_table` | `node scripts/inspect_table.js <file>` | Top-level tables with cell text snippets at `[row,col]` and the paragraph-index span each cell occupies (`paras: 60–89`). Use before composing a `cell` locator or a `range` that touches table content — the para spans show where row boundaries fall so a `range` doesn't accidentally cross cells. |
| `inspect_blockers` | `node scripts/inspect_blockers.js <file>` | Paragraphs `apply`'s edit phase will refuse — tracked changes, complex fields, SDT controls. |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. Validate `pattern_rules` regex coverage before applying. |
| `find_text` | `node scripts/find_text.js <file> <pattern> [--regex] [--paragraph N \| --range A-B] [--limit N] [--context N]` | Character-level locator. Returns paragraph index, run index (when in a direct-child run), char offset, length, and a context preview with the match bracketed. Annotates structural region (cross-run, in tracked change, in field result, in content control, in hyperlink) when applicable. Read-only — pair with `inspect_runs` for run rPr details, with `set-run` to replace at `runIndex`, or use standalone for browsing / counting / coverage validation. |
| `validate` | `node scripts/validate.js <file>` | Schema-aware OOXML check. `apply` runs this automatically; standalone for spot-checking arbitrary .docx files. |
| `apply --dry-run` | `node scripts/apply.js --dry-run <config.json>` | Iterate on a config without writing. **Use between every config edit.** |
| `apply` | `node scripts/apply.js <config.json>` | Unified writer. Default for any write task. |

## Cross-command invariants

- Original file is never modified; every applying CLI writes a fresh copy + validates before keeping. Validation failure → discard, surface to user — don't silently retry.
- Section properties (page size, margins, headers, footers, columns) never modified.
- Paragraph indexing is 1-based, matching `#NNN` in skeleton. Layout-table paragraphs are indexed; data / form-table paragraphs aren't (reachable via cell locator on edit path).
- All `edits[]` locators (paragraph indices, range, cell coordinates) resolve against the **pre-edits** document state — the `#NNN` you see in `overview` is what locators reference, regardless of whether earlier ops in the same `edits[]` array shift indices. Resolved Element refs survive subsequent mutations.
- Paths resolve against CWD; use absolute paths if you may have changed directories.
- Restyle: run-level direct formatting uniform across all runs gets stripped; per-run differences preserved as intentional inline emphasis.
- Field codes (`STYLEREF` / `TOC` / `REF` / `DATE` / …) preserved as-is; not editable inside.
- TOC content is not regenerated; user must right-click → "Update Field" in Word after opening.
- Edit blockers: the edit phase refuses paragraphs inside existing tracked changes / complex fields / SDT controls. Run `inspect_blockers` first.
- Style-name preflight: before any mutation, the engine catches `<w:name>` collisions (including en/zh-CN locale aliases) and aborts with a fix hint. In `--dry-run` collisions are listed as warnings so the agent can adjust all at once before committing.

Block-level config details: [`references/standardize.md`](references/standardize.md) (styles / numbering / pattern_rules / bulk_rules / assignments / exclude), [`references/edit.md`](references/edit.md) (locators / ops / MDF / track-changes), [`references/config-schema.md`](references/config-schema.md) (full field reference).
