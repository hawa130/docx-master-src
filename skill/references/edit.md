# `edits` block (used inside `apply`)

`edits` is a sub-block of `apply`'s config — it covers surgical content + format changes at specific locations: replace / insert / delete paragraphs, swap a table cell, embed an image, restyle a paragraph or range. Optional Word tracked-changes mode.

*Illustrative phrasings: "把第 3 段改成 ...", "在 X 章后面插一段", "这个表格第 2 行换成 ...", "删掉那段开题摘要", "给第 5 段加粗", "在结论后插张图". For role-based whole-doc reshape, drive the change via `pattern_rules` / `bulk_rules` in the same config; for surgical location-based work, use this `edits` block.*

There's no separate `apply_edits` CLI — content edits run through `apply` with `edits[]`. Inserts that introduce structural roles (prose body, list items, sub-headings) need the matching style installed in `styles[]` so Blocks bind via `styleId` rather than ad-hoc per-op `paraFormat` / `runFormat`. `numbering` and `pattern_rules` slot in alongside when the task spans new structure + chrome retags. **One config, one call.**

## Reconnaissance first

Always inspect before composing edits.

- `inspect_table` — top-level tables with `[row,col]` cell snippets (before composing a `cell` locator).
- `inspect_blockers` — paragraphs `apply` will refuse to touch in the edit phase (existing tracked changes / fields / SDT controls).
- `overview` / `find_paragraphs` / `inspect_range` / `inspect_neighbors` — same inspect tools the standardize-shape config uses.

## Config shape (within an `apply` config)

```json
{
  "source": "input.docx",
  "output": "output.docx",
  // ... styles[], numbering, pattern_rules, etc. as needed ...
  "edits": [
    { "op": "...", "at": { "type": "..." } /* op-specific fields */ }
  ],
  "trackChanges": false
}
```

`edits[]` runs in array order during `apply`'s pipeline (after style/numbering install, before pattern_rules cleanup). Failures abort atomically; the original file is never modified.

### Locators (`at`)

Index conventions vary by field: paragraph-level (matching `#NNN` in `overview`) is **1-based**; OOXML-internal positions (table cell coords, run indices) are **0-based**. Each row below states its own convention to avoid the mixup.

| `type` | Selects |
|---|---|
| `paragraph` | The Nth indexed paragraph (`{ "type": "paragraph", "index": N }`). **1-based**, matches `#NNN` in `overview`. |
| `range` | `{ ..., "from": A, "to": B }`, **1-based** inclusive (same indexing as `paragraph`). Endpoints must share a container (body, or one specific layout-table cell). |
| `cell` | `{ ..., "table": T, "row": R, "col": C }`, **0-based** (T/R/C). Only way to reach data/form-table cell paragraphs (those are unindexed). |
| `heading` | `{ ..., "text": "...", "level"?: L }`. First paragraph whose rendered text matches and whose outline level is L. Disambiguate with `find_paragraphs` if multiple match, then switch to `paragraph` index. |
| `whole-body` | Every body paragraph. Pairs naturally with `format`; rarely with `replace`. |
| `run` | A specific `<w:r>` inside a paragraph. `{ "type": "run", "paragraph": N, "blank"?: K, "runIndex"?: M }`. `paragraph` is 1-based; `blank` and `runIndex` are **0-based**. With `blank: K`, targets the Kth (0-indexed) run whose text is whitespace-only and rPr carries `<w:u/>` (form-fill placeholder); with `runIndex: M`, targets the Mth (0-indexed) run. If neither `blank` nor `runIndex` is given, defaults to `blank: 0`. Pair only with `set-run`. |

### Ops

- **`replace`** — `{ ..., "with": [Block, ...] }`. Removes targets, inserts fragment in their place.
- **`insert-before` / `insert-after`** — `{ ..., "content": [Block, ...] }`. Inserts fragment immediately before / after the target.
- **`delete`** — `{ ... }`. Removes the targeted paragraph(s).
- **`format`** — `{ ..., "styleId"?, "runFormat"?, "paraFormat"? }`. Mutates existing paragraphs without changing their content. At least one of styleId / runFormat / paraFormat required.
- **`set-run`** — `{ "at": <run-locator>, "with": "value text", "format"?: { ... } }`. Replaces the targeted run's text while preserving its rPr (font / underline / size carry through). Use for filling form-fill placeholder runs without manually reconstructing label + value runs. `format` accepts the same fields as `runFormat` (bold / italic / underline / strike / color / fontLatin / fontCJK / size); absent, the run's existing rPr stays verbatim.

### Match-destination formatting (default)

`replace` / `insert-before` / `insert-after` make new `paragraph` blocks inherit the **anchor** paragraph's `<w:pPr>` — same semantics as Word's "Match Destination Formatting" paste mode. Anchor: first replaced (replace), first target (insert-before), last target (insert-after). Inheritance is additive at pPr-child granularity — explicit `styleId` / `paraFormat` on the Block always wins. Set `"styleId": "Normal"` to opt out. `image` / `page-break` / `horizontal-rule` blocks don't inherit.

**Bold-pMark trap**: a label paragraph (heading-style) often has bold paragraph-mark rPr; the empty placeholder row beneath it inherits that bold. When you `replace` or `insert-after` against either, MDF would propagate the bold into your new paragraph's pPr-mark — and Word's style cascade can't undo it (it's not run rPr, not paragraph rPr). The engine handles the common case: when the new paragraph carries an explicit `styleId`, the anchor's pPr-mark rPr is skipped on inheritance and the style cascade governs. If you skip styleId (rare; Block uses MDF fallback), explicit `runFormat: { bold: false }` on each run is the override (writes `<w:b w:val="0"/>`).

### Blocks (in `with` / `content`)

```json
{ "type": "paragraph", "text": "...", "styleId"?, "paraFormat"?, "runFormat"?, "numbering"? }
{ "type": "image", "src": "path", "widthPt": N, "heightPt": N, "alt"? }
{ "type": "page-break" }
{ "type": "horizontal-rule" }
```

`text` is either a plain string (single run, no inline formatting) or an array of `{ text, format }` for mixed run-level formatting. Image dimensions are required.

**Express structure semantically.** Hierarchy and list shape bind via `styleId` and `numbering` — not by typing markers in `text`. Two paths to numbering:
- **styleId-bound** (preferred): if the styleId you set is bound to a numbering scheme via `numbering[].levels[].styleId`, the binding handles auto-numbering automatically — don't supply a `numbering` field on the block. Use for headings (`Heading1..N`) and list-bound styles (`ListNumber` / `ListBullet`).
- **ad-hoc** `numbering: { numId, level }`: for one-off paragraph-level numbering not tied to a style. Rare; used when you want a paragraph numbered without committing the style to a scheme.

If the styleId or numId you reference doesn't exist in the doc, add `styles[]` / `numbering` to the same `apply` config so they get installed before `edits[]` runs (see SKILL.md Target state).

### Quote handling

`text` is emitted verbatim. Default to smart quotes in prose (Chinese `“…”` / `「…」`, English `“…”` / `‘…’`). Use ASCII `"` `'` only inside literal tokens (code, URLs, identifiers, shell commands). Smart quotes also bypass the JSON `\"` escape footgun.

### Format fields

`runFormat`: `bold` / `italic` / `underline` / `strike` (boolean, tri-state — `false` emits explicit off-toggle to override inherited true), `color` (`"RRGGBB"`), `fontLatin` / `fontCJK`, `size` (pt).

`paraFormat`: `alignment` (`"left" | "center" | "right" | "both"`), `spaceBefore` / `spaceAfter` (pt), `lineSpacing` + `lineRule` (same convention as a style definition), `firstLineIndent` / `hangingIndent` / `indentLeft` / `indentRight` (`"Nchar"` / `"Npt"` / number), `outlineLevel` (0–9).

Mirrors the style schema — what you'd put inside a `styles[]` entry, you put as direct paragraph formatting here.

## Track-changes mode

`"trackChanges": true` emits edits as Word revision markup: text changes via `<w:ins>` / `<w:del>`, format changes via `<w:rPrChange>` / `<w:pPrChange>` snapshots. The user accepts / rejects in Word's Review tab. Author info is intentionally blank — this is a review affordance, not an authorship claim. Paragraphs that *already* contain tracked changes are blocked; ask the user to accept / reject existing revisions first.

## Cell-fill strategy

Form cells often hold a label paragraph + several empty placeholder rows (the form designer pre-allocates blanks for handwriting). Two strategies:

- **`replace` a range** covering the empty placeholder rows: drops the placeholders, only the new content remains. Cleaner cell, no trailing blank rows.
- **`insert-after` the label**: leaves the empty placeholder rows below the inserted content. Word auto-grows the cell; visually fine for most forms but trailing blank rows are visible.

Default to `replace` for content-bearing fills; `insert-after` for non-destructive additions where the original layout is meant to be preserved.

For multi-paragraph content, pass multiple Blocks in one op rather than chaining many `insert-after`s — order is preserved and no stale-index issue.

## Edge cases

- **Cross-container range**: `range` cannot span body ↔ table-cell. Split into two edits.
- **Heading locator ambiguity**: returns first match. Use `find_paragraphs` to disambiguate, then refer by `paragraph` index.
- **Stale element**: if op A removes paragraphs and op B targets one of them, op B fails. Reorder so deletes / replaces happen last, or split into separate runs.
- **Empty body**: `whole-body` on a doc with zero paragraphs only accepts `insert-*` (appends before the trailing `<w:sectPr>`).

## Compose with other shapes

- **Messy template + content**: combine `styles[]` / `numbering` / `pattern_rules` (standardize-shape) with `edits[]` in one `apply` config. The engine installs structure first, then `edits[]` references the just-installed styleIds. Match-Destination-Formatting on a dirty template otherwise propagates the mess.
- **Whole-doc role-based reshape**: drop `edits[]` and use the standardize-shape blocks (see [standardize.md](standardize.md)).
- **Read-only conformance check**: `audit`. `apply` always writes when not in dry-run.
- **Single-style change covering all paragraphs of role X**: `restyle` (or standardize-shape `pattern_rules`) — narrower than per-locator edits.
