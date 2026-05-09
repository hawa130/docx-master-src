# Command: `edit`

Surgical content + format changes at specific locations on an existing docx — replace / insert / delete paragraphs, swap a table cell, embed an image, restyle a paragraph or range. Optional Word tracked-changes mode.

*Illustrative phrasings: "把第 3 段改成 ...", "在 X 章后面插一段", "这个表格第 2 行换成 ...", "删掉那段开题摘要", "给第 5 段加粗", "在结论后插张图". Same words can land in `standardize` when the user wants role-based whole-doc reshape — ask one focused question if intent is genuinely ambiguous.*

## Reconnaissance first

Always inspect before composing edits.

- `inspect_table` — top-level tables with `[row,col]` cell snippets (before composing a `cell` locator).
- `inspect_blockers` — paragraphs `apply_edits` will refuse (existing tracked changes / fields / SDT controls).
- `overview` / `find_paragraphs` / `inspect_range` / `inspect_neighbors` — same as the standardize path.

## Config shape

```json
{
  "source": "input.docx",
  "output": "output.docx",
  "edits": [
    { "op": "...", "at": { "type": "..." } /* op-specific fields */ }
  ],
  "trackChanges": false
}
```

`edits[]` runs in array order. Failures abort atomically; the original file is never modified.

### Locators (`at`)

| `type` | Selects |
|---|---|
| `paragraph` | The Nth indexed paragraph (`{ "type": "paragraph", "index": N }`). 1-based, matches `#NNN` in `overview`. |
| `range` | `{ ..., "from": A, "to": B }`, inclusive. Endpoints must share a container (body, or one specific layout-table cell). |
| `cell` | `{ ..., "table": T, "row": R, "col": C }`, 0-based. Only way to reach data/form-table cell paragraphs (those are unindexed). |
| `heading` | `{ ..., "text": "...", "level"?: L }`. First paragraph whose rendered text matches and whose outline level is L. Disambiguate with `find_paragraphs` if multiple match, then switch to `paragraph` index. |
| `whole-body` | Every body paragraph. Pairs naturally with `format`; rarely with `replace`. |

### Ops

- **`replace`** — `{ ..., "with": [Block, ...] }`. Removes targets, inserts fragment in their place.
- **`insert-before` / `insert-after`** — `{ ..., "content": [Block, ...] }`. Inserts fragment immediately before / after the target.
- **`delete`** — `{ ... }`. Removes the targeted paragraph(s).
- **`format`** — `{ ..., "styleId"?, "runFormat"?, "paraFormat"? }`. Mutates existing paragraphs without changing their content. At least one of styleId / runFormat / paraFormat required.

### Match-destination formatting (default)

`replace` / `insert-before` / `insert-after` make new `paragraph` blocks inherit the **anchor** paragraph's `<w:pPr>` — same semantics as Word's "Match Destination Formatting" paste mode. Anchor: first replaced (replace), first target (insert-before), last target (insert-after). Inheritance is additive at pPr-child granularity — explicit `styleId` / `paraFormat` on the Block always wins. Set `"styleId": "Normal"` to opt out. `image` / `page-break` / `horizontal-rule` blocks don't inherit.

**Bold-pMark trap**: a label paragraph (heading-style) often has bold paragraph-mark rPr; the empty placeholder row beneath it inherits that bold. When you `replace` or `insert-after` against either, the new content inherits bold via Match-Destination-Formatting and renders wrong for prose. **Inspect the anchor's rPr first.** If it carries unwanted bold, override per-Block with `runFormat: { bold: false }` (the emitter writes `<w:b w:val="0"/>` for explicit-off, distinct from omitting the field). For systemic template defects (slot-rows uniformly bold-styled), the cleaner fix is a `standardize` pass that resets the slot's pStyle.

### Blocks (in `with` / `content`)

```json
{ "type": "paragraph", "text": "...", "styleId"?, "paraFormat"?, "runFormat"?, "numbering"? }
{ "type": "image", "src": "path", "widthPt": N, "heightPt": N, "alt"? }
{ "type": "page-break" }
{ "type": "horizontal-rule" }
```

`text` is either a plain string (single run, no inline formatting) or an array of `{ text, format }` for mixed run-level formatting. Image dimensions are required.

**Express structure semantically.** Hierarchy and list shape bind via `styleId` and `numbering` — not by typing markers in `text`. List items use `numbering: { numId, level }`; sub-headings use `styleId: "Heading3"`. If the styleId or numId you reference doesn't exist in the doc, `standardize` first to install it (see SKILL.md Target state).

### Quote handling

`text` is emitted verbatim. Default to smart quotes in prose (Chinese `“…”` / `「…」`, English `“…”` / `‘…’`). Use ASCII `"` `'` only inside literal tokens (code, URLs, identifiers, shell commands). Smart quotes also bypass the JSON `\"` escape footgun.

### Format fields

`runFormat`: `bold` / `italic` / `underline` / `strike` (boolean, tri-state — `false` emits explicit off-toggle to override inherited true), `color` (`"RRGGBB"`), `fontLatin` / `fontCJK`, `size` (pt).

`paraFormat`: `alignment` (`"left" | "center" | "right" | "both"`), `spaceBefore` / `spaceAfter` (pt), `lineSpacing` + `lineRule` (same convention as `apply_styles`), `firstLineIndent` / `hangingIndent` / `indentLeft` / `indentRight` (`"Nchar"` / `"Npt"` / number), `outlineLevel` (0–9).

Mirrors the standardize style schema — what you'd put inside a style definition there, you put as direct paragraph formatting here.

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

## Compose with other commands

- **Messy template + content**: `standardize` first to install / fix the style system, then `edit` to insert content. Match-Destination-Formatting on a dirty template propagates the mess.
- **Whole-doc role-based reshape**: `standardize`, not `edit`.
- **Read-only conformance check**: `audit`. `edit` always writes.
- **Single-style change covering all paragraphs of role X**: `standardize`'s targeted-restyle path — narrower than `edit`'s per-locator approach.
