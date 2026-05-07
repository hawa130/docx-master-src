# Command: `edit`

Surgical content + format changes at specific locations on an existing docx — replace / insert / delete paragraphs, swap a table cell, embed an image, restyle a paragraph or range, optionally as Word tracked changes.

*Illustrative phrasings: "把第 3 段改成 ...", "在 X 章后面插一段", "这个表格第 2 行换成 ...", "删掉那段开题摘要", "给第 5 段加粗", "在结论后插张图". Same words can land in `standardize` when the user wants role-based whole-doc reshape — ask one focused question if intent is genuinely ambiguous.*

## Reconnaissance first

Always inspect before composing edits.

- `inspect_table` — top-level tables with `[row,col]` cell snippets (before composing a `cell` locator)
- `inspect_blockers` — paragraphs `apply_edits` will refuse (existing tracked changes / fields / SDT controls)
- `overview` / `find_paragraphs` / `inspect_range` / `inspect_neighbors` — same as the standardize path

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

#### Match-destination formatting (default)

`replace` / `insert-before` / `insert-after` make new `paragraph` blocks inherit the **anchor** paragraph's `<w:pPr>` — same semantics as Word's "Match Destination Formatting" paste mode. Anchor: first replaced (replace), first target (insert-before), last target (insert-after). Inheritance is additive at pPr-child granularity — explicit `styleId` / `format` on the Block always wins. Set `"styleId": "Normal"` to opt out. `image` / `page-break` / `horizontal-rule` blocks don't inherit.

### Blocks (in `with` / `content`)

```json
{ "type": "paragraph", "text": "...", "styleId"?, "format"?, "runFormat"?, "numbering"? }
{ "type": "image", "src": "path", "widthPt": N, "heightPt": N, "alt"? }
{ "type": "page-break" }
{ "type": "horizontal-rule" }
```

`text` is either a plain string (single run, no inline formatting) or an array of `{ text, format }` for mixed run-level formatting. Image dimensions are required — the tool does not infer from the file.

#### Quote handling

`text` is emitted verbatim. Default to smart quotes in prose (Chinese `"…"` / `「…」`, English `"…"` / `'…'`). Use ASCII `"` `'` only inside literal tokens (code, URLs, identifiers, shell commands). Smart quotes also bypass the JSON `\"` escape footgun.

### Format fields

`runFormat`: `bold` / `italic` / `underline` / `strike` (boolean), `color` (`"RRGGBB"`), `fontLatin` / `fontCJK`, `size` (pt).

`paraFormat`: `alignment` (`"left" | "center" | "right" | "both"`), `spaceBefore` / `spaceAfter` (pt), `lineSpacing` + `lineRule` (same convention as `apply_styles`), `firstLineIndent` / `hangingIndent` / `indentLeft` / `indentRight` (`"Nchar"` / `"Npt"` / number), `outlineLevel` (0–9).

Mirrors the standardize style schema — what you'd put inside a style definition there, you put as direct paragraph formatting here.

## Track-changes mode

`"trackChanges": true` emits edits as Word revision markup: text changes via `<w:ins>` / `<w:del>`, format changes via `<w:rPrChange>` / `<w:pPrChange>` snapshots. The user accepts / rejects in Word's Review tab. Author info is intentionally blank — this is a review affordance, not an authorship claim. Paragraphs that *already* contain tracked changes are blocked; ask the user to accept / reject existing revisions first.

## Filling a template

Form templates pair a **label paragraph** (heading-ish style) with **empty body paragraph(s)** styled for prose. Fill the empty body slot — `replace` it and let inheritance pick up the body style. Replacing the label inherits the label's heading style onto your prose, which renders wrong. When a template has no empty slot after a label, `insert-after` the label with an explicit `styleId` from the doc's body style (find via `inspect_style_def` or `overview`).

## Edge cases

- **Cross-container range**: `range` cannot span body ↔ table-cell. Split into two edits.
- **Heading locator ambiguity**: returns first match. Use `find_paragraphs` to disambiguate, then refer by `paragraph` index.
- **Stale element**: if op A removes paragraphs and op B targets one of them, op B fails. Reorder so deletes / replaces happen last, or split into separate runs.
- **Empty body**: `whole-body` on a doc with zero paragraphs only accepts `insert-*` (which appends before the trailing `<w:sectPr>`).

## When to use a different command

- Whole-doc role-based reshape ("classify body paragraphs as BodyText, normalize heading numbering"): `standardize`. `edit` does not classify by role or fingerprint.
- Read-only conformance check: `audit`. `edit` always writes.
- Single-style change without locator inspection: `standardize`'s targeted-edit path. `edit` requires explicit locators.
