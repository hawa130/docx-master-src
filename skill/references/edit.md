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

`replace` / `insert-before` / `insert-after` make new `paragraph` blocks inherit the **anchor** paragraph's `<w:pPr>` — same semantics as Word's "Match Destination Formatting" paste mode. Anchor: first replaced (replace), first target (insert-before), last target (insert-after). Inheritance is additive at pPr-child granularity — explicit `styleId` / `paraFormat` on the Block always wins. Set `"styleId": "Normal"` to opt out. `image` / `page-break` / `horizontal-rule` blocks don't inherit.

### Blocks (in `with` / `content`)

```json
{ "type": "paragraph", "text": "...", "styleId"?, "paraFormat"?, "runFormat"?, "numbering"? }
{ "type": "image", "src": "path", "widthPt": N, "heightPt": N, "alt"? }
{ "type": "page-break" }
{ "type": "horizontal-rule" }
```

`text` is either a plain string (single run, no inline formatting) or an array of `{ text, format }` for mixed run-level formatting. Image dimensions are required — the tool does not infer from the file.

#### Express structure semantically, not in text

If content has hierarchy or list shape, bind it via `styleId` and `numbering` — never type the markers in `text`:

- List items → `{ "type": "paragraph", "numbering": { "numId": "5", "level": 0 }, "text": "..." }`. **Not** `"text": "1. ..."`.
- Sub-headings → `{ "type": "paragraph", "styleId": "Heading3", "text": "..." }`. **Not** `"text": "（1）..."` with a bold runFormat.

If the styleId or numId you need doesn't exist in the document, the **right move is `standardize` to install it**, then come back to `edit`. Falling back to typed prefixes is a footgun: the result fails to typeset as a real list or heading; Word loses outline navigation, TOC binding, and accept-changes granularity; the document's logical structure no longer matches its visual structure.

When in doubt — should I install Heading3, or fold sub-headings into bold? — **ask the user**, per SKILL.md "Ask, don't decide". Don't decide silently.

#### Form chrome is not a hierarchy strategy

Many templates carry typed prefixes — "一、论文概况", "（一）选题意义" — as scaffolding for the form itself. These are **template chrome**. They tell you nothing about how to express your content's hierarchy.

The trap: agent sees "the template uses typed prefixes" → agent reasons "so my content's sub-headings should also use typed prefixes" → agent types `1. 理论意义` / `（1）模型压缩方向` as text. **This is wrong every time.** Template chrome and content hierarchy are independent decisions:

- Template chrome (outer section labels that came with the template): leave as text. Don't try to convert them to Heading styles unless explicitly asked.
- Your content's hierarchy (sub-headings, lists, anything from the markdown): express semantically via `styleId` + `numbering`. If the template lacks the styles you need, `standardize` to install them.

When the choice between these strategies is genuinely user-discretionary — e.g., the user might want a fast form-style fill, or might want full semantic typesetting — **ask** (per SKILL.md "Ask, don't decide"). Do not infer the answer from the template's chrome convention.

#### Quote handling

`text` is emitted verbatim. Default to smart quotes in prose (Chinese `"…"` / `「…」`, English `"…"` / `'…'`). Use ASCII `"` `'` only inside literal tokens (code, URLs, identifiers, shell commands). Smart quotes also bypass the JSON `\"` escape footgun.

### Format fields

`runFormat`: `bold` / `italic` / `underline` / `strike` (boolean), `color` (`"RRGGBB"`), `fontLatin` / `fontCJK`, `size` (pt).

`paraFormat`: `alignment` (`"left" | "center" | "right" | "both"`), `spaceBefore` / `spaceAfter` (pt), `lineSpacing` + `lineRule` (same convention as `apply_styles`), `firstLineIndent` / `hangingIndent` / `indentLeft` / `indentRight` (`"Nchar"` / `"Npt"` / number), `outlineLevel` (0–9).

Mirrors the standardize style schema — what you'd put inside a style definition there, you put as direct paragraph formatting here.

## Track-changes mode

`"trackChanges": true` emits edits as Word revision markup: text changes via `<w:ins>` / `<w:del>`, format changes via `<w:rPrChange>` / `<w:pPrChange>` snapshots. The user accepts / rejects in Word's Review tab. Author info is intentionally blank — this is a review affordance, not an authorship claim. Paragraphs that *already* contain tracked changes are blocked; ask the user to accept / reject existing revisions first.

## Filling a template

Plan first — see SKILL.md Core Principle. The plan-survey-execute loop for fill tasks:

**Step 1: Survey the content** to be inserted. Hierarchy depth (how many heading levels)? Lists (ordered, bulleted, nested)? Inline emphasis? Tables / images / code? If the content is markdown, read its AST shape, not just the text.

**Step 2: Survey the template's expressiveness** via `overview`. Note:

- Defined styles: how many Heading levels? Body styles? List-bound styles (any pStyle with attached `<w:numPr>`)? Caption / Quote / Code styles?
- Defined numbering schemes: how many levels does each cover? Are any pre-bound to heading styles?
- Empty body slots inside cells: what's their pStyle, and is it actually styled for prose? Some templates leave bold or other surprising direct formatting on the empty rows.

**Step 3: Compare**. If the template covers what the content needs, proceed. If it doesn't, **stop and run `standardize` first** to install the missing pieces. Examples that should trigger standardize:

- Content has H3 / H4 but template defines only H1 / H2 → inject Heading3 / Heading4 styles + extend numbering to match.
- Content has bullet lists but template has only ordered numbering → install a bullet-bound style.
- Empty slot inherits unintended formatting (bold for prose, weird spacing) → either override at the source via standardize, or compensate per-Block — the trade-off is yours, ask the user when it matters.

If you face a genuine choice — flatten H4 to H3 vs install Heading4? install bullet style vs reuse ordered? — ask one focused question. Don't pick a default silently.

**Step 4: Execute edit ops**. With the style toolbox now complete:

- Find the empty body slot for prose content (`text: ""` with a body-style pStyle, often `Normal` / `Body Text` / `a3`).
- For list items, use `numbering: { numId, level }` on the Block — don't type `1.` / `（1）` in `text`.
- For sub-headings within content, use `styleId: "Heading3"` — don't approximate with bold + bigger font.
- Match-Destination-Formatting picks up the slot's pPr by default. `inspect_range` the slot first to confirm its formatting is what you want; override per-Block when it isn't.

Anti-pattern: replacing the **label paragraph** ("选题来源：") instead of the **empty body slot** that follows it. The label's heading style inherits onto your prose and renders wrong.

Inverse anti-pattern (just as common): `insert-after` on a label whose paragraph-mark `<w:rPr>` carries unwanted formatting — typically bold left over from the label's text. Match-Destination-Formatting picks up that pMark rPr and makes your inserted prose bold. Two ways to handle:

- Prefer `replace`-ing the empty body slot beneath the label (its formatting is meant for prose).
- If you must `insert-after` the label, override explicitly on each new Block: either `styleId: "Normal"` (clean opt-out of inheritance) or `runFormat: { bold: false }` (negates the inherited toggle — the emitter writes `<w:b w:val="0"/>` for explicit-off, distinct from omitting the field).

When the empty slot is itself styled wrong (a "body" placeholder that's actually bold), this is a template defect surfacing as inheritance — surface it to the user; the right fix is usually a `standardize` pass to clean up the template's defaults rather than working around it per-Block.

#### insert-after vs replace ranges in form cells

Form cells often hold a label paragraph + several empty placeholder rows (the form designer pre-allocates blank rows for handwriting). Two strategies:

- **`insert-after` the label**: appends content but leaves the empty placeholder rows below. Word auto-grows the cell; visually fine for most forms. Non-destructive; preserves the original layout.
- **`replace` a range covering the label + placeholders**: drops the placeholders, only your content remains. Tighter result; risks losing the label if you include it in the range.

Default to `insert-after` unless the user explicitly wants a clean cell (no remnants). For multi-paragraph content, pass multiple Blocks in the same op rather than chaining many `insert-after`s.

## Edge cases

- **Cross-container range**: `range` cannot span body ↔ table-cell. Split into two edits.
- **Heading locator ambiguity**: returns first match. Use `find_paragraphs` to disambiguate, then refer by `paragraph` index.
- **Stale element**: if op A removes paragraphs and op B targets one of them, op B fails. Reorder so deletes / replaces happen last, or split into separate runs.
- **Empty body**: `whole-body` on a doc with zero paragraphs only accepts `insert-*` (which appends before the trailing `<w:sectPr>`).

## Compose with other commands

- **Messy template + content**: run `standardize` first to install a clean style system, then `edit` to fill — Match-Destination-Formatting on a dirty template propagates the mess.
- **Whole-doc role-based reshape**: `standardize`, not `edit`. `edit` does not classify by role or fingerprint.
- **Read-only conformance check**: `audit`. `edit` always writes.
- **Single-style change without locator inspection**: `standardize`'s targeted-edit shape — narrower than `edit`'s explicit locators when the change is "all paragraphs of role X."
