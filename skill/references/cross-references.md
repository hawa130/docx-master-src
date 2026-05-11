# Cross-references (REF fields)

## When to use this

**Any time a body paragraph mentions a numbered target — figures, tables, sections, chapters, theorems, equations, reference list entries — use a cross-reference, not literal text.**

Concretely: never write `"如图 1 所示"` or `"见 [3]"` as plain runs in `edits[]`. The moment a figure gets inserted or a reference reordered, every literal cite silently desyncs. Use `InlineRef` so Word resolves the cite at render time.

The same rule applies to:
- "See Section 3.2" → cross-ref to the Heading2 paragraph
- "Chapter 5 discusses" → cross-ref to the Heading1
- "Equation (4)" → cross-ref to the auto-numbered equation caption
- "Table 2-1 shows" → cross-ref to the Caption

Page-number citations ("on page 12") are still out of scope (PAGEREF, not REF) — surface to user.

## Schema

`InlineRef` is an inline node alongside `InlineRun` inside a paragraph's `text` array:

```jsonc
{
  "type": "paragraph",
  "text": [
    { "text": "结果如" },
    { "refTo": { "type": "paragraph", "index": 42 }, "display": "label" },
    { "text": " 所示。" }
  ]
}
```

Fields:
- `refTo` — locator. Two forms:
  - `{ "type": "paragraph", "index": N }` — pre-edit 1-based paragraph index, same as every other `edits[]` locator. Target must exist in the source document.
  - `{ "type": "anchor", "name": "fig-arch" }` — named bookmark. Resolves against (a) anchors declared on `ParagraphBlock.anchor` earlier in this same `edits[]` array, or (b) bookmarks already in the source document (only those wrapping a single paragraph). Name must match `^[A-Za-z_][A-Za-z0-9_-]{0,39}$`.
- `display` — what Word shows. Each maps to a specific REF switch (see "Display option behavior" below):
  - `"label"` (default) — full numbered paragraph text resolved from `lvlText`, e.g. `图 1` / `1.2` / `[3]` / `第二章`. Word switch: `\n \h`.
  - `"number"` — relative-context paragraph number via `\r \h`. **Reliable only for single-level numbering schemes** (figure captions, reference lists `[%1]`), where it equals the level's counter — `1` / `2` / `3`. For multi-level schemes use `"label"` and rely on the lvlText to format what you want; `\r`'s relative-context semantics get hierarchy-dependent.
  - `"full"` — the target paragraph's body text content (caption title without the auto-numbered prefix, since the prefix is rendered from numbering and lives outside the bookmark's text). Word switch: `\h`. **Does not require an auto-numbered target** — works on any paragraph.
- `format` — optional `RunFormat` (color, italic, size, …) applied to the rendered text

## Target requirements

For `display: "label"` and `display: "number"`: target must be **bound to a numbering scheme in this apply run** — either pre-existing in the source's `pStyle → numId` binding, or freshly bound by this run's `numbering[]` config. Word's `\n` / `\r` switches render from the numbering binding, so an unbound target has nothing to display.

For `display: "full"`: any paragraph works. The bookmark resolves to the target's text content directly, no numbering required.

When `label` / `number` hits an unbound target, the engine refuses at apply time:

```
edits[5] (insert-after): InlineRef: target paragraph #42 is not bound
to a numbering scheme. display="label" requires an auto-numbered target …
Either bind the target's pStyle to a numbering[] level, or set display:
"full" to use the paragraph's body text instead.
```

## Named anchors — ref new paragraphs in the same apply

When an `edits[]` insert creates a paragraph that later refs will cite, give it an `anchor` so the refs can address it by name:

```jsonc
{
  "edits": [
    {
      "op": "insert-after",
      "at": { "type": "paragraph", "index": 50 },
      "content": [
        { "type": "image", "src": "diagrams/arch.png", "widthPt": 360, "heightPt": 240 },
        {
          "type": "paragraph",
          "styleId": "FigureCaption",
          "anchor": "fig-architecture",
          "text": "系统总体架构"
        }
      ]
    },
    {
      "op": "insert-after",
      "at": { "type": "paragraph", "index": 60 },
      "content": [
        {
          "type": "paragraph",
          "text": [
            { "text": "本文系统的架构如" },
            { "refTo": { "type": "anchor", "name": "fig-architecture" }, "display": "label" },
            { "text": " 所示。" }
          ]
        }
      ]
    }
  ]
}
```

Anchors are processed in `edits[]` array order: a ref can address any anchor declared earlier in the array, but not a later one (forward refs not supported in v1 — Reorder edits if needed). Anchor names collide with source bookmarks (same flat namespace) — picking a name that already exists in the source throws at apply time.

`anchor` can also be used on an inserted paragraph that no current `InlineRef` cites — the bookmark stays in the output for future apply runs to ref by name.

## Word's first-open behavior

The skill writes:
1. A `<w:bookmarkStart>/<w:bookmarkEnd>` pair wrapping the target paragraph
2. A REF field in the body run sequence: `REF _RefXXXXXX \n \h` (switches per `display`)
3. The numbering counter is simulated at apply time and the **placeholder text** between `<w:fldChar fldCharType="separate"/>` and `<w:fldChar fldCharType="end"/>` is set to the predicted value (e.g. `图 1`). This means the doc reads correctly even if the user opens it without updating fields.
4. `<w:updateFields w:val="true"/>` is set in `word/settings.xml`. On next open, Word automatically updates every field — the user sees a brief prompt and after confirming, all REFs resolve to their live target text.
5. **When `InlineRef.format` is set** (and produces non-empty rPr): the same `<w:rPr>` is replicated onto every one of the 5 field runs (begin / instrText / separate / placeholder / end), and `\* MERGEFORMAT` is appended to the field code. Together these tell Word to preserve the formatting when it rewrites the result run on field update. Without both, Word reads rPr from the surrounding (empty) field runs after F9 and the format silently disappears.

If the user dismisses the prompt without updating, the placeholder text from step 3 is still visible and correct for the as-emitted state. Updates take effect on the next manual `Ctrl+A` → `F9`.

### Verifying format-bearing refs

Inspecting the apply-time XML alone is **not enough** for refs that carry a `format`. Word rewrites the field's result run on every update; if rPr replication or `\* MERGEFORMAT` were missing, the emit-time XML would still look correct, validation would still pass, and the format would only disappear after the first update.

Round-trip in Word to verify:
1. Open the apply output in Word
2. Accept the "update fields" prompt
3. Save
4. Unzip and grep `word/document.xml` for the relevant `instrText` — confirm the surrounding result run still carries the expected `<w:rPr>` after the update

This is the only reliable check for format preservation. Schema validation and emit-time inspection both pass the broken case; only a real Word round-trip catches the regression.

## Display option behavior in detail

OOXML REF switches map directly to display choices:

| display    | switches | Word renders                                                          |
|------------|----------|-----------------------------------------------------------------------|
| `"label"`  | `\n \h`  | Full numbered paragraph text resolved from lvlText: `图 1` / `1.2.3` / `第二章` |
| `"number"` | `\r \h`  | Paragraph number in **relative context**. For single-level schemes this equals the counter (`1` / `3`). For multi-level schemes, returns the position within the level, which is not equivalent to "just the leaf digit" — prefer `"label"` for multi-level. |
| `"full"`   | `\h`     | The bookmark's body text — the caption title or paragraph text without auto-num |

`\h` is included on every form so the rendered text is clickable in Word, matching what `Insert → Cross-reference` produces.

## Examples

### Figure caption

```jsonc
{ "type": "paragraph", "text": [
  { "text": "如" },
  { "refTo": { "type": "paragraph", "index": 50 }, "display": "label" },
  { "text": " 所示，系统采用三层架构。" }
]}
// Renders: 如图 1 所示，系统采用三层架构。
```

### Reference list cite (numeric only)

```jsonc
{ "type": "paragraph", "text": [
  { "text": "本文方法借鉴了 [" },
  { "refTo": { "type": "paragraph", "index": 180 }, "display": "number" },
  { "text": "] 中的思路。" }
]}
// When ref target uses lvlText="[%1]" and bookmark count is 3:
// Renders: 本文方法借鉴了 [3] 中的思路。
//
// Note: `display: "number"` returns just "3" — wrapping with literal "["
// and "]" lets the agent control the bracket form independently of the
// target's lvlText. For "[1]" with brackets baked into lvlText use
// `display: "label"` instead.
```

### Section / chapter ref

```jsonc
{ "type": "paragraph", "text": [
  { "text": "详细方法见 " },
  { "refTo": { "type": "paragraph", "index": 30 }, "display": "label" },
  { "text": "。" }
]}
// Heading2 bound to lvlText="%1.%2": "详细方法见 3.2。"
```

## Limitations

- **No PAGEREF**: page-number citations ("on page 12") aren't supported — they require a different field type and depend on Word's pagination layer.
- **No HYPERLINK to external resources**: only intra-document REF fields.
- **No `\p` (above/below) switch**: REF targeting relative position isn't supported.
- **No forward anchor refs**: a ref can only address an anchor declared earlier in the same `edits[]` array. Reorder edits if needed.
- **Source bookmarks must wrap a single paragraph**: source bookmarks spanning multiple paragraphs (or sitting at body level outside any `<w:p>`) aren't resolvable via `refTo: { type: "anchor" }` — there's no paragraph-level target to surface for REF rendering. Such bookmarks are silently skipped at allocator construction.
- **Custom `lvlRestart` not honored in placeholder**: the simulator uses Word's default reset-on-higher-level behavior. Custom restart points produce correct text after `F9` (Word re-renders) but the placeholder may diverge. Cosmetic only.
