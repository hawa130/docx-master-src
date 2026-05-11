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
- `refTo` — locator. **Pre-edit paragraph index** (1-based, same as every `edits[]` locator). The target must be a paragraph that exists in the source document.
- `display` — what Word shows. Each maps to a specific REF switch (see "Display option behavior" below):
  - `"label"` (default) — full numbered paragraph text resolved from `lvlText`, e.g. `图 1` / `1.2` / `[3]` / `第二章`. Word switch: `\n \h`.
  - `"number"` — relative-context paragraph number via `\r \h`. **Reliable only for single-level numbering schemes** (figure captions, reference lists `[%1]`), where it equals the level's counter — `1` / `2` / `3`. For multi-level schemes use `"label"` and rely on the lvlText to format what you want; `\r`'s relative-context semantics get hierarchy-dependent.
  - `"full"` — the target paragraph's body text content (caption title without the auto-numbered prefix, since the prefix is rendered from numbering and lives outside the bookmark's text). Word switch: `\h`.
- `format` — optional `RunFormat` (color, italic, size, …) applied to the rendered text

## Target requirements

The target paragraph must be **bound to a numbering scheme in this apply run** — either pre-existing in the source's `pStyle → numId` binding, or freshly bound by this run's `numbering[]` config.

If the target has no `<w:numPr>` after the apply pipeline's numbering pass, the engine refuses at apply time:

```
edits[5] (insert-after): InlineRef: target paragraph #42 is not bound
to a numbering scheme. Cross-references require an auto-numbered target —
bind the target's pStyle to a numbering[] level, or change the locator
to a numbered paragraph.
```

Fix by adding a `numbering[]` entry that binds the target's style, or by re-pointing the locator.

## Word's first-open behavior

The skill writes:
1. A `<w:bookmarkStart>/<w:bookmarkEnd>` pair wrapping the target paragraph
2. A REF field in the body run sequence: `REF _RefXXXXXX \n \h` (switches per `display`)
3. The numbering counter is simulated at apply time and the **placeholder text** between `<w:fldChar fldCharType="separate"/>` and `<w:fldChar fldCharType="end"/>` is set to the predicted value (e.g. `图 1`). This means the doc reads correctly even if the user opens it without updating fields.
4. `<w:updateFields w:val="true"/>` is set in `word/settings.xml`. On next open, Word automatically updates every field — the user sees a brief prompt and after confirming, all REFs resolve to their live target text.

If the user dismisses the prompt without updating, the placeholder text from step 3 is still visible and correct for the as-emitted state. Updates take effect on the next manual `Ctrl+A` → `F9`.

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

## Limitations (v1)

- **No source-bookmark refs**: this v1 doesn't address bookmarks that exist in the source document but not bound to numbering. The skill only creates bookmarks for paragraphs it auto-numbers in this run.
- **No PAGEREF**: page-number citations ("on page 12") aren't supported — they require a different field type and depend on Word's pagination layer.
- **No HYPERLINK to external resources**: only intra-document REF fields.
- **No `\p` (above/below) switch**: REF targeting relative position isn't supported.
- **Custom `lvlRestart` not honored in placeholder**: the simulator uses Word's default reset-on-higher-level behavior. Custom restart points produce correct text after `F9` (Word re-renders) but the placeholder may diverge. Cosmetic only.
