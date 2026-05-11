# Cross-references (REF fields)

## When to use this

Never write `"如图 1 所示"` or `"见 [3]"` as plain runs — use `InlineRef` so Word resolves the cite at render time. Page-number citations (`PAGEREF`) are out of scope; surface to the user.

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
  - `{ "type": "anchor", "name": "fig-arch" }` — named bookmark. Resolves against (a) `ParagraphBlock.anchor` declared earlier in this same `edits[]` array, or (b) bookmarks already in the source document that wrap a single paragraph. Name must match `^[A-Za-z_][A-Za-z0-9_-]{0,39}$`.
- `display` — what Word renders. `\h` is always added so the rendered text is clickable:
  - `"label"` (default) — full numbered paragraph text from `lvlText`: `图 1` / `1.2` / `第二章`. Switch: `\n \h`.
  - `"number"` — paragraph number in relative context via `\r \h`. **Reliable only for single-level schemes** (figure captions, `[%1]` reference lists), where it equals the counter. For multi-level schemes prefer `"label"`.
  - `"full"` — target paragraph's text content (caption title without auto-num prefix). Switch: `\h`. **Does not require an auto-numbered target.**
- `format` — optional `RunFormat` (color, italic, size, …) applied to the rendered text. Format-bearing refs require Word round-trip to verify — see below.

## Target requirements

`label` / `number` require an auto-numbered target (bound to a `numbering[]` level — either pre-existing in the source's `pStyle → numId` binding or freshly bound by this run). Word's `\n` / `\r` switches render from the numbering binding; an unbound target produces nothing.

`full` works on any paragraph — the bookmark resolves to text content directly.

When `label` / `number` hits an unbound target, apply refuses with a message naming the fix (bind a `numbering[]` level, or switch to `display: "full"`).

## Named anchors — ref paragraphs created in the same apply

When an `edits[]` insert creates a paragraph later refs will cite, give it an `anchor` so the refs can address it by name:

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

Anchors are processed in `edits[]` array order: a ref can address any anchor declared earlier, but not a later one (reorder edits if needed). Anchor names share a flat namespace with source bookmarks — picking one that already exists throws at apply time.

`anchor` on an insert with no current ref is still useful: the bookmark stays in the output for future apply runs to ref by name.

## First-open behavior and format verification

The placeholder text between `<w:fldChar fldCharType="separate"/>` and `<w:fldChar fldCharType="end"/>` is pre-computed at apply time from the simulated numbering counters, so the document reads correctly even if the user dismisses Word's field-update prompt. `<w:updateFields w:val="true"/>` is set in `settings.xml`, so the prompt fires automatically on next open; manual `Ctrl+A` → `F9` updates after that.

### Verifying format-bearing refs

Inspecting the apply-time XML alone is **not enough** for refs carrying a `format`. Word rewrites the field's result run on every update; a broken emit (missing rPr replication or `\* MERGEFORMAT`) still validates and reads correctly on first open, but the format disappears the moment Word updates.

Round-trip to verify:
1. Open the apply output in Word
2. Accept the "update fields" prompt and save
3. Unzip and grep `word/document.xml` — confirm the result run still carries the expected `<w:rPr>`

Schema validation and emit-time inspection both pass the broken case; only a real Word round-trip catches the regression.

## Example — reference list with explicit brackets

`display: "number"` returns just the counter; wrap with literal `[` / `]` to control the bracket form independently of the target's lvlText. For brackets baked into lvlText (`lvlText="[%1]"`) use `display: "label"` instead.

```jsonc
{ "type": "paragraph", "text": [
  { "text": "本文方法借鉴了 [" },
  { "refTo": { "type": "paragraph", "index": 180 }, "display": "number" },
  { "text": "] 中的思路。" }
]}
// → 本文方法借鉴了 [3] 中的思路。
```

## Limitations

- **No PAGEREF**: page-number citations aren't supported (different field type, depends on Word's pagination layer).
- **No forward anchor refs**: a ref can only address an anchor declared earlier in the same `edits[]`. Reorder edits if needed.
- **Source bookmarks must wrap a single paragraph**: source bookmarks spanning multiple paragraphs or sitting at body level aren't resolvable via `refTo: { type: "anchor" }` — silently skipped at allocator construction.
- **Custom `lvlRestart` not honored in placeholder**: the counter simulator uses Word's default reset-on-higher-level. Custom restart points still produce correct text after Word updates the fields; the pre-update placeholder may diverge. Cosmetic only.
