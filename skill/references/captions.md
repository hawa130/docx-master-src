# Captions (`captions` table + `CaptionBlock` + `EquationBlock.captionId`)

Word-native caption-class numbering — SEQ + STYLEREF fields + bookmark
+ REF `\h`. Used for figure / table / equation / theorem / lemma /
... — anything that needs a chapter-prefixed enumerator with cross-
references.

## Why SEQ (not numPr)

Word's data model splits "numbered things" into two kinds:

- **`numPr` / `numbering.xml`** — outline numbering (H1–H6), bullet /
  ordered lists. Structural. (Keep using this for headings + lists.)
- **`SEQ` fields** — caption-class enumerators. Document-local counters
  keyed by identifier. (Use this for captions.)

A previous version of the skill used numPr for captions too. That worked
mechanically but lost: Word UI's References → Cross-reference dialog
integration, clean chapter-prefix decoupling, independent counters per
caption type. The captions table fixes all three.

## Quick start

```jsonc
// Top-level apply config
{
  "captions": {
    "Equation": {
      "prefix": "(",
      "suffix": ")",
      "chapterPrefix": ["Heading1"],
      "styleId": "EquationNumber"
    },
    "Figure": {
      "prefix": "图 ",
      "chapterPrefix": ["Heading1"],
      "bodySeparator": "  ",
      "styleId": "FigureCaption"
    },
    "Table": {
      "prefix": "表 ",
      "chapterPrefix": ["Heading1"],
      "bodySeparator": "  ",
      "styleId": "TableCaption"
    }
  },
  "edits": [
    {
      "op": "insert-after",
      "at": { "type": "paragraph", "index": 10 },
      "content": [
        { "type": "equation", "latex": "a^2 + b^2 = c^2",
          "captionId": "Equation", "anchor": "eq-pythagoras" },
        { "type": "caption", "captionId": "Figure",
          "text": "系统架构示意图", "anchor": "fig-arch" }
      ]
    }
  ]
}
```

## Schema reference

```ts
captions: Record<string, CaptionEntry>

CaptionEntry {
  prefix?: string                   // literal before the counter (default "")
  suffix?: string                   // literal after the counter (default "")
  format?: "arabic" | "alphabetic" | "ALPHABETIC"
         | "roman" | "ROMAN" | "chinese" | "chinese-formal"
                                    // default "arabic"
  chapterPrefix?: string[]          // ordered styleIds (any depth);
                                    // default [] (global, no restart)
  chapterSeparator?: string         // joins chapter levels + counter
                                    // (default ".")
  bodySeparator?: string            // between counter and CaptionBlock.text
                                    // (default " ")
  styleId: string                   // REQUIRED — caption paragraph's style
  subCounter?: {                    // enables subequations (1a)(1b)
    format?: "arabic" | "alphabetic" | ...   // default "alphabetic"
    prefix?: string                 // default ""
    suffix?: string                 // default ""
  }
}
```

Block-level:

```ts
CaptionBlock {
  type: "caption"
  captionId: string                 // references captions[<id>]
  text: string                      // may be ""
  anchor?: string                   // bookmark for cross-references
}

CaptionCounterReset {
  type: "caption-counter-reset"
  captionId: string
  newValue?: number                 // default 1
}
```

## Templates

### 中文学术 (GB/T 7713 — chapter-prefixed, Heading 1 outline-numbered)

```jsonc
"captions": {
  "Equation": { "prefix": "(", "suffix": ")",
                "chapterPrefix": ["Heading1"], "styleId": "EquationNumber" },
  "Figure":   { "prefix": "图 ", "chapterPrefix": ["Heading1"],
                "bodySeparator": "  ", "styleId": "FigureCaption" },
  "Table":    { "prefix": "表 ", "chapterPrefix": ["Heading1"],
                "bodySeparator": "  ", "styleId": "TableCaption" }
}
```

### English academic (chapter-prefixed)

```jsonc
"captions": {
  "Equation": { "prefix": "(", "suffix": ")",
                "chapterPrefix": ["Heading1"], "styleId": "EquationNumber" },
  "Figure":   { "prefix": "Figure ", "chapterPrefix": ["Heading1"],
                "bodySeparator": ": ", "styleId": "FigureCaption" },
  "Table":    { "prefix": "Table ", "chapterPrefix": ["Heading1"],
                "bodySeparator": ": ", "styleId": "TableCaption" }
}
```

### Short paper / no chapters (global continuous)

```jsonc
"captions": {
  "Equation": { "prefix": "(", "suffix": ")", "styleId": "EquationNumber" },
  "Figure":   { "prefix": "Figure ", "bodySeparator": ": ", "styleId": "FigureCaption" }
}
```

### Theorem-class (custom identifier)

```jsonc
"captions": {
  "Theorem": { "prefix": "定理 ", "chapterPrefix": ["Heading1"],
               "bodySeparator": "  ", "styleId": "TheoremStmt" },
  "Lemma":   { "prefix": "引理 ", "chapterPrefix": ["Heading1"],
               "bodySeparator": "  ", "styleId": "LemmaStmt" }
}
```

### Subequations

```jsonc
"captions": {
  "Equation": {
    "prefix": "(", "suffix": ")",
    "chapterPrefix": ["Heading1"],
    "styleId": "EquationNumber",
    "subCounter": { "format": "alphabetic" }
  }
}
```

Then on EquationBlock blocks: omit `subGroup` for standalone (1) / (2) /
(3); set `"subGroup": "start"` for (2a); set `"subGroup": "continue"`
for (2b) / (2c) / ...

## How rendering works (under the hood)

For a caption like `图 2.3` with `Figure` identifier + `chapterPrefix:
["Heading1"]`:

```
[bookmarkStart name="fig-arch"]
  "图 "                                 ← prefix
  { STYLEREF "Heading1" \n }            ← chapter prefix → "2"
  "."                                   ← chapterSeparator
  { SEQ Figure \* ARABIC \s 1 }         ← counter → "3"
[bookmarkEnd]
"  "                                    ← bodySeparator
"系统架构示意图"                          ← body text
```

REF `\h` targeting `fig-arch` returns `"图 2.3"` — the whole bookmark
range, decoration included.

`\s 1` makes the SEQ counter reset at each Heading 1 boundary (counter
sim mirrors this for the placeholder backfill). For two-level
chapter prefix (`["Heading1", "Heading2"]`), `\s 2` instead.

## Cross-references to captions

```jsonc
{ "refTo": { "type": "anchor", "name": "fig-arch" }, "display": "label" }
```

`display: "label"` and `display: "number"` collapse on caption-class
targets — both return the SEQ result with full decoration. `display:
"full"`:
- On CaptionBlock: returns the entire paragraph (number + bodySeparator
  + body). Triggers a secondary internal bookmark allocation around
  the whole paragraph.
- On EquationBlock: throws (no body to return).

## Editing existing captions

Use the `edit-caption` op:

```jsonc
{ "op": "edit-caption",
  "target": { "anchor": "fig-arch" },
  "text": "更新后的标题" }
```

Or target by identifier + body-order index:

```jsonc
{ "op": "edit-caption",
  "target": { "captionId": "Figure", "index": 3 },
  "text": "..." }
```

`edit-caption` replaces only the body text — the SEQ / STYLEREF fields
and bookmark stay intact so cross-references keep resolving. Throws on
EquationBlock targets (no body to edit) — to change an equation, delete
+ re-emit it.

Why not `replace` op on the whole caption paragraph? The field blocker
scan refuses it: a caption paragraph contains a complex field, and
paragraph-level destructive ops would break the SEQ chain. Use
`edit-caption` for body changes, `delete-paragraph` for whole removal.

## Resetting a caption counter mid-document

```jsonc
{ "type": "caption-counter-reset", "captionId": "Equation", "newValue": 1 }
```

Standalone marker. Emits an invisible SEQ `\r N \h`; counter sim resets
the identifier to `newValue` (default 1). Useful for appendix sequences
that don't align with outline-level restart, or multi-section docs
where each major section gets its own caption counter.

## Conflicts and constraints

- A `styleId` referenced in `captions[<id>].styleId` MUST NOT also
  appear in `numbering[].levels[].styleId`. The two mechanisms would
  fight at render time. Engine throws at apply with a clear message.
- `chapterPrefix` references unknown `styleId` → schema throws.
- `chapterPrefix` references a style that isn't outline-numbered →
  pre-scan warns (Word's STYLEREF returns 0 at render time).
- `captionId` references undeclared identifier → schema throws.
- `subGroup` set on EquationBlock without matching `subCounter` config
  → schema throws.
- `anchor` set on EquationBlock without `captionId` → schema throws.

## Standardize re-emit (source-doc captions)

When standardize runs on a doc that already contains SEQ-based captions
(from a prior apply or Word's Insert Caption), it walks the body and
rebuilds the pre-body run sequence of each caption paragraph in place
using the current `captions[<id>]` config. Bookmark id / SEQ identifier
/ body text preserved. Identifier mismatch (SEQ in source whose
identifier isn't declared) → passed through unchanged + warned.

Useful for iterating on captions config across apply runs: change the
prefix once, all existing captions re-render with the new shape on the
next apply.

## Discovering existing captions

- `overview` shows a Captions section listing SEQ identifiers found in
  body (skip-if-empty).
- `inspect_caption <doc>` lists all identifiers with occurrence counts +
  referencing-REF counts.
- `inspect_caption <doc> <identifier>` dumps per-paragraph details.
- `migrate_captions <doc>` detects manually-numbered caption-shaped
  paragraphs (e.g. "图 2.1: ..." typed by hand, no SEQ field) and
  suggests identifiers — agent builds the apply config to convert.
