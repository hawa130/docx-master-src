# Captions (`captions` table + `CaptionBlock` + `EquationBlock.captionId`)

Word-native caption-class numbering — SEQ + STYLEREF fields + bookmark
+ REF `\h`. Used for figure / table / equation / theorem / lemma /
... — anything that needs a chapter-prefixed enumerator with cross-
references.

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

### Chapter-prefixed (中文学术 / GB/T 7713)

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

For English academic: same shape, replace `"图 "` / `"表 "` with
`"Figure "` / `"Table "` and `bodySeparator: "  "` with `": "`.

For short papers / no chapters: drop `chapterPrefix` (global counter).

For theorem-class / lemmas / corollaries: same shape, custom
identifier + prefix.

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

On EquationBlock: omit `subGroup` for standalone `(1)` / `(2)` / `(3)`;
`"subGroup": "start"` for `(2a)`; `"subGroup": "continue"` for `(2b)` /
`(2c)` / ...

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

For caption-class targets, `display: "label"` / `"number"` / `"full"`
all return the SEQ-rendered text with full decoration (prefix +
chapter + counter + suffix). The variants collapse on this target
class — the bookmark wraps just the number range, so REF `\h` returns
that whether you ask for label / number / full. Full routing rules in
[`cross-references.md`](cross-references.md).

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

`edit-caption` replaces only the body text — SEQ / STYLEREF fields and
bookmark stay intact so cross-references keep resolving. Throws on
EquationBlock targets (no body to edit) — to change an equation,
delete + re-emit. `replace_paragraph` / paragraph-level destructive ops
on caption paragraphs are blocked by the field scan.

## Resetting a caption counter mid-document

```jsonc
{ "type": "caption-counter-reset", "captionId": "Equation", "newValue": 1 }
```

Standalone marker — the **next** caption of `captionId` renders with
its counter at `newValue` (default 1). Useful for appendix sequences
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

Existing SEQ captions in the source doc get re-rendered with the
current `captions[<id>]` config on each apply — useful when iterating
on prefix / chapter-prefix shape across runs. Bookmark + identifier +
body text preserved. Identifier mismatch (SEQ exists for an
unconfigured identifier) → pass through + warn.

## Discovering existing captions

- `overview` shows a Captions section listing SEQ identifiers found in
  body (skip-if-empty).
- `inspect_caption <doc>` lists all identifiers with occurrence counts +
  referencing-REF counts.
- `inspect_caption <doc> <identifier>` dumps per-paragraph details.
- `migrate_captions <doc>` detects manually-numbered caption-shaped
  paragraphs (e.g. "图 2.1: ..." typed by hand, no SEQ field) and
  suggests identifiers — agent builds the apply config to convert.
