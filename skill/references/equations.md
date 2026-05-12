# Equations (`{ "type": "equation", ... }` and `{ "math": "..." }`)

LaTeX → OMML rendering for display equations and inline math.

## Quick start

```jsonc
// Display (block)
{ "type": "equation", "latex": "E = mc^2" }

// Inline (inside a paragraph's text array)
{ "type": "paragraph", "text": [
  { "text": "By Pythagoras, " },
  { "math": "a^2 + b^2 = c^2" },
  { "text": "." }
]}
```

## Schema reference

```ts
EquationBlock {
  type: "equation"
  latex: string              // LaTeX expression (Temml subset — see below)
  styleId?: string           // paragraph style for the wrapping <w:p>
  paraFormat?: ParagraphFormat
  anchor?: string            // bookmark; emits valid, but caption paragraph is the practical InlineRef target — see below
}

InlineEquation { math: string }   // appears inside Paragraph.text[] alongside { text, format } and { refTo }
```

## LaTeX coverage

The renderer is [Temml](https://temml.org/) — supports most of LaTeX math mode. Inline math doesn't number itself; use the caption pattern below for any numbered equation, regardless of display vs inline shape.

## Numbering and cross-references

Standard 学术 / IEEE / GB/T 7713 convention: equation centered on its line, number flush right on the same line — `equation ........... (1)`. Compose via a 3-column borderless `TableBlock` (empty / equation / numbered paragraph):

```jsonc
{ "type": "table", "borders": "none",
  "cols": [{ "width": "auto" }, { "width": 300 }, { "width": "auto" }],
  "rows": [[
    "",
    [{ "type": "equation", "latex": "\\sum_{i=1}^{n} a_i" }],
    [{ "type": "paragraph", "styleId": "EquationNumber",
       "paraFormat": { "alignment": "right" },
       "anchor": "eq-sum", "text": "" }]
  ]]}
```

`EquationNumber` is bound to a single-level counter scheme (e.g. `lvlText: "(%1)"`) so the right cell renders as `(1)`, `(2)`, … automatically. Body-text refs target the numbered paragraph via `InlineRef` against its `anchor` (see `cross-references.md`).

The equation paragraph itself holds no readable text, so cross-refs must anchor on the numbered paragraph, not on the equation.

## Integration with the style system

| What | How |
|---|---|
| Centered display | Display equations center by default. Override via `styleId` + `paraFormat.alignment` for left/right. |
| Caption numbering | EquationCaption gets a single-level continuous numbering scheme — same shape as FigureCaption / TableCaption. See `numbering-formats.md`. |
| pattern_rules / bulk_rules | Apply to the caption paragraph normally. The equation paragraph contains no `<w:r>` runs (the OMML is its sole child), so text-pattern matching skips it. |

## Edge cases

- **Inline math inside an InlineRef's display text.** Not supported. InlineRef emits the resolved text as `<w:r>` runs; embedding `<m:oMath>` inside a REF field is non-standard. Place the InlineRef and the math as sibling InlineNodes instead.
- **trackChanges + equation insert.** Engine throws at emit. OOXML's tracked-change wrappers don't have a clean "equation inserted" shape. Run equation insertion without trackChanges; use trackChanges for subsequent surrounding edits.

## What's not supported (v1)

- **n-ary operator structural bug.** `\sum_{i=1}^{n} i^2`, `\int_0^1 x \, dx`, `\prod_{k=1}^{n} a_k` render with a dashed empty box (`<m:e/>`) before the operand — Word still computes correctly, but the visual is wrong. Comes from `mathml2omml` (LGPL); fix planned for v2 (self-built MathML→OMML translator).
- **Equation numbering via LaTeX `\tag{}` / `\label{}` / `equation` environment.** Use the caption + anchor pattern instead — gives you the same cross-ref reach and matches the figure / table convention.
- **Editing an existing equation in the docx.** Locators target paragraphs, not OMML subtrees. To change an equation, `replace` the whole paragraph.
