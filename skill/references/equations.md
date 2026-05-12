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

## Cross-references — caption + anchor pattern

Same as figures and tables: a separate caption paragraph carries the numbering binding + `anchor`; body text uses `InlineRef` against the anchor (see `cross-references.md`).

```jsonc
{ "type": "equation", "latex": "\\sum_{i=1}^{n} a_i" },
{ "type": "paragraph", "styleId": "EquationCaption",
  "anchor": "eq-sum", "text": "求和定义" }
```

Caption sits below the equation. The caption paragraph (not the equation) is the InlineRef target — equation paragraphs hold no readable text, so `display: "full"` against them resolves to empty.

For right-aligned equation numbering on the same line as the equation (IEEE convention), wrap both in a 3-column borderless `TableBlock`. Construction follows `tables.md`.

## Integration with the style system

| What | How |
|---|---|
| Centered display | Word's default OMML rendering centers `<m:oMathPara>`. Override via `styleId` + `paraFormat.alignment` if you want left/right. |
| Caption numbering | EquationCaption gets a single-level continuous numbering scheme — same shape as FigureCaption / TableCaption. See `numbering-formats.md`. |
| pattern_rules / bulk_rules | Apply to the caption paragraph normally. The equation paragraph contains no `<w:r>` runs (the OMML is its sole child), so text-pattern matching skips it. |

## Edge cases

- **Inline math inside an InlineRef's display text.** Not supported. InlineRef emits the resolved text as `<w:r>` runs; embedding `<m:oMath>` inside a REF field is non-standard. Place the InlineRef and the math as sibling InlineNodes instead.
- **trackChanges + equation insert.** Engine throws at emit. OOXML's tracked-change wrappers don't have a clean "equation inserted" shape. Run equation insertion without trackChanges; use trackChanges for subsequent surrounding edits.

## What's not supported (v1)

- **n-ary operator structural bug.** `\sum_{i=1}^{n} i^2`, `\int_0^1 x \, dx`, `\prod_{k=1}^{n} a_k` render with a dashed empty box (`<m:e/>`) before the operand — Word still computes correctly, but the visual is wrong. Comes from `mathml2omml` (LGPL); fix planned for v2 (self-built MathML→OMML translator).
- **Equation numbering via LaTeX `\tag{}` / `\label{}` / `equation` environment.** Use the caption + anchor pattern instead — gives you the same cross-ref reach and matches the figure / table convention.
- **Editing an existing equation in the docx.** Locators target paragraphs, not OMML subtrees. To change an equation, `replace` the whole paragraph.
