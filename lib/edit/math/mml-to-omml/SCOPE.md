# MathML → OMML — design scope

## Principle

Translate MathML semantics into OMML semantics **faithfully**. Where
OMML can't express a MathML concept, render the closest available
shape *without inventing new semantics*. Don't bend one OMML idiom to
mimic an unrelated MathML feature — those patches accumulate, drift
visually, and confuse the next maintainer about what we actually
represent.

If a user wants behaviour we don't cover, the `omml` escape hatch on
`EquationBlock` ships hand-authored OMML through unchanged.

## What OMML cannot express (don't patch — document)

These belong in the `omml` escape hatch when needed, not in this
converter.

- **`\big` / `\Big` / `\bigg` / `\Bigg` standalone sized delimiters.**
  LaTeX intent is fixed-scale (1.2×/1.8×/2.4×/3.0× base). OMML has no
  per-character size modifier on `<m:r>` (Word/LO ignore `<w:sz>`
  inside math runs). Auto-grow via `<m:d>` is a *different* semantic
  (stretches to body height); using it for `\big` would silently
  swap one user intent for another. Render as plain `<mo>` at default
  size and accept that subscripted bodies may look taller.
- **Negative or graded spacing — `\!`, `\,`, `\:`, `\;`, `\quad`,
  `\hspace{-1em}`.** OMML `<m:r>` text content has no width or
  kerning control beyond character flow. We emit `<mspace>` as a
  single regular space and rely on Word's math spacing rules; finer
  gradations don't survive the round-trip.
- **Phantom dimensions — `\hphantom`, `\vphantom`, `\smash`.**
  `<m:phant>` exists in OMML but lacks the axis-selective sizing
  these LaTeX primitives need. We emit `<m:phant>` for `<mphantom>`;
  the directional variants approximate.
- **Overlap macros — `\mathrlap`, `\mathllap`, `\rlap`, `\llap`.**
  No OMML equivalent.
- **MathML attributes on `<mpadded>` — `lspace`, `voffset`, `width`,
  `height`, `depth`.** We treat `<mpadded>` as a transparent
  wrapper (emit children only). Fine spacing tweaks the attribute
  encodes are lost.
- **Right-to-left math.** OMML doesn't model RTL math; Arabic /
  Hebrew identifiers in `<mi>` flow LTR like any other character.
- **MathML `align` / `groupalign` / `decimalpoint` on `<mtable>`.**
  OMML `<m:m>` has alignment controls but at coarser granularity. We
  emit the matrix without alignment hints.
- **Stretchy operators with explicit `minsize`/`maxsize` ranges.**
  Sized delimiter detection in OMML is tied to `<m:d>` paired
  fences (auto-grow), not standalone characters. Bare sized `<mo>`
  renders at default.

## What Word/LibreOffice render differently from OMML semantics

These are renderer quirks on output that IS semantically correct
per ECMA-376. We document, not patch.

- **`\not\subset` and other negated relations with combining slash.**
  Render quality depends on the font's combining-character support
  (Cambria Math handles most; CJK font substitution can split the
  slash from the base relation).
- **`mathvariant` substitution coverage.** temml usually emits the
  Unicode codepoint directly (U+1D6AA for `\boldsymbol{\Gamma}`),
  so the variant attribute is informational by the time we see it.
  Variants that don't have Unicode codepoints (sans-serif math)
  depend on Cambria Math glyph coverage in the rendering Word.
- **`<m:t>` with mixed CJK + Latin.** Word's `<m:r>` font selection
  is `Cambria Math` for all four font slots; CJK substitution kicks
  in at the renderer level using the document's eastAsia font.
- **Element ordering / spacing in `<m:nary>` with empty operand.**
  `\sum_{i=1}^n` standalone (no operand) is technically valid; Word
  may render a dashed empty box for the missing operand. Add an
  operand or use the `omml` escape hatch if that placeholder is
  unacceptable.

## What we deliberately *do* fuse / re-shape

These are MathML idioms whose OMML representation differs structurally;
faithful translation requires recognition, not new semantic invention.

- **mrow with `[fence-prefix-mo, …body…, fence-postfix-mo]`** →
  `<m:d>` with begChr/endChr and `<m:grow/>`. The `fence` attribute
  is MathML's signal that these are paired delimiters; OMML's
  `<m:d>` is the same concept.
- **mrow with an n-ary operator (Σ/∫/∏/…) and following operand
  siblings** → `<m:nary>` with operand inside `<m:e>`. The siblings
  *are* the operand per LaTeX-from-MathML convention; OMML's
  `<m:nary>` requires the operand inline.
- **Single-child `<mrow>` / `<mstyle>` / `<mpadded>` / `<menclose>`**
  → flatten before fusion passes. These wrappers carry no
  semantic at single-child arity (mrow grouping); for mpadded/mstyle
  this is a knowing trade — we lose the spacing/style attributes
  to keep n-ary fusion working through temml's wrapping habit.
- **`<mover>` / `<munder>` with bracket-class character (⏞ ⏟ ⏜ ⏝
  ⎴ ⎵ ︷ ︸)** → `<m:groupChr>`. The character class is the OMML
  equivalent of MathML's grouping bracket semantic.
- **`<mover>` with single-char `<mo>`** → `<m:acc>`. LaTeX accents
  (`\hat`, `\vec`, `\bar`, `\tilde`, …) all emit this shape; the
  single-character condition matches the LaTeX intent.

## When in doubt

If a translation requires:
- inferring intent from sibling structure beyond one mrow scope, or
- mapping a MathML attribute that has no OMML peer, or
- selecting an OMML idiom whose semantic differs from the source —

stop and document instead. The `omml` escape hatch covers the long
tail; the converter handles the body.
