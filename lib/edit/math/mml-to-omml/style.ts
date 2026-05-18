/**
 * Math character style resolution. Determines whether to emit
 * `<m:sty m:val="…"/>` inside a run, based on:
 *   - the leaf kind (mi/mn/mo/mtext/ms)
 *   - the text (mi length-1 vs length>1 drives MathML's italic default)
 *   - an explicit `mathvariant` attribute on the source element
 *
 * Returns the OMML sty value or `undefined` (use Word's default).
 *
 * MathML default conventions (from MathML 3 §3.2.3):
 *   - mi length 1 → italic
 *   - mi length >1 → normal/upright (function-name convention)
 *   - mn, mo, mtext, ms → normal
 *
 * Word's defaults in Cambria Math math context:
 *   - Plain text in <m:r> renders italic for single-codepoint Latin
 *     letters, plain for digits/operators/multi-char text
 *   - <m:sty m:val="p"/> forces plain; "i" forces italic; "b" bold;
 *     "bi" bold-italic
 *
 * So emit sty ONLY when our intent diverges from Word's leaf default:
 *   - mi with length>1: emit "p" (Word would render italic on Latin chars)
 *   - explicit mathvariant: emit corresponding mapping
 */

import { MATHVARIANT_STYLE } from "./constants.ts"

export type LeafKind = "mi" | "mn" | "mo" | "mtext" | "ms" | "mspace"

export type StyleVal = "p" | "b" | "i" | "bi"

export function resolveStyle(
  kind: LeafKind,
  text: string,
  mathvariant: string | undefined,
): StyleVal | undefined {
  if (mathvariant !== undefined) {
    const explicit = MATHVARIANT_STYLE.get(mathvariant)
    if (explicit !== undefined) return explicit
    // Other mathvariant values (script, fraktur, double-struck, ...) are
    // realized via Unicode codepoint substitution. temml typically emits
    // the substituted character directly in mi text content, so by the
    // time we see it here the variant attribute is informational only.
    return undefined
  }
  if (kind === "mi") {
    // MathML convention: single-codepoint mi defaults to italic (no sty
    // needed — Word matches). Multi-codepoint mi is a function name or
    // identifier that must render upright.
    return codepointLength(text) > 1 ? "p" : undefined
  }
  // mn / mo / mtext / ms / mspace — Word defaults plain in Cambria Math
  // context, so no explicit sty.
  return undefined
}

/** Codepoint count, not UTF-16 unit count — keeps surrogate pairs like
 *  ℒ (U+2112) and 𝐀 (U+1D400) counted as one symbol. */
function codepointLength(s: string): number {
  let n = 0
  for (const _ of s) n++
  return n
}
