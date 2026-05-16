/**
 * <m:r> (math run) emission — the leaf carrier for every character of
 * mathematical text. Per ECMA-376 §22.1.2.87 CT_R the child sequence is:
 *
 *   1. <m:rPr> (math run properties)               — must come first
 *   2. EG_RPrMath, which includes <w:rPr>          — font selection
 *   3. text content elements (<m:t>, <m:br>, …)
 *
 * Putting <w:rPr> before <m:rPr> is the exact bug the
 * mathml2omml package shipped — libxml2 schema-validation rejected it.
 *
 *   - <m:rPr><m:sty m:val="…"/></m:rPr> — only when the style differs
 *     from what Word would default for the leaf kind:
 *       - mi length-1 → italic (no sty)
 *       - mi length>1 → plain (sty=p)
 *       - mn / mo / mtext / ms → plain (TEI omits sty here; we follow)
 *   - <w:rPr><w:rFonts w:ascii="Cambria Math" .../></w:rPr> — every
 *     run; matches Word's own export and keeps math text in the math
 *     font regardless of paragraph font.
 *   - <m:t>text</m:t>
 */

import type { LeafKind } from "./style.ts"
import { resolveStyle } from "./style.ts"
import { mEl, wEl, setMVal } from "./dom.ts"

const CAMBRIA = "Cambria Math"

/** Build a single <m:r> with the given text and styling. Empty text
 *  still emits a run so adjacent runs don't merge in serialization. */
export function buildRun(
  doc: Document,
  text: string,
  kind: LeafKind,
  mathvariant: string | undefined,
): Element {
  const r = mEl(doc, "r")

  // 1. <m:rPr> (math run properties) — must come before <w:rPr> per
  // CT_R's sequence ordering. Two flags can apply:
  //   - <m:nor/>: marks the run as literal (non-math) text. Required
  //     for <mtext> to prevent Word/LO from interpreting letter runs
  //     as math identifiers (e.g. "and" was rendering as ∧ in LO
  //     without nor — math-mode autocorrect kicked in).
  //   - <m:sty m:val="…"/>: explicit style override.
  const sty = resolveStyle(kind, text, mathvariant)
  const isLiteralText = kind === "mtext" || kind === "ms"
  if (sty !== undefined || isLiteralText) {
    const mRPr = mEl(doc, "rPr")
    if (isLiteralText) mRPr.appendChild(mEl(doc, "nor"))
    if (sty !== undefined) {
      const styEl = mEl(doc, "sty")
      setMVal(styEl, sty)
      mRPr.appendChild(styEl)
    }
    r.appendChild(mRPr)
  }

  // 2. <w:rPr> with font selection.
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const wRPr = wEl(doc, "rPr")
  const rFonts = wEl(doc, "rFonts")
  rFonts.setAttributeNS(W, "w:ascii", CAMBRIA)
  rFonts.setAttributeNS(W, "w:eastAsia", CAMBRIA)
  rFonts.setAttributeNS(W, "w:hAnsi", CAMBRIA)
  rFonts.setAttributeNS(W, "w:cs", CAMBRIA)
  wRPr.appendChild(rFonts)
  r.appendChild(wRPr)

  // 3. <m:t> text content. Literal-text runs (mtext, ms) are bound
  // to need xml:space="preserve" because their leading/trailing space
  // carries typographic intent (`\text{ if and only if }` would
  // collide with adjacent math without it).
  const t = mEl(doc, "t")
  // OMML's <m:t> with default xml:space="default" trims whitespace.
  // Math text often includes meaningful spaces (e.g. \, → thin space,
  // \text{ if and only if }), so tag preserve whenever the text has
  // leading or trailing whitespace, or whenever it's an mtext/ms run
  // (literal text always wants preservation of internal spacing).
  const needsPreserve =
    text.length > 0 && (text.startsWith(" ") || text.endsWith(" ") || isLiteralText)
  if (needsPreserve) {
    t.setAttribute("xml:space", "preserve")
  }
  t.textContent = text
  r.appendChild(t)

  return r
}
