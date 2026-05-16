/**
 * <m:r> (math run) emission — the leaf carrier for every character of
 * mathematical text. Every run has:
 *
 *   1. <w:rPr><w:rFonts w:ascii="Cambria Math" .../></w:rPr> — font.
 *      Word ignores any other font in math context; we emit this on
 *      every run to keep round-tripping into Word's own export shape.
 *   2. <m:rPr><m:sty m:val="…"/></m:rPr> — only when the style differs
 *      from what Word would default for the leaf kind:
 *        - mi length-1 → italic (no sty)
 *        - mi length>1 → plain (sty=p)
 *        - mn / mo / mtext / ms → plain (sty=p NOT needed; Cambria Math
 *          renders these plain by default — TEI omits it, we follow)
 *      An explicit `mathvariant` overrides the default.
 *   3. <m:t>text</m:t>
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

  const wRPr = wEl(doc, "rPr")
  const rFonts = wEl(doc, "rFonts")
  rFonts.setAttributeNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w:ascii",
    CAMBRIA,
  )
  rFonts.setAttributeNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w:eastAsia",
    CAMBRIA,
  )
  rFonts.setAttributeNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w:hAnsi",
    CAMBRIA,
  )
  rFonts.setAttributeNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w:cs",
    CAMBRIA,
  )
  wRPr.appendChild(rFonts)
  r.appendChild(wRPr)

  const sty = resolveStyle(kind, text, mathvariant)
  if (sty !== undefined) {
    const mRPr = mEl(doc, "rPr")
    const styEl = mEl(doc, "sty")
    setMVal(styEl, sty)
    mRPr.appendChild(styEl)
    r.appendChild(mRPr)
  }

  const t = mEl(doc, "t")
  // OMML's <m:t> with default xml:space="default" trims whitespace.
  // Math text often includes meaningful spaces (e.g. \, → thin space),
  // so always tag preserve when text has any leading/trailing space.
  if (text.length > 0 && (text.startsWith(" ") || text.endsWith(" "))) {
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
  }
  t.textContent = text
  r.appendChild(t)

  return r
}
