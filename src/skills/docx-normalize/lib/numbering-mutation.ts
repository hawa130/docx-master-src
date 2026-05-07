import { NS } from "@core/types.ts"
import {
  firstChildNS,
  getChildren,
  getChildrenNS,
  wAttr,
} from "@core/xml-utils.ts"
import type { NumberingConfig } from "./types.ts"

/* ------------- numbering.xml manipulation ------------- */

export type SuffValue = "tab" | "space" | "nothing"

/**
 * Decide the marker-suffix character and the lvlText that should actually be
 * written to OOXML. When the user sets `suff` explicitly, that wins. Otherwise
 * we read the user's hand-written `lvlText` for trailing ASCII spaces — 0 → no
 * gap, 1 → one space, 2+ → tab — which captures their format intent without
 * us having to detect "is this CJK or Western punctuation". Trailing spaces
 * are stripped from the emitted lvlText so the gap is owned solely by suff
 * (otherwise a `"%1. "` lvlText with `suff=space` would render double-spaced).
 */
export function resolveSuff(
  lvlText: string,
  explicit?: SuffValue,
): { suff: SuffValue; effectiveLvlText: string } {
  const trailing = lvlText.match(/ *$/)![0].length
  const inferred: SuffValue = trailing === 0 ? "nothing" : trailing === 1 ? "space" : "tab"
  return {
    suff: explicit ?? inferred,
    effectiveLvlText: lvlText.replace(/ +$/, ""),
  }
}

export function injectNumbering(numberingDoc: Document, config: NumberingConfig): string {
  const w = NS.w
  const root = numberingDoc.documentElement!
  // pick fresh abstractNumId and numId
  const existingAbsIds = getChildrenNS(root, w, "abstractNum").map((e) =>
    parseInt(wAttr(e, "abstractNumId") || "0", 10),
  )
  const existingNumIds = getChildrenNS(root, w, "num").map((e) =>
    parseInt(wAttr(e, "numId") || "0", 10),
  )
  const nextAbs = (existingAbsIds.length ? Math.max(...existingAbsIds) : -1) + 1
  const nextNum = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1

  const abs = numberingDoc.createElementNS(w, "w:abstractNum")
  abs.setAttributeNS(w, "w:abstractNumId", String(nextAbs))
  for (const lvl of config.levels) {
    const lvlEl = numberingDoc.createElementNS(w, "w:lvl")
    lvlEl.setAttributeNS(w, "w:ilvl", String(lvl.level))
    const start = numberingDoc.createElementNS(w, "w:start")
    start.setAttributeNS(w, "w:val", String(lvl.start ?? 1))
    lvlEl.appendChild(start)
    const numFmtEl = numberingDoc.createElementNS(w, "w:numFmt")
    numFmtEl.setAttributeNS(w, "w:val", lvl.numFmt)
    lvlEl.appendChild(numFmtEl)
    const { suff, effectiveLvlText } = resolveSuff(lvl.lvlText, lvl.suff)
    const suffEl = numberingDoc.createElementNS(w, "w:suff")
    suffEl.setAttributeNS(w, "w:val", suff)
    lvlEl.appendChild(suffEl)
    const lvlTextEl = numberingDoc.createElementNS(w, "w:lvlText")
    lvlTextEl.setAttributeNS(w, "w:val", effectiveLvlText)
    lvlEl.appendChild(lvlTextEl)
    const lvlJc = numberingDoc.createElementNS(w, "w:lvlJc")
    lvlJc.setAttributeNS(w, "w:val", "left")
    lvlEl.appendChild(lvlJc)
    const pStyle = numberingDoc.createElementNS(w, "w:pStyle")
    pStyle.setAttributeNS(w, "w:val", lvl.styleId)
    lvlEl.appendChild(pStyle)
    if (lvl.numRPr) {
      const rPr = buildLvlRPr(numberingDoc, lvl.numRPr)
      if (rPr) lvlEl.appendChild(rPr)
    }
    abs.appendChild(lvlEl)
  }
  // abstractNum must come before num children — insert before any existing num
  const firstNum = getChildrenNS(root, w, "num")[0]
  if (firstNum) root.insertBefore(abs, firstNum)
  else root.appendChild(abs)

  const num = numberingDoc.createElementNS(w, "w:num")
  num.setAttributeNS(w, "w:numId", String(nextNum))
  const absRef = numberingDoc.createElementNS(w, "w:abstractNumId")
  absRef.setAttributeNS(w, "w:val", String(nextAbs))
  num.appendChild(absRef)
  root.appendChild(num)

  return String(nextNum)
}

function buildLvlRPr(
  doc: Document,
  spec: NonNullable<NumberingConfig["levels"][number]["numRPr"]>,
): Element | null {
  const w = NS.w
  const rPr = doc.createElementNS(w, "w:rPr")
  if (spec.fontLatin || spec.fontCJK) {
    const rFonts = doc.createElementNS(w, "w:rFonts")
    if (spec.fontLatin) {
      rFonts.setAttributeNS(w, "w:ascii", spec.fontLatin)
      rFonts.setAttributeNS(w, "w:hAnsi", spec.fontLatin)
    }
    if (spec.fontCJK) rFonts.setAttributeNS(w, "w:eastAsia", spec.fontCJK)
    rPr.appendChild(rFonts)
  }
  if (spec.size !== undefined) {
    const sz = doc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(Math.round(spec.size * 2)))
    rPr.appendChild(sz)
  }
  if (spec.bold !== undefined) {
    const b = doc.createElementNS(w, "w:b")
    if (!spec.bold) b.setAttributeNS(w, "w:val", "0")
    rPr.appendChild(b)
  }
  if (spec.italic !== undefined) {
    const i = doc.createElementNS(w, "w:i")
    if (!spec.italic) i.setAttributeNS(w, "w:val", "0")
    rPr.appendChild(i)
  }
  if (spec.color) {
    const color = doc.createElementNS(w, "w:color")
    color.setAttributeNS(w, "w:val", spec.color)
    rPr.appendChild(color)
  }
  return getChildren(rPr).length > 0 ? rPr : null
}

export function attachNumberingToStyle(
  stylesDoc: Document,
  styleId: string,
  numId: string,
  level: number,
) {
  const w = NS.w
  const styleEl = getChildrenNS(stylesDoc.documentElement!, w, "style").find(
    (s) => wAttr(s, "styleId") === styleId,
  )
  if (!styleEl) return
  let pPr = firstChildNS(styleEl, w, "pPr")
  if (!pPr) {
    pPr = stylesDoc.createElementNS(w, "w:pPr")
    styleEl.appendChild(pPr)
  }
  // remove existing numPr
  const existing = firstChildNS(pPr, w, "numPr")
  if (existing) pPr.removeChild(existing)

  const numPr = stylesDoc.createElementNS(w, "w:numPr")
  const ilvl = stylesDoc.createElementNS(w, "w:ilvl")
  ilvl.setAttributeNS(w, "w:val", String(level))
  numPr.appendChild(ilvl)
  const numIdEl = stylesDoc.createElementNS(w, "w:numId")
  numIdEl.setAttributeNS(w, "w:val", numId)
  numPr.appendChild(numIdEl)
  pPr.insertBefore(numPr, pPr.firstChild)
}
