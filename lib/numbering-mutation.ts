import { NS } from "@lib/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml-utils.ts"
import type { NumberingConfig } from "./config-types.ts"
import { insertPPrIntoStyle } from "./style-mutation.ts"
import { RPR_CHILD_ORDER, insertChildInOrder } from "./xml-order.ts"

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

export function injectNumbering(
  numberingDoc: Document,
  config: NumberingConfig,
): { numId: string; abstractNumId: string } {
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
  // Word and other renderers require <w:multiLevelType>; without it numbering
  // silently fails to render even though the bindings resolve.
  const multiLevelType = numberingDoc.createElementNS(w, "w:multiLevelType")
  multiLevelType.setAttributeNS(w, "w:val", config.levels.length > 1 ? "multilevel" : "singleLevel")
  abs.appendChild(multiLevelType)
  for (const lvl of config.levels) {
    const lvlEl = numberingDoc.createElementNS(w, "w:lvl")
    lvlEl.setAttributeNS(w, "w:ilvl", String(lvl.level))
    const start = numberingDoc.createElementNS(w, "w:start")
    start.setAttributeNS(w, "w:val", String(lvl.start ?? 1))
    lvlEl.appendChild(start)
    const numFmtEl = numberingDoc.createElementNS(w, "w:numFmt")
    numFmtEl.setAttributeNS(w, "w:val", lvl.numFmt)
    lvlEl.appendChild(numFmtEl)
    // CT_Lvl child order matters: start, numFmt, lvlRestart?, pStyle?, isLgl?,
    // suff?, lvlText, lvlPicBulletId?, legacy?, lvlJc?, pPr?, rPr?. Word
    // ignores pStyle silently when written out of order, breaking the
    // style→numbering binding even though the file otherwise validates.
    const pStyle = numberingDoc.createElementNS(w, "w:pStyle")
    pStyle.setAttributeNS(w, "w:val", lvl.styleId)
    lvlEl.appendChild(pStyle)
    if (lvl.isLgl) {
      // <w:isLgl/> forces every cross-level placeholder in lvlText to render
      // as arabic regardless of the referenced level's numFmt. Required for
      // "X.X" headings where Heading1 uses chineseCounting (一、) but
      // Heading3's "%1.%3" should display "1.1" not "一.1".
      const isLgl = numberingDoc.createElementNS(w, "w:isLgl")
      lvlEl.appendChild(isLgl)
    }
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

  return { numId: String(nextNum), abstractNumId: String(nextAbs) }
}

function buildLvlRPr(
  doc: Document,
  spec: NonNullable<NumberingConfig["levels"][number]["numRPr"]>,
): Element | null {
  const w = NS.w
  const rPr = doc.createElementNS(w, "w:rPr")
  const add = (el: Element) => insertChildInOrder(rPr, el, RPR_CHILD_ORDER)
  if (spec.fontLatin || spec.fontCJK) {
    const rFonts = doc.createElementNS(w, "w:rFonts")
    if (spec.fontLatin) {
      rFonts.setAttributeNS(w, "w:ascii", spec.fontLatin)
      rFonts.setAttributeNS(w, "w:hAnsi", spec.fontLatin)
    }
    if (spec.fontCJK) rFonts.setAttributeNS(w, "w:eastAsia", spec.fontCJK)
    add(rFonts)
  }
  if (spec.size !== undefined) {
    const sz = doc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(Math.round(spec.size * 2)))
    add(sz)
  }
  if (spec.bold !== undefined) {
    const b = doc.createElementNS(w, "w:b")
    if (!spec.bold) b.setAttributeNS(w, "w:val", "0")
    add(b)
  }
  if (spec.italic !== undefined) {
    const i = doc.createElementNS(w, "w:i")
    if (!spec.italic) i.setAttributeNS(w, "w:val", "0")
    add(i)
  }
  if (spec.color) {
    const color = doc.createElementNS(w, "w:color")
    color.setAttributeNS(w, "w:val", spec.color)
    add(color)
  }
  return getChildren(rPr).length > 0 ? rPr : null
}

/**
 * Mint a fresh `<w:num>` pointing to `abstractNumId`, with
 * `<w:lvlOverride><w:startOverride val="1"/></w:lvlOverride>` so Word resets
 * the counter to 1. Returns the new numId. Used by the per-instance restart
 * pass for single-level (list-shaped) schemes — every "list instance" in the
 * document gets its own counter via this fork, so a `1.` `2.` `3.` list in
 * Chapter 1 doesn't continue as `4.` `5.` `6.` in Chapter 2.
 */
export function forkNumWithStartOverride(
  numberingDoc: Document,
  abstractNumId: string,
  level: number,
): string {
  const w = NS.w
  const root = numberingDoc.documentElement!
  const existingNumIds = getChildrenNS(root, w, "num").map((e) =>
    parseInt(wAttr(e, "numId") || "0", 10),
  )
  const nextNum = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1
  const num = numberingDoc.createElementNS(w, "w:num")
  num.setAttributeNS(w, "w:numId", String(nextNum))
  const absRef = numberingDoc.createElementNS(w, "w:abstractNumId")
  absRef.setAttributeNS(w, "w:val", abstractNumId)
  num.appendChild(absRef)
  const lvlOverride = numberingDoc.createElementNS(w, "w:lvlOverride")
  lvlOverride.setAttributeNS(w, "w:ilvl", String(level))
  const startOverride = numberingDoc.createElementNS(w, "w:startOverride")
  startOverride.setAttributeNS(w, "w:val", "1")
  lvlOverride.appendChild(startOverride)
  num.appendChild(lvlOverride)
  root.appendChild(num)
  return String(nextNum)
}

/**
 * Set (or replace) paragraph-level `<w:numPr>` on a `<w:p>` element. Inserted
 * at the correct position per CT_PPr schema (after `<w:pStyle>`, before the
 * formatting children). Used by the per-instance restart pass to override the
 * style-level numId binding with a per-run forked numId.
 */
export function setParagraphNumPr(pEl: Element, numId: string, level: number): void {
  const w = NS.w
  const ownerDoc = pEl.ownerDocument!
  let pPr = firstChildNS(pEl, w, "pPr")
  if (!pPr) {
    pPr = ownerDoc.createElementNS(w, "w:pPr")
    pEl.insertBefore(pPr, pEl.firstChild)
  }
  const existing = firstChildNS(pPr, w, "numPr")
  if (existing) pPr.removeChild(existing)

  const numPr = ownerDoc.createElementNS(w, "w:numPr")
  const ilvl = ownerDoc.createElementNS(w, "w:ilvl")
  ilvl.setAttributeNS(w, "w:val", String(level))
  numPr.appendChild(ilvl)
  const numIdEl = ownerDoc.createElementNS(w, "w:numId")
  numIdEl.setAttributeNS(w, "w:val", numId)
  numPr.appendChild(numIdEl)

  const pStyle = firstChildNS(pPr, w, "pStyle")
  if (pStyle && pStyle.nextSibling) pPr.insertBefore(numPr, pStyle.nextSibling)
  else if (pStyle) pPr.appendChild(numPr)
  else pPr.insertBefore(numPr, pPr.firstChild)
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
    insertPPrIntoStyle(styleEl, pPr)
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
