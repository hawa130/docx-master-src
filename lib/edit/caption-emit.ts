/**
 * Caption-class layout emitters.
 *
 * EquationBlock with captionId → 3-col borderless table:
 *   [ left spacer | centered OMML | caption (prefix + STYLEREFs + SEQ + suffix) ]
 *
 * EquationBlock without captionId → single centered paragraph with OMML.
 *
 * CaptionBlock → single paragraph: prefix + STYLEREFs + SEQ + suffix +
 * bodySeparator + text (the last two omitted when text is empty).
 *
 * CaptionCounterReset → single paragraph with one hidden SEQ field
 * (`\r N \h`) — invisible counter advance marker.
 *
 * All caption-bearing emitters return both the XML root element AND a
 * PendingCaptionFill / PendingCaptionReset record. The apply pipeline
 * collects these, passes to the counter simulator, then backfills the
 * computed values into the result text elements.
 *
 * Bookmark scope (per spec §4.7): the primary bookmark wraps prefix run
 * through suffix run — number + decoration, NOT the bodySeparator + body
 * for CaptionBlock. Caller supplies the pre-allocated `{ id, name }` via
 * `bookmark` option; emit emits bookmarkStart/End around the correct
 * range. When `bookmark` is omitted, no bookmark is emitted.
 */

import { NS } from "@lib/parse/types.ts"
import { parseXml } from "@lib/xml/reader.ts"
import { getOmmlSync } from "@lib/edit/math/latex-to-omml.ts"
import { emitSeqField } from "@lib/edit/fields/seq-field.ts"
import { emitStyleRefField } from "@lib/edit/fields/styleref-field.ts"
import type {
  ResolvedCaptionConfig,
  PendingCaptionFill,
  PendingCaptionReset,
} from "@lib/edit/caption-counter.ts"

const w = NS.w
const m = NS.m
const XML_NS = "http://www.w3.org/XML/1998/namespace"

export type MathSource = { latex: string } | { omml: string }

export interface BookmarkRange {
  id: number
  name: string
}

/* ============================ EquationBlock ============================ */

export interface NumberedEquationOptions {
  mathSource: MathSource
  equationStyleId: string
  captionConfig: ResolvedCaptionConfig
  subGroup?: "start" | "continue"
  bookmark?: BookmarkRange
  /** Section's usable text width (twips). Mid-cell width derives from
   * this minus left + right spacers. Default 8500 (~5.9in, A4 with
   * narrow margins). */
  usableWidthTwips?: number
}

/** Emit a numbered equation as a 3-col borderless table. Middle cell
 * holds the OMML; right cell holds the caption (prefix + chapter +
 * counter + suffix), bookmark-wrapped per spec §4.7. */
export function emitNumberedEquation(
  ownerDoc: Document,
  opts: NumberedEquationOptions,
): { table: Element; fill: PendingCaptionFill } {
  const total = opts.usableWidthTwips ?? 8500
  const leftW = 850
  const rightW = 1500
  const midW = Math.max(2000, total - leftW - rightW)

  const leftCell = buildCell(ownerDoc, leftW, [emptyParagraph(ownerDoc)])
  const middleCell = buildCell(ownerDoc, midW, [
    centeredEquationParagraph(ownerDoc, opts.mathSource, opts.equationStyleId),
  ])
  const captionResult = buildCaptionParagraph(ownerDoc, {
    captionConfig: opts.captionConfig,
    subGroup: opts.subGroup,
    bookmark: opts.bookmark,
    body: undefined,
  })
  const rightCell = buildCell(ownerDoc, rightW, [captionResult.paragraph])

  const table = buildBorderlessTable(
    ownerDoc,
    [leftW, midW, rightW],
    [leftCell, middleCell, rightCell],
  )

  const fill: PendingCaptionFill = {
    paragraph: captionResult.paragraph,
    identifier: opts.captionConfig.identifier,
    subGroup: opts.subGroup,
    chapterPrefixResults: captionResult.chapterPrefixResults,
    parentSeqResult: captionResult.parentSeqResult,
    subSeqResult: captionResult.subSeqResult,
  }

  return { table, fill }
}

export interface UnnumberedEquationOptions {
  mathSource: MathSource
  equationStyleId: string
}

/** Emit an unnumbered equation as a single centered paragraph (no
 * surrounding table, no caption). */
export function emitUnnumberedEquation(
  ownerDoc: Document,
  opts: UnnumberedEquationOptions,
): Element {
  return centeredEquationParagraph(ownerDoc, opts.mathSource, opts.equationStyleId)
}

/* ============================= CaptionBlock ============================ */

export interface CaptionBlockOptions {
  captionConfig: ResolvedCaptionConfig
  text: string
  bookmark?: BookmarkRange
}

/** Emit a caption paragraph (FigureCaption / TableCaption / Theorem /
 * ...). Single paragraph: prefix + chapter STYLEREFs + SEQ + suffix +
 * bodySeparator + text. Last two omitted when text is empty. */
export function emitCaptionBlock(
  ownerDoc: Document,
  opts: CaptionBlockOptions,
): { paragraph: Element; fill: PendingCaptionFill } {
  const built = buildCaptionParagraph(ownerDoc, {
    captionConfig: opts.captionConfig,
    subGroup: undefined,
    bookmark: opts.bookmark,
    body: opts.text === "" ? undefined : opts.text,
  })

  const fill: PendingCaptionFill = {
    paragraph: built.paragraph,
    identifier: opts.captionConfig.identifier,
    subGroup: undefined,
    chapterPrefixResults: built.chapterPrefixResults,
    parentSeqResult: built.parentSeqResult,
    subSeqResult: built.subSeqResult,
  }

  return { paragraph: built.paragraph, fill }
}

/* ========================== CaptionCounterReset ======================== */

export interface CaptionResetOptions {
  identifier: string
  newValue: number
}

/** Emit a hidden SEQ marker paragraph that resets the identifier's
 * counter. Paragraph carries Normal style; SEQ field uses `\r N \h`. */
export function emitCaptionReset(
  ownerDoc: Document,
  opts: CaptionResetOptions,
): { paragraph: Element; reset: PendingCaptionReset } {
  const p = ownerDoc.createElementNS(w, "w:p")
  const pPr = ownerDoc.createElementNS(w, "w:pPr")
  const pStyle = ownerDoc.createElementNS(w, "w:pStyle")
  pStyle.setAttributeNS(w, "w:val", "Normal")
  pPr.appendChild(pStyle)
  p.appendChild(pPr)

  const { runs } = emitSeqField(ownerDoc, {
    identifier: opts.identifier,
    format: "arabic",
    resetTo: opts.newValue,
    hidden: true,
  })
  for (const r of runs) p.appendChild(r)

  const reset: PendingCaptionReset = {
    paragraph: p,
    identifier: opts.identifier,
    newValue: opts.newValue,
  }
  return { paragraph: p, reset }
}

/* ============================= shared helpers ========================== */

/** Build the caption paragraph's run sequence:
 *
 *   [bookmarkStart?]
 *   <r>prefix</r>
 *   [STYLEREF + literal separator] (per chapterPrefix entry)
 *   [SEQ identifier (with switches)]
 *   [SEQ identifierSub (when subGroup is start/continue)]
 *   <r>suffix</r>
 *   [bookmarkEnd?]
 *   [<r>bodySeparator</r><r>text</r>] (when body provided)
 *
 * Returns the paragraph + references to the field result <w:t> elements
 * for the counter simulator to populate. */
function buildCaptionParagraph(
  ownerDoc: Document,
  opts: {
    captionConfig: ResolvedCaptionConfig
    subGroup: "start" | "continue" | undefined
    bookmark: BookmarkRange | undefined
    body: string | undefined
  },
): {
  paragraph: Element
  chapterPrefixResults: Element[]
  parentSeqResult: Element
  subSeqResult: Element | undefined
} {
  const config = opts.captionConfig
  const p = ownerDoc.createElementNS(w, "w:p")
  const pPr = ownerDoc.createElementNS(w, "w:pPr")
  const pStyle = ownerDoc.createElementNS(w, "w:pStyle")
  pStyle.setAttributeNS(w, "w:val", config.paragraphStyleId)
  pPr.appendChild(pStyle)
  p.appendChild(pPr)

  if (opts.bookmark) {
    const bmStart = ownerDoc.createElementNS(w, "w:bookmarkStart")
    bmStart.setAttributeNS(w, "w:id", String(opts.bookmark.id))
    bmStart.setAttributeNS(w, "w:name", opts.bookmark.name)
    p.appendChild(bmStart)
  }

  if (config.prefix !== "") {
    p.appendChild(textRun(ownerDoc, config.prefix))
  }

  const chapterPrefixResults: Element[] = []
  for (let i = 0; i < config.chapterPrefix.length; i++) {
    const entry = config.chapterPrefix[i]!
    const { runs, resultTextEl } = emitStyleRefField(ownerDoc, {
      styleName: entry.styleName,
      switches: ["\\n"],
    })
    for (const r of runs) p.appendChild(r)
    chapterPrefixResults.push(resultTextEl)
    // separator after every chapter ref (including the last, joining to SEQ)
    p.appendChild(textRun(ownerDoc, config.chapterSeparator))
  }

  const parent = emitSeqField(ownerDoc, {
    identifier: config.identifier,
    format: config.format,
    restartAtOutlineLevel: config.restartAtOutlineLevel,
    repeat: opts.subGroup === "continue",
  })
  for (const r of parent.runs) p.appendChild(r)

  let subSeqResult: Element | undefined
  if (opts.subGroup !== undefined && config.subCounter) {
    if (config.subCounter.prefix !== "") {
      p.appendChild(textRun(ownerDoc, config.subCounter.prefix))
    }
    const sub = emitSeqField(ownerDoc, {
      identifier: `${config.identifier}Sub`,
      format: config.subCounter.format,
      resetTo: opts.subGroup === "start" ? 1 : undefined,
    })
    for (const r of sub.runs) p.appendChild(r)
    subSeqResult = sub.resultTextEl
    if (config.subCounter.suffix !== "") {
      p.appendChild(textRun(ownerDoc, config.subCounter.suffix))
    }
  }

  if (config.suffix !== "") {
    p.appendChild(textRun(ownerDoc, config.suffix))
  }

  if (opts.bookmark) {
    const bmEnd = ownerDoc.createElementNS(w, "w:bookmarkEnd")
    bmEnd.setAttributeNS(w, "w:id", String(opts.bookmark.id))
    p.appendChild(bmEnd)
  }

  if (opts.body !== undefined) {
    p.appendChild(textRun(ownerDoc, config.bodySeparator))
    p.appendChild(textRun(ownerDoc, opts.body))
  }

  return {
    paragraph: p,
    chapterPrefixResults,
    parentSeqResult: parent.resultTextEl,
    subSeqResult,
  }
}

function centeredEquationParagraph(
  ownerDoc: Document,
  source: MathSource,
  equationStyleId: string,
): Element {
  const p = ownerDoc.createElementNS(w, "w:p")
  const pPr = ownerDoc.createElementNS(w, "w:pPr")
  const pStyle = ownerDoc.createElementNS(w, "w:pStyle")
  pStyle.setAttributeNS(w, "w:val", equationStyleId)
  pPr.appendChild(pStyle)
  p.appendChild(pPr)

  const oMath = buildOMath(source, ownerDoc, true)
  const oMathPara = ownerDoc.createElementNS(m, "m:oMathPara")
  oMathPara.appendChild(oMath)
  p.appendChild(oMathPara)
  return p
}

function buildOMath(source: MathSource, ownerDoc: Document, displayMode: boolean): Element {
  const xml = "latex" in source ? getOmmlSync(source.latex, displayMode) : source.omml
  const parsed = parseXml(xml)
  const root = parsed.documentElement
  if (!root) {
    throw new Error(`Math source produced no root element: ${JSON.stringify(source)}`)
  }
  return ownerDoc.importNode(root, true) as Element
}

function textRun(ownerDoc: Document, text: string): Element {
  const r = ownerDoc.createElementNS(w, "w:r")
  const t = ownerDoc.createElementNS(w, "w:t")
  t.setAttributeNS(XML_NS, "xml:space", "preserve")
  t.textContent = text
  r.appendChild(t)
  return r
}

function emptyParagraph(ownerDoc: Document): Element {
  return ownerDoc.createElementNS(w, "w:p")
}

function buildCell(ownerDoc: Document, widthTwips: number, content: Element[]): Element {
  const tc = ownerDoc.createElementNS(w, "w:tc")
  const tcPr = ownerDoc.createElementNS(w, "w:tcPr")
  const tcW = ownerDoc.createElementNS(w, "w:tcW")
  tcW.setAttributeNS(w, "w:w", String(widthTwips))
  tcW.setAttributeNS(w, "w:type", "dxa")
  tcPr.appendChild(tcW)
  tc.appendChild(tcPr)
  for (const c of content) tc.appendChild(c)
  return tc
}

function buildBorderlessTable(ownerDoc: Document, colWidths: number[], cells: Element[]): Element {
  const tbl = ownerDoc.createElementNS(w, "w:tbl")

  const tblPr = ownerDoc.createElementNS(w, "w:tblPr")
  const tblW = ownerDoc.createElementNS(w, "w:tblW")
  tblW.setAttributeNS(w, "w:w", String(colWidths.reduce((a, b) => a + b, 0)))
  tblW.setAttributeNS(w, "w:type", "dxa")
  tblPr.appendChild(tblW)
  const borders = ownerDoc.createElementNS(w, "w:tblBorders")
  for (const edge of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    const b = ownerDoc.createElementNS(w, `w:${edge}`)
    b.setAttributeNS(w, "w:val", "none")
    b.setAttributeNS(w, "w:sz", "0")
    b.setAttributeNS(w, "w:color", "auto")
    borders.appendChild(b)
  }
  tblPr.appendChild(borders)
  const layout = ownerDoc.createElementNS(w, "w:tblLayout")
  layout.setAttributeNS(w, "w:type", "fixed")
  tblPr.appendChild(layout)
  tbl.appendChild(tblPr)

  const tblGrid = ownerDoc.createElementNS(w, "w:tblGrid")
  for (const width of colWidths) {
    const gc = ownerDoc.createElementNS(w, "w:gridCol")
    gc.setAttributeNS(w, "w:w", String(width))
    tblGrid.appendChild(gc)
  }
  tbl.appendChild(tblGrid)

  const tr = ownerDoc.createElementNS(w, "w:tr")
  for (const c of cells) tr.appendChild(c)
  tbl.appendChild(tr)

  return tbl
}
