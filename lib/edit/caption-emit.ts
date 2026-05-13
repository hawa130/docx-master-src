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
import { buildPlainTextRun } from "@lib/xml/xml-utils.ts"
import { getOmmlSync } from "@lib/edit/math/latex-to-omml.ts"
import { emitSeqField } from "@lib/edit/fields/seq-field.ts"
import { emitStyleRefField } from "@lib/edit/fields/styleref-field.ts"
import type { BookmarkRange } from "@lib/edit/bookmark.ts"
import type {
  ResolvedCaptionConfig,
  PendingCaptionFill,
  PendingCaptionReset,
} from "@lib/edit/caption-counter.ts"

const w = NS.w
const m = NS.m

export type { BookmarkRange }
export type MathSource = { latex: string } | { omml: string }

/** SEQ identifier for the hidden chapter counter paired with a heading
 * style under chapterPrefix `format` override. Engine reserves the
 * `_chap_` prefix; agent-declared captionIds matching this pattern
 * throw at schema validation. */
export function chapterCounterIdentifier(styleId: string): string {
  return `_chap_${styleId}`
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
  /** Subequation membership for caption paragraphs that are themselves
   * subequation children. Almost always undefined on CaptionBlock —
   * only standardize re-emit uses this to preserve subGroup state when
   * rebuilding existing caption paragraphs. */
  subGroup?: "start" | "continue"
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
    subGroup: opts.subGroup,
    bookmark: opts.bookmark,
    body: opts.text === "" ? undefined : opts.text,
  })

  const fill: PendingCaptionFill = {
    paragraph: built.paragraph,
    identifier: opts.captionConfig.identifier,
    subGroup: opts.subGroup,
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
 * counter so the NEXT caption emits the agent's specified `newValue`.
 * Word's `SEQ \r N` resets the counter to N AND the marker field itself
 * emits N (hidden); the next visible SEQ increments to N+1. To match
 * the agent's intent ("the next caption is newValue"), emit `\r
 * (newValue - 1)`. Counter sim mirrors the same convention. */
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

  // Drop SEQ `\h` switch — Word's documented quirk is that `\h` does
  // NOT hide the result when a `\*` format switch is also present (and
  // we always emit `\* ARABIC`). Wrap each run's rPr in `<w:vanish/>`
  // for character-level hiding instead — the SEQ counter still
  // advances (Word evaluates the field), but no rendered artifact
  // appears in the paragraph.
  const { runs } = emitSeqField(ownerDoc, {
    identifier: opts.identifier,
    format: "arabic",
    resetTo: opts.newValue - 1,
  })
  for (const r of runs) {
    const rPr = ownerDoc.createElementNS(w, "w:rPr")
    rPr.appendChild(ownerDoc.createElementNS(w, "w:vanish"))
    r.insertBefore(rPr, r.firstChild)
    p.appendChild(r)
  }

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
export interface CaptionRunSequenceOptions {
  captionConfig: ResolvedCaptionConfig
  subGroup: "start" | "continue" | undefined
  bookmark: BookmarkRange | undefined
  body: string | undefined
}

export interface CaptionRunSequence {
  /** Ordered runs to append after `<w:pPr>` in the caption paragraph:
   * bookmarkStart? + prefix + STYLEREF (×N) + separator + SEQ + sub-SEQ?
   * + suffix + bookmarkEnd? + bodySeparator? + body? */
  runs: Element[]
  chapterPrefixResults: Element[]
  parentSeqResult: Element
  subSeqResult: Element | undefined
}

/** Build the run sequence that goes inside a caption paragraph (no
 * wrapping `<w:p>`, no `<w:pPr>`). Shared by `buildCaptionParagraph`
 * (which wraps in a fresh paragraph) and `standardize-captions` (which
 * appends into an existing paragraph after preserving its pPr). */
export function buildCaptionRunSequence(
  ownerDoc: Document,
  opts: CaptionRunSequenceOptions,
): CaptionRunSequence {
  const config = opts.captionConfig
  const runs: Element[] = []

  if (opts.bookmark) {
    const bmStart = ownerDoc.createElementNS(w, "w:bookmarkStart")
    bmStart.setAttributeNS(w, "w:id", String(opts.bookmark.id))
    bmStart.setAttributeNS(w, "w:name", opts.bookmark.name)
    runs.push(bmStart)
  }

  if (config.prefix !== "") {
    runs.push(buildPlainTextRun(ownerDoc, config.prefix))
  }

  const chapterPrefixResults: Element[] = []
  for (const entry of config.chapterPrefix) {
    let resultTextEl: Element
    if (entry.format !== undefined) {
      // Format override path: reference a hidden auto-chapter SEQ
      // counter that the apply pipeline injects into each heading
      // paragraph of the cited style. `\c` repeats the counter's
      // current value (no increment), so the caption reads whatever
      // the most recent heading set. Identifier convention:
      // `_chap_<styleId>` — engine-reserved, schema rejects agent use.
      //
      // Word `STYLEREF "Heading 1" \n \* ARABIC` does NOT reliably
      // re-format chineseCounting (or other non-Arabic) source numFmts
      // — `\n` returns the heading's full rendered lvlText ("第一章")
      // and `\* ARABIC` doesn't extract the numeric portion. The
      // hidden-SEQ approach sidesteps that by maintaining a parallel
      // Arabic counter in the chosen format.
      const { runs: seqRuns, resultTextEl: t } = emitSeqField(ownerDoc, {
        identifier: chapterCounterIdentifier(entry.styleId),
        format: entry.format,
        repeat: true,
      })
      for (const r of seqRuns) runs.push(r)
      resultTextEl = t
    } else {
      // No override: STYLEREF \n returns the heading's native paragraph
      // number rendering. Word's F9 re-resolves on H1 changes.
      const { runs: styleRefRuns, resultTextEl: t } = emitStyleRefField(ownerDoc, {
        styleName: entry.styleName,
        switches: ["\\n"],
      })
      for (const r of styleRefRuns) runs.push(r)
      resultTextEl = t
    }
    chapterPrefixResults.push(resultTextEl)
    // separator after every chapter ref (including the last, joining to SEQ)
    runs.push(buildPlainTextRun(ownerDoc, config.chapterSeparator))
  }

  const parent = emitSeqField(ownerDoc, {
    identifier: config.identifier,
    format: config.format,
    restartAtOutlineLevel: config.restartAtOutlineLevel,
    repeat: opts.subGroup === "continue",
  })
  for (const r of parent.runs) runs.push(r)

  let subSeqResult: Element | undefined
  if (opts.subGroup !== undefined && config.subCounter) {
    if (config.subCounter.prefix !== "") {
      runs.push(buildPlainTextRun(ownerDoc, config.subCounter.prefix))
    }
    const sub = emitSeqField(ownerDoc, {
      identifier: `${config.identifier}Sub`,
      format: config.subCounter.format,
      resetTo: opts.subGroup === "start" ? 1 : undefined,
    })
    for (const r of sub.runs) runs.push(r)
    subSeqResult = sub.resultTextEl
    if (config.subCounter.suffix !== "") {
      runs.push(buildPlainTextRun(ownerDoc, config.subCounter.suffix))
    }
  }

  if (config.suffix !== "") {
    runs.push(buildPlainTextRun(ownerDoc, config.suffix))
  }

  if (opts.bookmark) {
    const bmEnd = ownerDoc.createElementNS(w, "w:bookmarkEnd")
    bmEnd.setAttributeNS(w, "w:id", String(opts.bookmark.id))
    runs.push(bmEnd)
  }

  if (opts.body !== undefined) {
    runs.push(buildPlainTextRun(ownerDoc, config.bodySeparator))
    runs.push(buildPlainTextRun(ownerDoc, opts.body))
  }

  return {
    runs,
    chapterPrefixResults,
    parentSeqResult: parent.resultTextEl,
    subSeqResult,
  }
}

function buildCaptionParagraph(
  ownerDoc: Document,
  opts: CaptionRunSequenceOptions,
): {
  paragraph: Element
  chapterPrefixResults: Element[]
  parentSeqResult: Element
  subSeqResult: Element | undefined
} {
  const p = ownerDoc.createElementNS(w, "w:p")
  const pPr = ownerDoc.createElementNS(w, "w:pPr")
  const pStyle = ownerDoc.createElementNS(w, "w:pStyle")
  pStyle.setAttributeNS(w, "w:val", opts.captionConfig.paragraphStyleId)
  pPr.appendChild(pStyle)
  p.appendChild(pPr)
  const seq = buildCaptionRunSequence(ownerDoc, opts)
  for (const r of seq.runs) p.appendChild(r)
  return {
    paragraph: p,
    chapterPrefixResults: seq.chapterPrefixResults,
    parentSeqResult: seq.parentSeqResult,
    subSeqResult: seq.subSeqResult,
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
