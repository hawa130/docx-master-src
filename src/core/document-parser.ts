import {
  NS,
  type ComputedParaStyle,
  type ComputedRunStyle,
  type DocumentElement,
  type NeighborItem,
  type ParsedParagraph,
  type SectionInfo,
  type TableClassification,
} from "./types.ts"
import {
  descendantsNS,
  firstChildNS,
  getChildren,
  getChildrenNS,
  textContent,
  wAttr,
  wVal,
} from "./xml-utils.ts"
import type { StyleResolver } from "./style-resolver.ts"
import { summarizeTable } from "./table-classifier.ts"

interface ParseResult {
  paragraphs: ParsedParagraph[]
  elements: DocumentElement[]
  sections: SectionInfo[]
  /** Flat ordered stream for inspect_neighbors. Empty paragraphs not collapsed. */
  neighborItems: NeighborItem[]
}

/** Internal item produced while walking — fingerprint not yet assigned. */
type RawItem =
  | { kind: "para"; paragraph: ParsedParagraph; isImageOnly: boolean; isEquationOnly: boolean }
  | {
      kind: "table"
      classification: TableClassification
      rows: number
      cols: number
      headers: string[]
      sectionIndex: number
      // for layout tables: include paragraphs (as raw items)
      innerParagraphs?: ParsedParagraph[]
    }
  | { kind: "image"; widthCm: number; heightCm: number; sectionIndex: number; paraIndex?: number }
  | { kind: "equation"; sectionIndex: number; paraIndex?: number }
  | { kind: "pageBreak"; sectionIndex: number }
  | { kind: "sectionBreak"; sectionIndex: number }

export class DocumentParser {
  private docXml: Document
  private resolver: StyleResolver
  private numberingDoc: Document | null
  private headerDocs: Map<string, Document> // rId → header doc
  private footerDocs: Map<string, Document> // rId → footer doc
  private rels: Map<string, { type: string; target: string }> // rId → rel info
  private nextParaIndex = 1
  private currentSection = 0

  constructor(
    docXml: Document,
    resolver: StyleResolver,
    numberingDoc: Document | null,
    options?: {
      headerDocs?: Map<string, Document>
      footerDocs?: Map<string, Document>
      rels?: Map<string, { type: string; target: string }>
    },
  ) {
    this.docXml = docXml
    this.resolver = resolver
    this.numberingDoc = numberingDoc
    this.headerDocs = options?.headerDocs ?? new Map()
    this.footerDocs = options?.footerDocs ?? new Map()
    this.rels = options?.rels ?? new Map()
  }

  parse(): ParseResult {
    const root = this.docXml.documentElement
    const body = firstChildNS(root, NS.w, "body")
    if (!body) {
      return { paragraphs: [], elements: [], sections: [] }
    }

    const allParagraphs: ParsedParagraph[] = []
    const flatItems: RawItem[] = []
    const sections: SectionInfo[] = []

    // section start tracker
    let sectionStartIndex = 1

    const children = getChildren(body)
    for (let idx = 0; idx < children.length; idx++) {
      const child = children[idx]!
      if (child.namespaceURI !== NS.w) continue

      if (child.localName === "p") {
        this.handleParagraph(child, allParagraphs, flatItems, null)
        // check for sectPr inside this paragraph
        const pPr = firstChildNS(child, NS.w, "pPr")
        const sectPr = pPr ? firstChildNS(pPr, NS.w, "sectPr") : null
        if (sectPr) {
          const endIndex = this.nextParaIndex - 1
          sections.push(
            this.buildSectionInfo(sectPr, sections.length, sectionStartIndex, endIndex),
          )
          sectionStartIndex = this.nextParaIndex
          flatItems.push({ kind: "sectionBreak", sectionIndex: this.currentSection })
          this.currentSection++
        }
      } else if (child.localName === "tbl") {
        const summary = summarizeTable(child)
        if (summary.classification === "layout") {
          // expand: walk all <w:p> inside in document order
          const innerParas: ParsedParagraph[] = []
          flatItems.push({
            kind: "table",
            classification: "layout",
            rows: summary.rows,
            cols: summary.cols,
            headers: summary.headers,
            sectionIndex: this.currentSection,
            innerParagraphs: innerParas,
          })
          // walk rows → cells → tcContent in order
          for (const tr of getChildrenNS(child, NS.w, "tr")) {
            for (const tc of getChildrenNS(tr, NS.w, "tc")) {
              this.walkLayoutCell(tc, allParagraphs, flatItems, innerParas)
            }
          }
        } else {
          flatItems.push({
            kind: "table",
            classification: summary.classification,
            rows: summary.rows,
            cols: summary.cols,
            headers: summary.headers,
            sectionIndex: this.currentSection,
          })
        }
      } else if (child.localName === "sectPr") {
        // final section
        const endIndex = this.nextParaIndex - 1
        sections.push(
          this.buildSectionInfo(child, sections.length, sectionStartIndex, endIndex),
        )
        sectionStartIndex = this.nextParaIndex
      }
    }

    // If no final sectPr was emitted but we had paragraphs, ensure at least one section exists
    if (sections.length === 0 && allParagraphs.length > 0) {
      sections.push({
        index: 0,
        paraRange: [1, allParagraphs.length],
        pageSize: { width: 0, height: 0 },
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        orientation: "portrait",
        header: null,
        footer: null,
        headerHasImage: false,
      })
    }

    // Build the flat NeighborItem stream for inspect_neighbors. Done in a
    // single pass over flatItems; empty paragraphs are kept individually
    // (unlike compressElements). Layout-table inner paragraphs are inlined
    // into the stream alongside top-level paragraphs.
    const neighborItems = this.buildNeighborItems(flatItems)

    // Convert flat items → DocumentElement[] with empty-paragraph compression
    const elements = this.compressElements(flatItems)

    return {
      paragraphs: allParagraphs,
      elements,
      sections,
      neighborItems,
    }
  }

  private walkLayoutCell(
    tc: Element,
    all: ParsedParagraph[],
    flat: RawItem[],
    innerParas: ParsedParagraph[],
  ): void {
    for (const c of getChildren(tc)) {
      if (c.namespaceURI !== NS.w) continue
      if (c.localName === "p") {
        this.handleParagraph(c, all, flat, "layout", innerParas)
      } else if (c.localName === "tbl") {
        // nested table; classify and process
        const summary = summarizeTable(c)
        if (summary.classification === "layout") {
          flat.push({
            kind: "table",
            classification: "layout",
            rows: summary.rows,
            cols: summary.cols,
            headers: summary.headers,
            sectionIndex: this.currentSection,
            innerParagraphs: innerParas,
          })
          for (const tr of getChildrenNS(c, NS.w, "tr")) {
            for (const tc2 of getChildrenNS(tr, NS.w, "tc")) {
              this.walkLayoutCell(tc2, all, flat, innerParas)
            }
          }
        } else {
          flat.push({
            kind: "table",
            classification: summary.classification,
            rows: summary.rows,
            cols: summary.cols,
            headers: summary.headers,
            sectionIndex: this.currentSection,
          })
        }
      }
    }
  }

  private handleParagraph(
    pEl: Element,
    all: ParsedParagraph[],
    flat: RawItem[],
    insideTable: TableClassification | null,
    innerCollect?: ParsedParagraph[],
  ): void {
    const pPr = firstChildNS(pEl, NS.w, "pPr")
    const pStyleEl = pPr ? firstChildNS(pPr, NS.w, "pStyle") : null
    const styleId = (pStyleEl && wVal(pStyleEl)) || "Normal"

    // dominant run rPr — pick the run with the most visible text characters.
    // Fallback to first non-empty run, then paragraph-mark rPr in pPr.
    // Why dominant rather than first: a heading may begin with a stylistically
    // distinct number prefix run ("1. " in blue) before its bold black title;
    // a list item may begin with a bold lead phrase before non-bold body text.
    // Taking the first run misclassifies the paragraph by its prefix.
    const runs = getChildrenNS(pEl, NS.w, "r")
    let dominantRun: Element | null = null
    let dominantLen = -1
    let firstNonEmptyRun: Element | null = null
    let firstSubstantiveRun: Element | null = null
    for (const r of runs) {
      let text = ""
      for (const c of getChildren(r)) {
        if (c.namespaceURI === NS.w && c.localName === "t") text += textContent(c)
      }
      if (text.length > 0 && firstNonEmptyRun === null) firstNonEmptyRun = r
      // Skip runs that look like a numbering / bullet prefix: pure digits,
      // CJK numerals, list punctuation, or whitespace. Otherwise a heading
      // "4.5 导出器" (prefix 4 chars, title 3 chars) would be classified by
      // its prefix run instead of the title.
      const isNumberingPrefix = text.length > 0 &&
        /^[\d一二三四五六七八九十百千零０-９\.\(\)（）【】［］〔〕\[\]•·○●◆◇■□★☆※\-、，,\s]+$/.test(text)
      if (!isNumberingPrefix && text.length > 0 && firstSubstantiveRun === null) {
        firstSubstantiveRun = r
      }
      const len = isNumberingPrefix ? 0 : text.length
      // Use >= so later runs win on ties: a heading "1.1 标题" typically has a
      // numbering-prefix run followed by a title run; the title run carries
      // the role-defining formatting.
      if (len >= dominantLen) {
        dominantLen = len
        dominantRun = r
      }
    }
    const chosenRun =
      dominantLen > 0
        ? dominantRun
        : firstSubstantiveRun || firstNonEmptyRun || runs[0] || null
    const chosenRPr = chosenRun ? firstChildNS(chosenRun, NS.w, "rPr") : null
    const paraMarkRPr = pPr ? firstChildNS(pPr, NS.w, "rPr") : null

    const computed = this.resolver.computeRunStyle(
      styleId,
      chosenRPr || paraMarkRPr,
    )

    // overlay paragraph-level pPr direct formatting onto computed pPr
    const directPPr = this.resolver.parsePPr(pPr)
    const finalPPr: ComputedParaStyle = { ...computed.pPr }
    for (const k of Object.keys(directPPr) as (keyof ComputedParaStyle)[]) {
      const v = directPPr[k]
      if (v !== undefined) (finalPPr as any)[k] = v
    }

    // text
    let text = ""
    let hasDrawing = false
    let hasEquation = false
    let imageDims: { widthCm: number; heightCm: number } | null = null
    let hasPageBreak = false
    for (const r of runs) {
      for (const c of getChildren(r)) {
        if (c.namespaceURI === NS.w) {
          if (c.localName === "t") text += textContent(c)
          else if (c.localName === "tab") text += "\t"
          else if (c.localName === "br") {
            const t = wAttr(c, "type")
            if (t === "page") hasPageBreak = true
            else text += "\n"
          } else if (c.localName === "drawing") {
            hasDrawing = true
            const dims = extractDrawingDims(c)
            if (dims) imageDims = dims
          }
        } else if (c.namespaceURI === NS.m) {
          if (c.localName === "oMath" || c.localName === "oMathPara") hasEquation = true
        }
      }
      // also drawings can be direct children of run — covered above
    }
    // also math elements may be direct children of paragraph, not inside a run
    for (const c of getChildren(pEl)) {
      if (c.namespaceURI === NS.m) {
        if (c.localName === "oMath" || c.localName === "oMathPara") hasEquation = true
      }
    }

    this.resolver.incrementUsage(styleId)
    const styleDef = this.resolver.getStyleDefinition(styleId)
    const styleName = styleDef?.name || computed.styleName || styleId

    const para: ParsedParagraph = {
      index: this.nextParaIndex++,
      text,
      rPr: computed.rPr,
      pPr: finalPPr,
      styleId,
      styleName,
      fingerprint: "", // assigned later
      context: {
        insideTable,
        sectionIndex: this.currentSection,
      },
    }
    all.push(para)
    if (innerCollect) innerCollect.push(para)
    const isImageOnly = hasDrawing && text.trim().length === 0
    const isEquationOnly = hasEquation && text.trim().length === 0

    // Emit image element if image-bearing
    if (hasDrawing && imageDims) {
      flat.push({
        kind: "image",
        widthCm: imageDims.widthCm,
        heightCm: imageDims.heightCm,
        sectionIndex: this.currentSection,
        paraIndex: para.index,
      })
    }
    if (hasEquation) {
      flat.push({ kind: "equation", sectionIndex: this.currentSection, paraIndex: para.index })
    }
    if (hasPageBreak) {
      flat.push({ kind: "pageBreak", sectionIndex: this.currentSection })
    }
    flat.push({ kind: "para", paragraph: para, isImageOnly, isEquationOnly })
  }

  private buildSectionInfo(
    sectPr: Element,
    index: number,
    startIdx: number,
    endIdx: number,
  ): SectionInfo {
    const pgSz = firstChildNS(sectPr, NS.w, "pgSz")
    const pgMar = firstChildNS(sectPr, NS.w, "pgMar")
    const width = pgSz ? parseInt(wAttr(pgSz, "w") || "0", 10) : 0
    const height = pgSz ? parseInt(wAttr(pgSz, "h") || "0", 10) : 0
    const orient = pgSz ? wAttr(pgSz, "orient") : null
    const orientation: "portrait" | "landscape" =
      orient === "landscape" ? "landscape" : "portrait"
    const margins = {
      top: pgMar ? parseInt(wAttr(pgMar, "top") || "0", 10) : 0,
      bottom: pgMar ? parseInt(wAttr(pgMar, "bottom") || "0", 10) : 0,
      left: pgMar ? parseInt(wAttr(pgMar, "left") || "0", 10) : 0,
      right: pgMar ? parseInt(wAttr(pgMar, "right") || "0", 10) : 0,
    }

    let header: string | null = null
    let footer: string | null = null
    let headerHasImage = false
    let footerPageNumFormat: string | undefined

    for (const ref of getChildrenNS(sectPr, NS.w, "headerReference")) {
      const rId = wAttr(ref, "id")
      const type = wAttr(ref, "type")
      if (rId && (type === "default" || type === null)) {
        const rel = this.rels.get(rId)
        const doc = rel ? this.headerDocs.get(rel.target) : null
        if (doc) {
          const text = collectAllText(doc.documentElement!).trim()
          header = text.length > 0 ? text : "(empty)"
          headerHasImage = descendantsNS(doc.documentElement!, NS.w, "drawing").length > 0
        }
      }
    }
    for (const ref of getChildrenNS(sectPr, NS.w, "footerReference")) {
      const rId = wAttr(ref, "id")
      const type = wAttr(ref, "type")
      if (rId && (type === "default" || type === null)) {
        const rel = this.rels.get(rId)
        const doc = rel ? this.footerDocs.get(rel.target) : null
        if (doc) {
          const text = collectAllText(doc.documentElement!).trim()
          footer = text.length > 0 ? text : "(empty)"
          // detect page number field format (PAGE field)
          const instrTexts = descendantsNS(doc.documentElement!, NS.w, "instrText")
            .map((e) => textContent(e))
            .join(" ")
          if (/PAGE/i.test(instrTexts)) {
            const m = instrTexts.match(/\\\*\s*([A-Za-z]+)/)
            footerPageNumFormat = m ? m[1] : "decimal"
          }
        }
      }
    }
    const pgNumType = firstChildNS(sectPr, NS.w, "pgNumType")
    if (pgNumType) {
      const fmt = wAttr(pgNumType, "fmt")
      if (fmt) footerPageNumFormat = fmt
    }

    return {
      index,
      paraRange: [startIdx, endIdx],
      pageSize: { width, height },
      margins,
      orientation,
      header,
      footer,
      headerHasImage,
      footerPageNumFormat,
    }
  }

  /**
   * Flatten the internal RawItem list into the public NeighborItem stream.
   * Each empty paragraph stays as its own entry (unlike compressElements).
   * inspect_neighbors uses this to compute on-demand neighbor windows.
   */
  private buildNeighborItems(items: RawItem[]): NeighborItem[] {
    const out: NeighborItem[] = []
    for (const it of items) {
      switch (it.kind) {
        case "para": {
          const text = it.paragraph.text
          const isEmpty =
            text.trim().length === 0 && !it.isImageOnly && !it.isEquationOnly
          out.push({
            kind: "paragraph",
            paraIndex: it.paragraph.index,
            isEmpty,
            sectionIndex: it.paragraph.context.sectionIndex,
          })
          break
        }
        case "image":
          out.push({
            kind: "image",
            widthCm: it.widthCm,
            heightCm: it.heightCm,
            sectionIndex: it.sectionIndex,
          })
          break
        case "table":
          out.push({
            kind: "table",
            classification: it.classification,
            rows: it.rows,
            cols: it.cols,
            sectionIndex: it.sectionIndex,
          })
          break
        case "equation":
          out.push({ kind: "equation", sectionIndex: it.sectionIndex })
          break
        case "pageBreak":
          out.push({ kind: "pageBreak", sectionIndex: it.sectionIndex })
          break
        case "sectionBreak":
          out.push({ kind: "sectionBreak", sectionIndex: it.sectionIndex })
          break
      }
    }
    return out
  }

  private compressElements(items: RawItem[]): DocumentElement[] {
    const out: DocumentElement[] = []
    let emptyRunCount = 0
    let emptyRunFirstIdx = 0
    let emptyRunSection = 0

    const flushEmpty = () => {
      if (emptyRunCount > 0) {
        out.push({
          kind: "emptyRun",
          count: emptyRunCount,
          firstIndex: emptyRunFirstIdx,
          sectionIndex: emptyRunSection,
        })
        emptyRunCount = 0
      }
    }

    for (const it of items) {
      if (it.kind === "para") {
        // Layout-table inner paragraphs are rendered inside the table block; skip top-level emission
        if (it.paragraph.context.insideTable === "layout") continue
        const isEmpty =
          it.paragraph.text.trim().length === 0 && !it.isImageOnly && !it.isEquationOnly
        if (isEmpty) {
          if (emptyRunCount === 0) {
            emptyRunFirstIdx = it.paragraph.index
            emptyRunSection = it.paragraph.context.sectionIndex
          }
          emptyRunCount++
          continue
        }
        flushEmpty()
        out.push({ kind: "paragraph", paragraph: it.paragraph })
      } else if (it.kind === "table") {
        flushEmpty()
        out.push({
          kind: "table",
          classification: it.classification,
          rows: it.rows,
          cols: it.cols,
          headers: it.headers,
          sectionIndex: it.sectionIndex,
          paragraphs: it.innerParagraphs ?? [],
        })
      } else if (it.kind === "image") {
        flushEmpty()
        out.push({
          kind: "image",
          widthCm: it.widthCm,
          heightCm: it.heightCm,
          sectionIndex: it.sectionIndex,
        })
      } else if (it.kind === "equation") {
        flushEmpty()
        out.push({ kind: "equation", sectionIndex: it.sectionIndex })
      } else if (it.kind === "pageBreak") {
        flushEmpty()
        out.push({ kind: "pageBreak", sectionIndex: it.sectionIndex })
      } else if (it.kind === "sectionBreak") {
        flushEmpty()
        out.push({ kind: "sectionBreak", sectionIndex: it.sectionIndex })
      }
    }
    flushEmpty()
    return out
  }
}

function extractDrawingDims(drawingEl: Element): { widthCm: number; heightCm: number } | null {
  const extents = descendantsNS(drawingEl, NS.wp, "extent")
  for (const e of extents) {
    const cx = parseInt(attrLocal(e, "cx") || "0", 10)
    const cy = parseInt(attrLocal(e, "cy") || "0", 10)
    if (cx > 0 && cy > 0) {
      // EMU → cm: 1cm = 360000 EMU
      return { widthCm: cx / 360000, heightCm: cy / 360000 }
    }
  }
  return null
}

function attrLocal(el: Element, name: string): string | null {
  return el.getAttribute(name)
}

function collectAllText(el: Element): string {
  const ts = descendantsNS(el, NS.w, "t")
  return ts.map((t) => textContent(t)).join("")
}
