import { NS, type TableClassification } from "./types.ts"
import {
  descendantsNS,
  firstChildNS,
  getChildrenNS,
  textContent,
  wAttr,
  wVal,
} from "./xml-utils.ts"

export interface TableSummary {
  classification: TableClassification
  rows: number
  cols: number
  headers: string[]
}

export function summarizeTable(tbl: Element): TableSummary {
  const rows = getChildrenNS(tbl, NS.w, "tr")
  const rowCount = rows.length
  let maxCols = 0
  const rowEffectiveCols: number[] = []

  for (const tr of rows) {
    const tcs = getChildrenNS(tr, NS.w, "tc")
    let cells = 0
    for (const tc of tcs) {
      const tcPr = firstChildNS(tc, NS.w, "tcPr")
      const gridSpan = tcPr ? firstChildNS(tcPr, NS.w, "gridSpan") : null
      const span = gridSpan ? parseInt(wVal(gridSpan) || "1", 10) : 1
      cells += span
    }
    rowEffectiveCols.push(cells)
    if (cells > maxCols) maxCols = cells
  }

  // collect headers from first row
  const headers: string[] = []
  if (rows.length > 0) {
    const first = rows[0]!
    for (const tc of getChildrenNS(first, NS.w, "tc")) {
      const text = collectCellText(tc).trim()
      headers.push(text)
    }
  }

  // count total paragraphs across all cells
  const totalParas = descendantsNS(tbl, NS.w, "p").length

  let classification: TableClassification = "data"
  const allSingleCol = rowEffectiveCols.every((c) => c === 1)
  if (allSingleCol && totalParas > 3) {
    classification = "layout"
  } else if (rowCount > 1 && maxCols > 1) {
    if (looksLikeHeaderRow(rows[0]!)) {
      classification = "data"
    } else if (looksLikeForm(rows)) {
      classification = "form"
    } else {
      classification = "data"
    }
  } else {
    classification = "data"
  }

  return {
    classification,
    rows: rowCount,
    cols: maxCols,
    headers,
  }
}

function collectCellText(tc: Element): string {
  const ts = descendantsNS(tc, NS.w, "t")
  return ts.map((t) => textContent(t)).join("")
}

function looksLikeHeaderRow(tr: Element): boolean {
  const tcs = getChildrenNS(tr, NS.w, "tc")
  if (tcs.length === 0) return false
  let boldCells = 0
  let shortCells = 0
  let total = 0
  for (const tc of tcs) {
    total++
    const text = collectCellText(tc).trim()
    if (text.length > 0 && text.length <= 20) shortCells++
    // check any run rPr has <w:b/>
    const rPrs = descendantsNS(tc, NS.w, "rPr")
    let hasBold = false
    for (const rPr of rPrs) {
      const b = firstChildNS(rPr, NS.w, "b")
      if (b && wVal(b) !== "0") {
        hasBold = true
        break
      }
    }
    if (hasBold) boldCells++
  }
  const ratio = (boldCells + shortCells) / (total * 2)
  return boldCells >= Math.ceil(total / 2) || ratio > 0.6
}

function looksLikeForm(rows: Element[]): boolean {
  // label-value pattern: short text in left column(s), longer/empty in right
  let formish = 0
  for (const tr of rows) {
    const tcs = getChildrenNS(tr, NS.w, "tc")
    if (tcs.length < 2) continue
    const leftText = collectCellText(tcs[0]!).trim()
    const rightText = collectCellText(tcs[tcs.length - 1]!).trim()
    if (leftText.length > 0 && leftText.length <= 12) {
      if (rightText.length === 0 || rightText.length >= leftText.length) formish++
    }
  }
  return formish >= Math.ceil(rows.length / 2)
}
