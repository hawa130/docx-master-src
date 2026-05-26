import { NS, type TableClassification } from "@lib/parse/types.ts"
import {
  descendantsNS,
  firstChildNS,
  getChildrenNS,
  textContent,
  wVal,
} from "@lib/xml/xml-utils.ts"

export interface TableSummary {
  classification: TableClassification
  rows: number
  cols: number
  headers: string[]
  firstRowLooksLikeHeader: boolean
}

/** Single cell holding more paragraphs than this is treated as a body
 * container and forces the table into the `layout` bucket. Calibrated
 * against survey of real fixtures: true data cells never exceed 3
 * paragraphs; multi-line form cells (申报书 checkbox lists, proposal
 * evaluation rubrics) sit at 5; layout content containers start at 4
 * but typically reach 10+. Threshold 5 keeps form / data cells stable
 * while catching content-heavy multi-tc layout cases that S1 misses. */
const BULK_CELL_PARA_THRESHOLD = 5

export function summarizeTable(tbl: Element): TableSummary {
  const rows = getChildrenNS(tbl, NS.w, "tr")
  const rowCount = rows.length
  let maxCols = 0
  let maxCellParas = 0
  const rowEffectiveCols: number[] = []
  const rowTcCounts: number[] = []

  for (const tr of rows) {
    const tcs = getChildrenNS(tr, NS.w, "tc")
    let cells = 0
    for (const tc of tcs) {
      const tcPr = firstChildNS(tc, NS.w, "tcPr")
      const gridSpan = tcPr ? firstChildNS(tcPr, NS.w, "gridSpan") : null
      const span = gridSpan ? parseInt(wVal(gridSpan) || "1", 10) : 1
      cells += span
      const cellParas = descendantsNS(tc, NS.w, "p").length
      if (cellParas > maxCellParas) maxCellParas = cellParas
    }
    rowEffectiveCols.push(cells)
    rowTcCounts.push(tcs.length)
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

  // Layout signals — overrule structural form/data shape:
  //   S1 single-cell-per-row stack: every row has exactly one <w:tc>
  //      (regardless of gridSpan) and the table holds >3 paragraphs.
  //      Catches both true 1-grid-column tables AND the common Word
  //      pattern "table grid is N-column but body rows all merge
  //      across all columns" — visually identical, both are stacks
  //      of full-width single cells used as content containers.
  //      Counting <w:tc> instead of gridSpan-effective cols is what
  //      makes the second form get recognized.
  //   S2 outlineLvl: any descendant <w:outlineLvl/> means the table
  //      contains heading-bound paragraphs. Data cells never carry
  //      outline level.
  //   S3 bulk cell: a single cell holding many paragraphs is a body
  //      container, not a tabular data cell.
  const allSingleTcPerRow = rowTcCounts.every((c) => c === 1)
  const hasOutlineHeading = descendantsNS(tbl, NS.w, "outlineLvl").length > 0

  // S0: 1×1 tables are always layout boxes. Structurally a data table
  // requires at least a header + one data row; 1×1 is by elimination a
  // single-cell content container (cover page title boxes, callout
  // frames, etc.). Falls through the normal classification path to
  // "data" otherwise, which hides the box's content.
  if (rowCount === 1 && maxCols === 1) {
    return { classification: "layout", rows: rowCount, cols: maxCols, headers, firstRowLooksLikeHeader: false }
  }

  let classification: TableClassification = "data"
  if (allSingleTcPerRow && totalParas > 3) {
    classification = "layout"
  } else if (hasOutlineHeading || maxCellParas > BULK_CELL_PARA_THRESHOLD) {
    classification = "layout"
  } else if (rowCount > 1 && maxCols > 1) {
    // First-row-looks-like-header is now a renderer hint, not a
    // classifier branch — both data and what was previously "form"
    // collapse to "data" since their downstream behavior is identical.
    classification = "data"
  }
  // else: default "data" from initialization

  return {
    classification,
    rows: rowCount,
    cols: maxCols,
    headers,
    firstRowLooksLikeHeader: rows.length > 0 ? looksLikeHeaderRow(rows[0]!) : false,
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
