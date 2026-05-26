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
  /** First-row cell text snippets, truncated. Empty array if the
   * first row has no cells or all cells are empty. Renderer uses
   * this as the source of truth for row1 display — agent decides
   * whether row1 is "headers" or "first data row" based on context. */
  row1Texts: string[]
  /** Short label naming which classifier signal fired. One of:
   *   - "singleTcStack" — every row has 1 <w:tc> AND totalParas > 3 (S1)
   *   - "outlineLvl"     — table contains direct <w:outlineLvl> (S2)
   *   - "bulkCell"       — some cell has > threshold paragraphs (S2)
   *   - "1x1"            — 1 row × 1 col, layout by elimination (S0)
   *   - "multiColData"   — rowCount > 1 && maxCols > 1, default data
   *   - "fallback"       — degenerate (single-cell tables that aren't
   *                        1x1 layouts, etc.) → data */
  classificationReason: string
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
    rowTcCounts.push(tcs.length)
    if (cells > maxCols) maxCols = cells
  }

  // collect row1 text snippets from first row
  const row1Texts: string[] = []
  if (rows.length > 0) {
    for (const tc of getChildrenNS(rows[0]!, NS.w, "tc")) {
      row1Texts.push(collectCellText(tc).trim())
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
    return { classification: "layout", rows: rowCount, cols: maxCols, row1Texts, classificationReason: "1x1" }
  }

  let classification: TableClassification = "data"
  let classificationReason = "fallback"

  if (allSingleTcPerRow && totalParas > 3) {
    classification = "layout"
    classificationReason = "singleTcStack"
  } else if (hasOutlineHeading) {
    classification = "layout"
    classificationReason = "outlineLvl"
  } else if (maxCellParas > BULK_CELL_PARA_THRESHOLD) {
    classification = "layout"
    classificationReason = "bulkCell"
  } else if (rowCount > 1 && maxCols > 1) {
    classification = "data"
    classificationReason = "multiColData"
  }
  // else: default "data" / "fallback"

  return {
    classification,
    rows: rowCount,
    cols: maxCols,
    row1Texts,
    classificationReason,
  }
}

function collectCellText(tc: Element): string {
  const ts = descendantsNS(tc, NS.w, "t")
  return ts.map((t) => textContent(t)).join("")
}
