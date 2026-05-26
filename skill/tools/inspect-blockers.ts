import { DocxReader } from "@lib/xml/reader.ts"
import { buildResolverContext, paragraphStyleId, paragraphText } from "@lib/edit/locator.ts"
import { detectBlockers, explainBlockerReason, summarizeBlockers } from "@lib/edit/blockers.ts"
import { DocumentParser } from "@lib/parse/document-parser.ts"
import { StyleResolver } from "@lib/parse/style-resolver.ts"
import { summarizeTable } from "@lib/parse/table-classifier.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS } from "@lib/xml/xml-utils.ts"
import { pad } from "@lib/parse/format.ts"
import type { CellCoords } from "@lib/edit/text-search.ts"

/**
 * `inspect_blockers` — list paragraphs that `apply`'s edit phase will refuse
 * to touch. Causes: existing tracked changes, complex field regions
 * (TOC / STYLEREF / cross-references), SDT content controls.
 *
 * Use before composing `edits[]` so locators can be chosen to avoid
 * conflicts. Indexed paragraphs appear as `#NNN`; paragraphs inside
 * data-table cells appear as `T<n>R<n>C<n> K<n>`.
 */

/** Build a Map from every <w:p> inside data-table cells to its cell coords.
 * Only tables classified as "data" are included; layout-table paragraphs
 * receive index assignments via walkIndexedParagraphs already. */
function buildDataCellIndex(body: Element): Map<Element, CellCoords> {
  const out = new Map<Element, CellCoords>()
  let tableIdx = 0
  for (const child of getChildren(body)) {
    if (child.namespaceURI !== NS.w || child.localName !== "tbl") continue
    tableIdx++
    const summary = summarizeTable(child)
    if (summary.classification !== "data") continue
    const rows = getChildrenNS(child, NS.w, "tr")
    for (let ri = 0; ri < rows.length; ri++) {
      const cells = getChildrenNS(rows[ri]!, NS.w, "tc")
      for (let ci = 0; ci < cells.length; ci++) {
        const paras = getChildrenNS(cells[ci]!, NS.w, "p")
        for (let pi = 0; pi < paras.length; pi++) {
          out.set(paras[pi]!, { table: tableIdx, row: ri + 1, col: ci + 1, paragraph: pi + 1 })
        }
      }
    }
  }
  return out
}

function formatCellCoord(c: CellCoords): string {
  return `T${c.table}R${c.row}C${c.col} K${c.paragraph}`
}

async function main() {
  const file = process.argv[2]
  if (!file) {
    console.error("Usage: node scripts/inspect_blockers.js <docx-path>")
    process.exit(1)
  }
  try {
    const reader = await DocxReader.open(file)
    const documentDoc = await reader.readXml("word/document.xml")
    if (!documentDoc) {
      console.error("word/document.xml not found")
      process.exit(1)
    }
    const stylesDoc = await reader.readXml("word/styles.xml")
    const themeDoc = await reader.readXml("word/theme/theme1.xml")
    const numberingDoc = await reader.readXml("word/numbering.xml")
    const resolver = new StyleResolver(stylesDoc, themeDoc)
    if (stylesDoc) resolver.expandThemedFontsInStyles(stylesDoc)
    const parsed = new DocumentParser(documentDoc, resolver, numberingDoc).parse()

    const ctx = buildResolverContext(documentDoc, parsed.paragraphs)
    const scan = detectBlockers(documentDoc, ctx.indexByElement)
    const counts = summarizeBlockers(scan)

    const body = firstChildNS(documentDoc.documentElement, NS.w, "body")
    const dataCellIndex = body ? buildDataCellIndex(body) : new Map<Element, CellCoords>()

    const out: string[] = []
    out.push(
      `=== Blocker scan: ${scan.byElement.size} paragraph(s) blocked ` +
        `(tracked: ${counts["tracked-change"]}, field: ${counts.field}, sdt: ${counts.sdt}) ===`,
    )
    if (scan.byElement.size === 0) {
      out.push("(none — every paragraph is editable)")
      console.log(out.join("\n"))
      return
    }
    // Sort indexed entries by paragraph number; cell entries by coord tuple.
    const indexedRows: Array<{ idx: number; element: Element }> = []
    const cellRows: Array<{ coords: CellCoords; element: Element }> = []
    const unknownRows: Element[] = []
    for (const [el] of scan.byElement) {
      const idx = ctx.indexByElement.get(el)
      if (idx !== undefined) {
        indexedRows.push({ idx, element: el })
      } else {
        const coords = dataCellIndex.get(el)
        if (coords) cellRows.push({ coords, element: el })
        else unknownRows.push(el)
      }
    }
    indexedRows.sort((a, b) => a.idx - b.idx)
    cellRows.sort((a, b) => {
      const c = a.coords
      const d = b.coords
      return c.table !== d.table
        ? c.table - d.table
        : c.row !== d.row
          ? c.row - d.row
          : c.col !== d.col
            ? c.col - d.col
            : c.paragraph - d.paragraph
    })

    for (const row of indexedRows) {
      const reason = scan.byElement.get(row.element)!
      const text = paragraphText(row.element).slice(0, 40)
      const style = paragraphStyleId(row.element)
      out.push(`  #${pad(row.idx)} [${reason}] style=${style}  ${JSON.stringify(text)}`)
    }
    for (const row of cellRows) {
      const reason = scan.byElement.get(row.element)!
      const text = paragraphText(row.element).slice(0, 40)
      const style = paragraphStyleId(row.element)
      out.push(
        `  ${formatCellCoord(row.coords).padEnd(14)} [${reason}] style=${style}  ${JSON.stringify(text)}`,
      )
    }
    if (unknownRows.length > 0) {
      out.push(
        `  (${unknownRows.length} blocked paragraph(s) at unresolved locations — likely nested inside an uncommon container (deeply nested tables, etc.). Not directly addressable via edit locators.)`,
      )
    }
    out.push("")
    out.push("Reasons:")
    for (const reason of new Set(scan.byElement.values())) {
      out.push(`  ${reason}: ${explainBlockerReason(reason)}`)
    }
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

void main()
