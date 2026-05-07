import { DocxReader } from "@lib/reader.ts"
import { buildResolverContext, paragraphStyleId, paragraphText } from "@lib/locator.ts"
import { detectBlockers, explainBlockerReason, summarizeBlockers } from "@lib/blockers.ts"
import { DocumentParser } from "@lib/document-parser.ts"
import { StyleResolver } from "@lib/style-resolver.ts"

/**
 * `inspect_blockers` — list paragraphs that `apply_edits` will refuse to
 * touch. Causes: existing tracked changes, complex field regions
 * (TOC / STYLEREF / cross-references), SDT content controls.
 *
 * Use before composing `edits[]` so locators can be chosen to avoid
 * conflicts. Output is paragraph-level (cell paragraphs unindexed).
 */

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

    const out: string[] = []
    out.push(
      `=== Blocker scan: ${scan.byElement.size} paragraph(s) blocked ` +
        `(tracked: ${counts["tracked-change"]}, field: ${counts.field}, sdt: ${counts.sdt}) ===`,
    )
    if (scan.byElement.size === 0) {
      out.push("(none — every indexed paragraph is editable)")
      console.log(out.join("\n"))
      return
    }
    // Sort by paragraph index for indexed entries; group unindexed (cell)
    // entries at the end.
    const indexedRows: Array<{ idx: number; element: Element }> = []
    const unindexedRows: Element[] = []
    for (const [el] of scan.byElement) {
      const idx = ctx.indexByElement.get(el)
      if (idx !== undefined) indexedRows.push({ idx, element: el })
      else unindexedRows.push(el)
    }
    indexedRows.sort((a, b) => a.idx - b.idx)
    for (const row of indexedRows) {
      const reason = scan.byElement.get(row.element)!
      const text = paragraphText(row.element).slice(0, 40)
      const style = paragraphStyleId(row.element)
      out.push(`  #${pad(row.idx)} [${reason}] style=${style}  ${JSON.stringify(text)}`)
    }
    if (unindexedRows.length > 0) {
      out.push(
        `  (${unindexedRows.length} blocked cell paragraph(s) — addressable only via cell locator)`,
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

function pad(n: number): string {
  return String(n).padStart(3, " ")
}

void main()
