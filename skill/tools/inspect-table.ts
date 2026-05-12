import { DocxReader } from "@lib/xml/reader.ts"
import { paragraphText, walkIndexedParagraphs, walkTopLevelTables } from "@lib/edit/locator.ts"
import { NS } from "@lib/parse/types.ts"
import { getChildrenNS } from "@lib/xml/xml-utils.ts"
import { summarizeTable } from "@lib/parse/table-classifier.ts"

/**
 * `inspect_table` — list top-level tables with cell text snippets.
 *
 * Output is the cell coordinate space the `cell` locator addresses:
 *   table T row R col C  → text snippet (first 40 chars)
 *
 * Used before composing a `cell` locator. Lists every top-level table —
 * data, form, layout — because cell locators can address any of them
 * (paragraph indices skip data/form, but cell locators don't).
 */

async function main() {
  const file = process.argv[2]
  if (!file) {
    console.error("Usage: node scripts/inspect_table.js <docx-path>")
    process.exit(1)
  }
  try {
    const reader = await DocxReader.open(file)
    const documentDoc = await reader.readXml("word/document.xml")
    if (!documentDoc) {
      console.error("word/document.xml not found")
      process.exit(1)
    }
    const tables = walkTopLevelTables(documentDoc)
    if (tables.length === 0) {
      console.log("(no top-level tables)")
      return
    }
    // Build a paragraph-element → index map so each cell can report which
    // #NNN range its paragraphs occupy. Only layout-table paragraphs are
    // indexed; data/form-table cells return empty arrays here (cell
    // locator addresses them anyway, no index needed).
    const indexByElement = new Map<Element, number>()
    for (const p of walkIndexedParagraphs(documentDoc)) {
      indexByElement.set(p.element, p.index)
    }
    const out: string[] = []
    for (const t of tables) {
      const summary = summarizeTable(t.element)
      const rows = getChildrenNS(t.element, NS.w, "tr")
      out.push(
        `=== Table ${t.tableIndex} (${summary.classification}) ${rows.length}×${summary.cols} ===`,
      )
      for (let r = 0; r < rows.length; r++) {
        const cells = getChildrenNS(rows[r]!, NS.w, "tc")
        for (let c = 0; c < cells.length; c++) {
          const text = cellText(cells[c]!)
          const snippet = text.length > 40 ? text.slice(0, 40) + "…" : text
          const paraSpan = formatParaSpan(cells[c]!, indexByElement)
          // Display 1-based to match the cell locator's 1-based row/col fields.
          out.push(`  [${r + 1},${c + 1}] ${JSON.stringify(snippet)}${paraSpan}`)
        }
      }
      out.push("")
    }
    console.log(out.join("\n").trimEnd())
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function cellText(tc: Element): string {
  let out = ""
  for (const p of getChildrenNS(tc, NS.w, "p")) {
    if (out) out += " ⏎ "
    const pText = paragraphText(p)
    out += pText
  }
  return out
}

/** "  paras: 58–89" / "  paras: 60" / "" (empty when cell paragraphs are
 * unindexed, i.e. inside a data/form table — cell locator addresses them
 * by [r,c] anyway). Layout-table paragraphs each have a #NNN; surfacing
 * the span lets agents pick a non-cross-cell range locator. */
function formatParaSpan(tc: Element, indexByElement: Map<Element, number>): string {
  const indices: number[] = []
  for (const p of getChildrenNS(tc, NS.w, "p")) {
    const idx = indexByElement.get(p)
    if (idx !== undefined) indices.push(idx)
  }
  if (indices.length === 0) return ""
  if (indices.length === 1) return `  paras: ${indices[0]}`
  // Indices are document-order so already sorted; emit a compact range when
  // they're contiguous, otherwise list explicitly. (Layout-table cells
  // typically have contiguous paragraphs; non-contiguous would be unusual.)
  const first = indices[0]!
  const last = indices[indices.length - 1]!
  const contiguous = last - first + 1 === indices.length
  return contiguous ? `  paras: ${first}–${last}` : `  paras: ${indices.join(", ")}`
}

void main()
