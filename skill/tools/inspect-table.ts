import { DocxReader } from "@lib/reader.ts"
import { walkTopLevelTables } from "@lib/locator.ts"
import { NS } from "@lib/types.ts"
import { getChildren, getChildrenNS, textContent } from "@lib/xml-utils.ts"
import { summarizeTable } from "@lib/table-classifier.ts"

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
          out.push(`  [${r},${c}] ${JSON.stringify(snippet)}`)
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

function paragraphText(p: Element): string {
  let out = ""
  for (const r of getChildrenNS(p, NS.w, "r")) {
    for (const c of getChildren(r)) {
      if (c.namespaceURI === NS.w && c.localName === "t") out += textContent(c)
    }
  }
  return out
}

void main()
