import { loadDocx, parseNumbering } from "@lib/xml/load.ts"
import type { ComputedParaStyle, ComputedRunStyle } from "@lib/parse/types.ts"
import { formatComputedPPrParts, formatComputedRPrParts } from "@lib/parse/format.ts"

async function main() {
  const file = process.argv[2]
  const styleId = process.argv[3]
  if (!file || !styleId) {
    console.error("Usage: node scripts/inspect_style_def.js <docx-path> <styleId>")
    process.exit(1)
  }
  try {
    const doc = await loadDocx(file)
    const def = doc.resolver.getStyleDefinition(styleId)
    if (!def) {
      console.error(`Style "${styleId}" not found`)
      process.exit(1)
    }
    const out: string[] = []
    out.push(`Style "${def.id}" (${def.name})`)
    out.push(`  type: ${def.type}`)
    out.push(`  basedOn: ${def.basedOn ?? "(none)"}`)
    out.push("")
    out.push("  Inheritance chain:")
    const chain = doc.resolver.resolveStyleChain(styleId)
    const docDefaults = doc.resolver.getDocDefaults()
    out.push(`    docDefaults → ${formatStyleProps(docDefaults.rPr, docDefaults.pPr)}`)
    for (const id of chain.chain) {
      const raw = doc.resolver.getStyleDefinition(id)
      if (!raw) continue
      out.push(`    ${id.padEnd(12)} → ${formatStyleProps(raw.rPr, raw.pPr)}`)
    }
    out.push("")
    out.push(`  Final computed: ${formatStyleProps(chain.rPr, chain.pPr)}`)

    // numbering binding
    const numDefs = parseNumbering(doc.numberingDoc)
    let numBindingNote: string | null = null
    // 1) check style's own pPr.numId
    if (def.pPr.numId) {
      const numDef = numDefs.find((n) => n.numId === def.pPr.numId)
      if (numDef && def.pPr.numLevel !== undefined) {
        const lvl = numDef.levels.find((l) => l.level === def.pPr.numLevel)
        if (lvl)
          numBindingNote = `numId=${def.pPr.numId}, level=${lvl.level} (format: "${lvl.text}")`
      } else if (numDef) {
        numBindingNote = `numId=${def.pPr.numId}`
      }
    }
    // 2) check if any abstractNum lvl has pStyle === styleId
    if (!numBindingNote) {
      for (const n of numDefs) {
        const lvl = n.levels.find((l) => l.pStyle === styleId)
        if (lvl) {
          numBindingNote = `linked to numId=${n.numId}, level=${lvl.level} (format: "${lvl.text}")`
          break
        }
      }
    }
    if (numBindingNote) {
      out.push("")
      out.push(`  Numbering: ${numBindingNote}`)
    }

    out.push("")
    out.push(`  Usage: ${def.usageCount} paragraphs reference this style`)
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function formatStyleProps(r: ComputedRunStyle, pp: ComputedParaStyle): string {
  const parts = [
    ...formatComputedRPrParts(r),
    ...formatComputedPPrParts(pp, { includeIndentSides: true }),
  ]
  return parts.length === 0 ? "{ }" : `{ ${parts.join(", ")} }`
}

main()
