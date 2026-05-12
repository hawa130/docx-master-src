import { loadDocx } from "@lib/xml/load.ts"
import type { ParsedParagraph } from "@lib/parse/types.ts"
import { formatComputedPPrParts, formatComputedRPrParts, pad, truncate } from "@lib/parse/format.ts"

async function main() {
  const file = process.argv[2]
  const label = process.argv[3]
  if (!file || !label) {
    console.error("Usage: node scripts/inspect_style.js <docx-path> <fingerprint-label>")
    process.exit(1)
  }
  try {
    const doc = await loadDocx(file)
    const sum = doc.summary.find((s) => s.label === label)
    if (!sum) {
      console.error(`Fingerprint "${label}" not found`)
      process.exit(1)
    }
    const matches = doc.paragraphs.filter((p) => p.fingerprint === label)
    const styleIds = Array.from(new Set(matches.map((p) => p.styleId)))

    const out: string[] = []
    out.push(`Fingerprint ${label}: ${sum.description}`)
    if (matches.length > 0) {
      const first = matches[0]!
      out.push(`Computed rPr: ${formatRPr(first)}`)
      out.push(`Computed pPr: ${formatPPr(first, matches)}`)
    }
    out.push(`Referenced pStyles: [${styleIds.map((s) => `"${s}"`).join(", ")}]`)
    out.push(`Occurrences: ${matches.length}`)
    out.push("")
    const limit = Math.min(20, matches.length)
    for (let i = 0; i < limit; i++) {
      const p = matches[i]!
      out.push(`  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`)
    }
    if (matches.length > limit) out.push(`  ... (showing first ${limit}, total ${matches.length})`)
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function formatRPr(p: ParsedParagraph): string {
  const r = p.rPr
  const parts = formatComputedRPrParts(r, { truthyToggles: true, filterAutoColor: true })
  if (r.underline) parts.push(`underline: ${r.underline}`)
  return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`
}

function formatPPr(p: ParsedParagraph, all: ParsedParagraph[]): string {
  const parts = formatComputedPPrParts(p.pPr, { numIdDisplay: resolveNumId(all) })
  return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`
}

function resolveNumId(all: ParsedParagraph[]): string | null {
  const seen = new Set<string>()
  let hasUnset = false
  for (const p of all) {
    if (p.pPr.numId !== undefined) seen.add(p.pPr.numId)
    else hasUnset = true
  }
  if (seen.size === 0) return null
  if (seen.size === 1 && !hasUnset) return Array.from(seen)[0]!
  return "mixed"
}

main()
