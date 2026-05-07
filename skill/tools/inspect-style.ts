import { loadDocx } from "@lib/load.ts"
import type { ParsedParagraph } from "@lib/types.ts"
import { formatLineSpacing, pad, truncate } from "@lib/format.ts"

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
      out.push(
        `  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`,
      )
    }
    if (matches.length > limit)
      out.push(`  ... (showing first ${limit}, total ${matches.length})`)
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function formatRPr(p: ParsedParagraph): string {
  const r = p.rPr
  const parts: string[] = []
  if (r.fontEastAsia) parts.push(`fontCJK: "${r.fontEastAsia}"`)
  const latin = r.fontAscii ?? r.fontHAnsi
  if (latin && latin !== r.fontEastAsia) parts.push(`fontLatin: "${latin}"`)
  if (r.size !== undefined) parts.push(`size: ${r.size / 2}pt`)
  if (r.bold) parts.push(`bold: true`)
  if (r.italic) parts.push(`italic: true`)
  if (r.color && r.color !== "auto") parts.push(`color: ${r.color}`)
  if (r.underline) parts.push(`underline: ${r.underline}`)
  return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`
}

function formatPPr(p: ParsedParagraph, all: ParsedParagraph[]): string {
  const pp = p.pPr
  const parts: string[] = []
  if (pp.alignment) parts.push(`alignment: ${pp.alignment}`)
  if (pp.spaceBefore !== undefined)
    parts.push(`spaceBefore: ${pp.spaceBefore / 20}pt`)
  if (pp.spaceAfter !== undefined)
    parts.push(`spaceAfter: ${pp.spaceAfter / 20}pt`)
  if (pp.lineSpacing !== undefined)
    parts.push(`lineSpacing: ${formatLineSpacing(pp.lineSpacing, pp.lineRule)}`)
  if (pp.firstLineIndentChars !== undefined)
    parts.push(`firstLineIndent: ${pp.firstLineIndentChars / 100}char`)
  else if (pp.firstLineIndent !== undefined)
    parts.push(`firstLineIndent: ${pp.firstLineIndent / 20}pt`)
  if (pp.hangingIndentChars !== undefined)
    parts.push(`hangingIndent: ${pp.hangingIndentChars / 100}char`)
  else if (pp.hangingIndent !== undefined)
    parts.push(`hangingIndent: ${pp.hangingIndent / 20}pt`)
  if (pp.outlineLevel !== undefined)
    parts.push(`outlineLevel: ${pp.outlineLevel}`)
  const numIdDisplay = resolveNumId(all)
  if (numIdDisplay) parts.push(`numId: ${numIdDisplay}`)
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
