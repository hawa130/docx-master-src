import { loadDocx } from "../core/load.ts"
import type { ParsedParagraph } from "../core/types.ts"

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
      out.push(`Computed style: ${formatComputed(first)}`)
    }
    out.push(`Referenced pStyles: [${styleIds.map((s) => `"${s}"`).join(", ")}]`)
    out.push(`Occurrences: ${matches.length}`)
    out.push("")
    const limit = Math.min(20, matches.length)
    for (let i = 0; i < limit; i++) {
      const p = matches[i]!
      const pred = p.context.predecessor ? p.context.predecessor.type : "none"
      out.push(
        `  #${pad(p.index)} [${p.fingerprint}] pred:${pred.padEnd(10)} "${truncate(p.text, 40)}"`,
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

function formatComputed(p: ParsedParagraph): string {
  const r = p.rPr
  const parts: string[] = []
  const font = r.fontEastAsia || r.fontAscii || r.fontHAnsi
  if (font) parts.push(`font: "${font}"`)
  if (r.size !== undefined) parts.push(`size: ${r.size} (${r.size / 2}pt)`)
  if (r.bold !== undefined) parts.push(`bold: ${r.bold}`)
  if (r.italic !== undefined) parts.push(`italic: ${r.italic}`)
  if (r.color) parts.push(`color: ${r.color}`)
  if (p.pPr.alignment) parts.push(`alignment: ${p.pPr.alignment}`)
  if (p.pPr.firstLineIndent) parts.push(`firstLineIndent: ${p.pPr.firstLineIndent}`)
  return `{ ${parts.join(", ")} }`
}

function truncate(s: string, n: number): string {
  const collapsed = s.replace(/\s+/g, " ").trim()
  if (collapsed.length <= n) return collapsed
  return collapsed.slice(0, n) + "…"
}

function pad(n: number): string {
  return n.toString().padStart(3, "0")
}

main()
