import { loadDocx } from "@core/load.ts"
import type { ComputedParaStyle, ComputedRunStyle, ParsedParagraph } from "@core/types.ts"
import { formatLineSpacing, pad } from "@core/format.ts"

async function main() {
  const file = process.argv[2]
  const fromArg = process.argv[3]
  const toArg = process.argv[4]
  if (!file || !fromArg || !toArg) {
    console.error("Usage: node scripts/inspect_range.js <docx-path> <from> <to>")
    process.exit(1)
  }
  const from = parseInt(fromArg, 10)
  const to = parseInt(toArg, 10)
  if (isNaN(from) || isNaN(to) || from < 1 || to < from) {
    console.error("Invalid range")
    process.exit(1)
  }
  try {
    const doc = await loadDocx(file)
    const out: string[] = []
    for (const p of doc.paragraphs) {
      if (p.index < from || p.index > to) continue
      out.push(...renderPara(p))
      out.push("")
    }
    if (out.length === 0) {
      out.push(`No paragraphs in range ${from}-${to}`)
    }
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function renderPara(p: ParsedParagraph): string[] {
  const lines: string[] = []
  lines.push(`#${pad(p.index)} [${p.fingerprint}]`)
  lines.push(`  text: ${JSON.stringify(p.text)}`)
  lines.push(
    `  style: "${p.styleId}" (${p.styleName})  insideTable: ${p.context.insideTable ?? "(none)"}`,
  )
  lines.push(`  computed rPr: ${formatRPr(p.rPr)}`)
  lines.push(`  computed pPr: ${formatPPr(p.pPr)}`)
  lines.push(`  section: ${p.context.sectionIndex}`)
  return lines
}

function formatRPr(r: ComputedRunStyle): string {
  const parts: string[] = []
  const font = r.fontEastAsia || r.fontAscii || r.fontHAnsi
  if (font) parts.push(`font: "${font}"`)
  if (r.fontEastAsia && r.fontAscii && r.fontEastAsia !== r.fontAscii)
    parts.push(`fontAscii: "${r.fontAscii}"`)
  if (r.size !== undefined) parts.push(`size: ${r.size / 2}pt`)
  if (r.bold !== undefined) parts.push(`bold: ${r.bold}`)
  if (r.italic !== undefined) parts.push(`italic: ${r.italic}`)
  if (r.underline) parts.push(`underline: ${r.underline}`)
  if (r.color) parts.push(`color: ${r.color}`)
  if (r.highlight) parts.push(`highlight: ${r.highlight}`)
  if (r.strike) parts.push(`strike: ${r.strike}`)
  if (r.caps) parts.push(`caps: ${r.caps}`)
  return `{ ${parts.join(", ")} }`
}

function formatPPr(pp: ComputedParaStyle): string {
  const parts: string[] = []
  if (pp.alignment) parts.push(`alignment: ${pp.alignment}`)
  if (pp.spaceBefore !== undefined) parts.push(`spaceBefore: ${pp.spaceBefore / 20}pt`)
  if (pp.spaceAfter !== undefined) parts.push(`spaceAfter: ${pp.spaceAfter / 20}pt`)
  if (pp.lineSpacing !== undefined) {
    parts.push(`lineSpacing: ${formatLineSpacing(pp.lineSpacing, pp.lineRule)}`)
  }
  if (pp.indentLeft !== undefined) parts.push(`indentLeft: ${pp.indentLeft / 20}pt`)
  if (pp.indentRight !== undefined) parts.push(`indentRight: ${pp.indentRight / 20}pt`)
  if (pp.firstLineIndentChars !== undefined)
    parts.push(`firstLineIndent: ${pp.firstLineIndentChars / 100}char`)
  else if (pp.firstLineIndent !== undefined)
    parts.push(`firstLineIndent: ${pp.firstLineIndent / 20}pt`)
  if (pp.hangingIndentChars !== undefined)
    parts.push(`hangingIndent: ${pp.hangingIndentChars / 100}char`)
  else if (pp.hangingIndent !== undefined)
    parts.push(`hangingIndent: ${pp.hangingIndent / 20}pt`)
  if (pp.outlineLevel !== undefined) parts.push(`outlineLevel: ${pp.outlineLevel}`)
  if (pp.numId) parts.push(`numId: ${pp.numId}`)
  if (pp.numLevel !== undefined) parts.push(`numLevel: ${pp.numLevel}`)
  return `{ ${parts.join(", ")} }`
}

main()
