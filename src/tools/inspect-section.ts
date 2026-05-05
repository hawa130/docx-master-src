import { loadDocx } from "../core/load.ts"
import type { SectionInfo } from "../core/types.ts"

async function main() {
  const file = process.argv[2]
  const idxArg = process.argv[3]
  if (!file || idxArg === undefined) {
    console.error("Usage: node scripts/inspect_section.js <docx-path> <section-index>")
    process.exit(1)
  }
  const idx = parseInt(idxArg, 10)
  try {
    const doc = await loadDocx(file)
    const sec = doc.sections[idx]
    if (!sec) {
      console.error(`Section ${idx} not found (have ${doc.sections.length})`)
      process.exit(1)
    }
    const out: string[] = []
    out.push(`Section ${sec.index} (paragraphs #${sec.paraRange[0]}-#${sec.paraRange[1]})`)
    const paper = paperName(sec.pageSize.width, sec.pageSize.height)
    const tw2mm = (t: number) => +(t / 56.6929).toFixed(1)
    out.push(`  Paper:       ${paper} (${tw2mm(sec.pageSize.width)} × ${tw2mm(sec.pageSize.height)} mm)`)
    out.push(`  Orientation: ${sec.orientation}`)
    out.push(
      `  Margins:     top=${tw2mm(sec.margins.top)} bottom=${tw2mm(sec.margins.bottom)} left=${tw2mm(sec.margins.left)} right=${tw2mm(sec.margins.right)} mm`,
    )
    out.push(`  Header: ${sec.header ?? "(none)"}`)
    if (sec.headerHasImage) out.push(`    contains image`)
    out.push(`  Footer: ${sec.footer ?? "(none)"}`)
    if (sec.footerPageNumFormat) out.push(`  Footer page format: ${sec.footerPageNumFormat}`)

    if (idx > 0) {
      const prev = doc.sections[idx - 1]!
      out.push("")
      out.push(`  Differs from Section ${prev.index}:`)
      const diffs = diffSections(prev, sec)
      if (diffs.length === 0) out.push(`    (no differences)`)
      for (const d of diffs) out.push(`    ${d}`)
    }
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function diffSections(a: SectionInfo, b: SectionInfo): string[] {
  const out: string[] = []
  const tw2mm = (t: number) => +(t / 56.6929).toFixed(1)
  if (a.pageSize.width !== b.pageSize.width || a.pageSize.height !== b.pageSize.height)
    out.push(`paper changed: ${tw2mm(a.pageSize.width)}×${tw2mm(a.pageSize.height)}mm → ${tw2mm(b.pageSize.width)}×${tw2mm(b.pageSize.height)}mm`)
  if (a.orientation !== b.orientation)
    out.push(`orientation: ${a.orientation} → ${b.orientation}`)
  for (const k of ["top", "bottom", "left", "right"] as const) {
    if (a.margins[k] !== b.margins[k])
      out.push(`margin.${k}: ${tw2mm(a.margins[k])}mm → ${tw2mm(b.margins[k])}mm`)
  }
  if ((a.header ?? "") !== (b.header ?? "")) {
    out.push(
      a.header === null ? `+ Header added: "${b.header}"` : b.header === null ? `- Header removed` : `Header changed`,
    )
  }
  if ((a.footer ?? "") !== (b.footer ?? "")) {
    out.push(
      a.footer === null ? `+ Footer added` : b.footer === null ? `- Footer removed` : `Footer changed`,
    )
  }
  if ((a.footerPageNumFormat ?? "") !== (b.footerPageNumFormat ?? ""))
    out.push(
      `Footer page format: ${a.footerPageNumFormat ?? "(none)"} → ${b.footerPageNumFormat ?? "(none)"}`,
    )
  return out
}

function paperName(w: number, h: number): string {
  const known: Array<[string, number, number]> = [
    ["A4", 11906, 16838],
    ["A4", 16838, 11906],
    ["A3", 16838, 23811],
    ["A5", 8392, 11906],
    ["Letter", 12240, 15840],
    ["Legal", 12240, 20160],
  ]
  for (const [n, ww, hh] of known) {
    if (Math.abs(w - ww) < 50 && Math.abs(h - hh) < 50) return n
  }
  return "Custom"
}

main()
