import { loadDocx, parseNumbering } from "../core/load.ts"
import type { DocumentElement, ParsedParagraph, SectionInfo } from "../core/types.ts"
import type { LoadedDoc } from "../core/load.ts"

async function main() {
  const file = process.argv[2]
  if (!file) {
    console.error("Usage: node scripts/overview.js <docx-path>")
    process.exit(1)
  }
  try {
    const doc = await loadDocx(file)
    const out: string[] = []
    out.push(...renderMetadata(doc))
    out.push("")
    out.push(...renderPageSetup(doc))
    out.push("")
    out.push(...renderTheme(doc))
    out.push("")
    out.push(...renderStyleDefinitions(doc))
    out.push("")
    out.push(...renderNumbering(doc))
    out.push("")
    out.push(...renderVisualSummary(doc))
    out.push("")
    out.push(...renderSkeleton(doc))
    console.log(out.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function renderMetadata(doc: LoadedDoc): string[] {
  const lines = ["=== Document Metadata ==="]
  lines.push(`File:    ${doc.metadata.fileName}`)
  lines.push(`Size:    ${formatSize(doc.metadata.fileSize)}`)
  if (doc.metadata.title) lines.push(`Title:   ${doc.metadata.title}`)
  if (doc.metadata.author) lines.push(`Author:  ${doc.metadata.author}`)
  lines.push(`Paragraphs: ${doc.paragraphs.length}`)
  lines.push(`Sections:   ${doc.sections.length}`)
  return lines
}

function renderPageSetup(doc: LoadedDoc): string[] {
  const lines = ["=== Page Setup ==="]
  const s = doc.sections[doc.sections.length - 1] || doc.sections[0]
  if (!s) {
    lines.push("(no section properties)")
    return lines
  }
  const paper = paperName(s.pageSize.width, s.pageSize.height)
  const tw2mm = (t: number) => +(t / 56.6929).toFixed(1)
  lines.push(
    `Paper:       ${paper} (${tw2mm(s.pageSize.width)} × ${tw2mm(s.pageSize.height)} mm)`,
  )
  lines.push(`Orientation: ${s.orientation}`)
  lines.push(
    `Margins:     top=${tw2mm(s.margins.top)} bottom=${tw2mm(s.margins.bottom)} left=${tw2mm(s.margins.left)} right=${tw2mm(s.margins.right)} mm`,
  )
  return lines
}

function renderTheme(doc: LoadedDoc): string[] {
  const lines = ["=== Theme ==="]
  const fonts = doc.resolver.getThemeFonts()
  const colors = doc.resolver.getThemeColors()
  if (fonts.majorLatin || fonts.majorEastAsia)
    lines.push(
      `Major font: ${fonts.majorLatin || "?"}${fonts.majorEastAsia ? ` / ${fonts.majorEastAsia}` : ""}`,
    )
  if (fonts.minorLatin || fonts.minorEastAsia)
    lines.push(
      `Minor font: ${fonts.minorLatin || "?"}${fonts.minorEastAsia ? ` / ${fonts.minorEastAsia}` : ""}`,
    )
  const accentSlots = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
  const accents = accentSlots
    .filter((s) => colors[s])
    .map((s) => `${s}=#${colors[s]}`)
  if (accents.length > 0) lines.push(`Accent colors: ${accents.join(" ")}`)
  if (lines.length === 1) lines.push("(no theme)")
  return lines
}

function renderStyleDefinitions(doc: LoadedDoc): string[] {
  const lines = ["=== Style Definitions (styles.xml) ==="]
  const all = doc.resolver
    .getAllStyles()
    .filter((s) => !s.isDefault)
    .sort((a, b) => b.usageCount - a.usageCount || a.id.localeCompare(b.id))
  if (all.length === 0) {
    lines.push("(no styles)")
    return lines
  }
  for (const s of all) {
    const props: string[] = []
    if (s.rPr.size !== undefined) props.push(`${s.rPr.size / 2}pt`)
    if (s.rPr.bold) props.push("Bold")
    if (s.rPr.italic) props.push("Italic")
    if (s.rPr.fontEastAsia || s.rPr.fontAscii)
      props.push(s.rPr.fontEastAsia || s.rPr.fontAscii || "")
    if (s.pPr.alignment) props.push(s.pPr.alignment)
    if (s.pPr.outlineLevel !== undefined)
      props.push(`outlineLvl=${s.pPr.outlineLevel}`)
    if (s.pPr.numId) props.push(`numId=${s.pPr.numId}`)
    const based = s.basedOn ? ` basedOn=${s.basedOn}` : ""
    lines.push(
      `  ${s.id} "${s.name}" [${s.type}]${based}  usage=${s.usageCount}  {${props.join(", ")}}`,
    )
  }
  return lines
}

function renderNumbering(doc: LoadedDoc): string[] {
  const lines = ["=== Numbering Definitions (numbering.xml) ==="]
  const defs = parseNumbering(doc.numberingDoc)
  if (defs.length === 0) {
    lines.push("(no numbering)")
    return lines
  }

  // Cluster numIds whose abstractNum has identical lvlText pattern across all
  // levels — Word frequently emits N near-identical abstractNums (one per
  // list region, varying only by start values). Showing them collapsed beats
  // dumping 24 nearly-identical blocks.
  const clusters = new Map<string, typeof defs>()
  for (const def of defs) {
    const sig = def.levels
      .map((l) => `${l.level}:${l.format}|${l.text}|${l.pStyle ?? ""}`)
      .join(";;")
    if (!clusters.has(sig)) clusters.set(sig, [])
    clusters.get(sig)!.push(def)
  }

  let clusterIdx = 0
  for (const [, group] of clusters) {
    clusterIdx++
    const sample = group[0]!
    const ids = group.map((d) => d.numId).join(", ")
    const starts = new Set(group.map((d) => d.levels[0]?.start ?? 1))
    const startsStr =
      starts.size === 1
        ? `start=${[...starts][0]}`
        : `starts: {${[...starts].sort((a, b) => a - b).join(", ")}}`

    if (group.length === 1) {
      lines.push(`  numId=${sample.numId} (abstract=${sample.abstractNumId})`)
    } else {
      lines.push(`  Scheme ${clusterIdx} × ${group.length} numIds: [${ids}]  ${startsStr}`)
    }
    // Only show non-empty levels; trailing decimal-with-empty-text levels are
    // Word's default filler that distract from the actual scheme
    const meaningful = sample.levels.filter(
      (l) => l.text.length > 0 || l.pStyle !== undefined,
    )
    const shown = meaningful.length > 0 ? meaningful : sample.levels.slice(0, 1)
    for (const lvl of shown) {
      const ps = lvl.pStyle ? ` pStyle=${lvl.pStyle}` : ""
      const startNote = group.length === 1 ? ` start=${lvl.start}` : ""
      lines.push(
        `    L${lvl.level}: numFmt=${lvl.format} lvlText="${lvl.text}"${startNote}${ps}`,
      )
    }
  }
  return lines
}

function renderVisualSummary(doc: LoadedDoc): string[] {
  const lines = ["=== Visual Style Summary (deduplicated) ==="]
  for (const s of doc.summary) {
    // Letter is stable within this run; hash is content-derived and stable
    // across runs / edits — use it in persisted configs to survive
    // frequency-rank shuffles when paragraphs are added or removed.
    lines.push(`${s.label} [${s.hash}]: ${s.description.padEnd(36, " ")} ×${s.count}`)
  }
  return lines
}

function renderSkeleton(doc: LoadedDoc): string[] {
  const lines = ["=== Document Skeleton ==="]
  // group elements by section
  const sectionElements = new Map<number, DocumentElement[]>()
  for (const el of doc.elements) {
    const idx = elementSectionIndex(el)
    if (!sectionElements.has(idx)) sectionElements.set(idx, [])
    sectionElements.get(idx)!.push(el)
  }

  for (const sec of doc.sections) {
    const range = `(para #${sec.paraRange[0]}-#${sec.paraRange[1]})`
    lines.push(`--- Section ${sec.index + 1} ${range} ---`)
    if (sec.header !== null) lines.push(`Header: ${truncate(sec.header, 60)}`)
    else lines.push(`Header: (none)`)
    if (sec.footer !== null) lines.push(`Footer: ${truncate(sec.footer, 60)}`)
    else lines.push(`Footer: (none)`)
    if (sec.footerPageNumFormat) lines.push(`Footer page num: ${sec.footerPageNumFormat}`)
    lines.push("")
    const elems = sectionElements.get(sec.index) || []
    for (const el of elems) {
      lines.push(...renderElement(el, ""))
    }
    lines.push("")
  }
  return lines
}

function elementSectionIndex(el: DocumentElement): number {
  if (el.kind === "paragraph") return el.paragraph.context.sectionIndex
  return (el as any).sectionIndex ?? 0
}

function renderElement(el: DocumentElement, indent: string): string[] {
  const lines: string[] = []
  if (el.kind === "paragraph") {
    const p = el.paragraph
    lines.push(
      `${indent}  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`,
    )
  } else if (el.kind === "table") {
    if (el.classification === "layout") {
      lines.push(`${indent}--- LAYOUT TABLE ---`)
      for (const p of el.paragraphs) {
        lines.push(
          `${indent}    #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`,
        )
      }
      lines.push(`${indent}--- END LAYOUT TABLE ---`)
    } else if (el.classification === "data") {
      const headers = el.headers
        .map((h) => `"${truncate(h, 12)}"`)
        .slice(0, el.cols)
        .join(",")
      lines.push(
        `${indent}--- TABLE (${el.rows}×${el.cols}) headers:[${headers}] ---`,
      )
    } else {
      lines.push(`${indent}--- FORM TABLE (${el.rows}×${el.cols}) ---`)
    }
  } else if (el.kind === "image") {
    lines.push(
      `${indent}--- IMAGE (${el.widthCm.toFixed(1)}cm × ${el.heightCm.toFixed(1)}cm) ---`,
    )
  } else if (el.kind === "equation") {
    lines.push(`${indent}--- EQUATION ---`)
  } else if (el.kind === "pageBreak") {
    lines.push(`${indent}--- PAGEBREAK ---`)
  } else if (el.kind === "sectionBreak") {
    // section break is implicit in section header; skip
  } else if (el.kind === "emptyRun") {
    if (el.count === 1) {
      lines.push(`${indent}  #${pad(el.firstIndex)} --- empty ---`)
    } else {
      lines.push(`${indent}  #${pad(el.firstIndex)} --- empty ×${el.count} ---`)
    }
  }
  return lines
}

function truncate(s: string, n: number): string {
  const collapsed = s.replace(/\s+/g, " ").trim()
  if (collapsed.length <= n) return collapsed
  return collapsed.slice(0, n) + "…"
}

function pad(n: number): string {
  return n.toString().padStart(3, "0")
}

function formatSize(n: number): string {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`
  return `${(n / 1024 / 1024).toFixed(2)} MB`
}

function paperName(width: number, height: number): string {
  // common paper sizes in twips
  const known: Array<[string, number, number]> = [
    ["A4", 11906, 16838],
    ["A4", 16838, 11906],
    ["A3", 16838, 23811],
    ["A5", 8392, 11906],
    ["Letter", 12240, 15840],
    ["Legal", 12240, 20160],
  ]
  for (const [name, w, h] of known) {
    if (Math.abs(width - w) < 50 && Math.abs(height - h) < 50) return name
  }
  return "Custom"
}

main()
