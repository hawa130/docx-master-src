/**
 * overview <docx> [--paras=A..B | --paras=none] [--include-unused]
 *
 * Default output is the full survey (metadata, page setup, theme, style
 * definitions, numbering schemes, visual style summary, direct-format per
 * fingerprint, document skeleton). Two scoping flags trim long-doc output
 * without losing what an agent needs for first-time orientation:
 *
 *   --paras=A..B      Render the skeleton only for paragraphs in [A, B].
 *                     Structural elements (tables / images / breaks) that
 *                     fall within the range are kept; everything else is
 *                     dropped. Use for second-look drilling after the
 *                     initial full survey.
 *   --paras=none      Skip the skeleton entirely. Useful when the agent
 *                     only needs the style / numbering / theme dictionary
 *                     (e.g. composing a small-scope `edits[]` config).
 *   --include-unused  Show styles with usage=0 (hidden by default — Word
 *                     and template residue often leaves dozens of declared-
 *                     but-unused style definitions that drown the table).
 */
import { loadDocx, parseNumbering } from "@lib/xml/load.ts"
import { walkIndexedParagraphs } from "@lib/edit/locator.ts"
import { NS, type DocumentElement, type ParsedParagraph } from "@lib/parse/types.ts"
import type { LoadedDoc } from "@lib/xml/load.ts"
import { firstChildNS, getChildren } from "@lib/xml/xml-utils.ts"
import { pad, paperName, truncate, tw2mm } from "@lib/parse/format.ts"

type ParasMode = "full" | "none" | { from: number; to: number }

interface Args {
  file: string
  paras: ParasMode
  includeUnused: boolean
}

function parseArgs(): Args {
  const argv = process.argv.slice(2)
  let paras: ParasMode = "full"
  let includeUnused = false
  const positional: string[] = []
  for (const a of argv) {
    if (a === "--include-unused") {
      includeUnused = true
    } else if (a.startsWith("--paras=")) {
      const spec = a.slice("--paras=".length)
      if (spec === "none") {
        paras = "none"
      } else {
        const m = spec.match(/^(\d+)\.\.(\d+)$/)
        if (!m) {
          console.error(`--paras: expected "A..B" or "none", got "${spec}"`)
          process.exit(1)
        }
        const from = parseInt(m[1]!, 10)
        const to = parseInt(m[2]!, 10)
        if (from < 1 || to < from) {
          console.error(`--paras: invalid range ${from}..${to} (require 1 ≤ from ≤ to)`)
          process.exit(1)
        }
        paras = { from, to }
      }
    } else if (a.startsWith("--")) {
      console.error(`unknown flag: ${a}`)
      process.exit(1)
    } else {
      positional.push(a)
    }
  }
  const file = positional[0]
  if (!file) {
    console.error(
      "Usage: node scripts/overview.js <docx-path> [--paras=A..B | --paras=none] [--include-unused]",
    )
    process.exit(1)
  }
  return { file, paras, includeUnused }
}

async function main() {
  const args = parseArgs()
  try {
    const doc = await loadDocx(args.file)
    const out: string[] = []
    out.push(...renderMetadata(doc))
    out.push("")
    out.push(...renderPageSetup(doc))
    out.push("")
    out.push(...renderTheme(doc))
    out.push("")
    out.push(...renderStyleDefinitions(doc, args.includeUnused))
    out.push("")
    out.push(...renderNumbering(doc))
    out.push("")
    out.push(...renderVisualSummary(doc))
    out.push("")
    out.push(...renderDirectFormatPerFingerprint(doc))
    out.push("")
    out.push(...renderSkeleton(doc, args.paras))
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
  lines.push(`Paper:       ${paper} (${tw2mm(s.pageSize.width)} × ${tw2mm(s.pageSize.height)} mm)`)
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
  const accents = accentSlots.filter((s) => colors[s]).map((s) => `${s}=#${colors[s]}`)
  if (accents.length > 0) lines.push(`Accent colors: ${accents.join(" ")}`)
  if (lines.length === 1) lines.push("(no theme)")
  return lines
}

function renderStyleDefinitions(doc: LoadedDoc, includeUnused: boolean): string[] {
  const lines = ["=== Style Definitions (styles.xml) ==="]
  const allNonDefault = doc.resolver.getAllStyles().filter((s) => !s.isDefault)
  const shown = includeUnused ? allNonDefault : allNonDefault.filter((s) => s.usageCount > 0)
  shown.sort((a, b) => b.usageCount - a.usageCount || a.id.localeCompare(b.id))
  if (shown.length === 0) {
    lines.push("(no styles)")
  } else {
    for (const s of shown) {
      const props: string[] = []
      if (s.rPr.size !== undefined) props.push(`${s.rPr.size / 2}pt`)
      if (s.rPr.bold) props.push("Bold")
      if (s.rPr.italic) props.push("Italic")
      if (s.rPr.fontEastAsia || s.rPr.fontAscii)
        props.push(s.rPr.fontEastAsia || s.rPr.fontAscii || "")
      if (s.pPr.alignment) props.push(s.pPr.alignment)
      if (s.pPr.outlineLevel !== undefined) props.push(`outlineLvl=${s.pPr.outlineLevel}`)
      if (s.pPr.numId) props.push(`numId=${s.pPr.numId}`)
      const based = s.basedOn ? ` basedOn=${s.basedOn}` : ""
      lines.push(
        `  ${s.id} "${s.name}" [${s.type}]${based}  usage=${s.usageCount}  {${props.join(", ")}}`,
      )
    }
  }
  const hidden = allNonDefault.length - shown.length
  if (hidden > 0) {
    lines.push(
      `  (${hidden} unused style${hidden === 1 ? "" : "s"} hidden — pass --include-unused to show)`,
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
    const meaningful = sample.levels.filter((l) => l.text.length > 0 || l.pStyle !== undefined)
    const shown = meaningful.length > 0 ? meaningful : sample.levels.slice(0, 1)
    for (const lvl of shown) {
      const ps = lvl.pStyle ? ` pStyle=${lvl.pStyle}` : ""
      const startNote = group.length === 1 ? ` start=${lvl.start}` : ""
      lines.push(`    L${lvl.level}: numFmt=${lvl.format} lvlText="${lvl.text}"${startNote}${ps}`)
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
    // Trailing facts (avg text length + dominant pStyle binding) are the cheap
    // signals that distinguish content fingerprints from chrome / form labels:
    // long avg + repeating pStyle = content; short avg + scattered styles =
    // probably pre-printed labels.
    const facts: string[] = [`×${s.count}`, `avg ${s.avgTextLength}ch`]
    if (s.boundStyleId) {
      facts.push(
        s.boundStyleName ? `via "${s.boundStyleName}"/${s.boundStyleId}` : `via ${s.boundStyleId}`,
      )
    }
    lines.push(`${s.label} [${s.hash}]: ${s.description.padEnd(36, " ")} ${facts.join("  ")}`)
  }
  return lines
}

/** For each fingerprint, list which pPr children the first sample paragraph
 * carries DIRECTLY (not via style cascade), plus the union of run-level rPr
 * children across the paragraph's `<w:r>` children. Lets the agent decide
 * whether an attribute is already set as chrome convention — in which case
 * the style shouldn't redeclare it. The values themselves aren't shown here
 * (use `inspect_range` for those); presence/absence is what governs the
 * "don't override what chrome already provides" rule across both axes. */
function renderDirectFormatPerFingerprint(doc: LoadedDoc): string[] {
  const lines = ["=== Direct Format per Fingerprint ==="]
  const indexToElement = new Map<number, Element>()
  for (const p of walkIndexedParagraphs(doc.documentDoc)) {
    indexToElement.set(p.index, p.element)
  }
  const firstByLabel = new Map<string, ParsedParagraph>()
  for (const p of doc.paragraphs) {
    if (!firstByLabel.has(p.fingerprint)) firstByLabel.set(p.fingerprint, p)
  }
  for (const s of doc.summary) {
    const sample = firstByLabel.get(s.label)
    if (!sample) continue
    const el = indexToElement.get(sample.index)
    if (!el) continue
    const pPrNames: string[] = []
    const pPr = firstChildNS(el, NS.w, "pPr")
    if (pPr) {
      for (const c of getChildren(pPr)) {
        if (c.namespaceURI !== NS.w) continue
        if (c.localName === "rPr") {
          // paragraph-mark rPr — note it but don't list its inner attrs
          pPrNames.push("rPr(pMark)")
        } else {
          pPrNames.push(c.localName!)
        }
      }
    }
    const rPrNames = new Set<string>()
    for (const c of getChildren(el)) {
      if (c.namespaceURI !== NS.w || c.localName !== "r") continue
      const rPr = firstChildNS(c, NS.w, "rPr")
      if (!rPr) continue
      for (const cc of getChildren(rPr)) {
        if (cc.namespaceURI !== NS.w) continue
        rPrNames.add(cc.localName!)
      }
    }
    const pPrStr = pPrNames.length > 0 ? pPrNames.join(", ") : "(none)"
    const rPrStr = rPrNames.size > 0 ? [...rPrNames].join(", ") : "(none)"
    lines.push(`  ${s.label}  pPr: ${pPrStr}  |  rPr: ${rPrStr}`)
  }
  return lines
}

function renderSkeleton(doc: LoadedDoc, paras: ParasMode): string[] {
  if (paras === "none") {
    return [
      "=== Document Skeleton ===",
      "(omitted — pass --paras=A..B or remove the flag to render the skeleton)",
    ]
  }
  const range = paras === "full" ? null : paras
  const header =
    range === null
      ? "=== Document Skeleton ==="
      : `=== Document Skeleton (paragraphs #${pad(range.from)}–#${pad(range.to)} of ${doc.paragraphs.length}) ===`
  const lines = [header]

  // group elements by section
  const sectionElements = new Map<number, DocumentElement[]>()
  for (const el of doc.elements) {
    const idx = elementSectionIndex(el)
    if (!sectionElements.has(idx)) sectionElements.set(idx, [])
    sectionElements.get(idx)!.push(el)
  }

  let renderedAnything = false
  for (const sec of doc.sections) {
    if (range && (sec.paraRange[1] < range.from || sec.paraRange[0] > range.to)) continue
    const headerLine = `--- Section ${sec.index + 1} (para #${sec.paraRange[0]}-#${sec.paraRange[1]}) ---`
    lines.push(headerLine)
    if (sec.header !== null) lines.push(`Header: ${truncate(sec.header, 60)}`)
    else lines.push(`Header: (none)`)
    if (sec.footer !== null) lines.push(`Footer: ${truncate(sec.footer, 60)}`)
    else lines.push(`Footer: (none)`)
    if (sec.footerPageNumFormat) lines.push(`Footer page num: ${sec.footerPageNumFormat}`)
    lines.push("")
    const elems = sectionElements.get(sec.index) || []
    // Track running paragraph index so non-indexed elements (data/form tables,
    // images, page breaks, equations) can be filtered by their positional
    // context — they belong to the slice iff a preceding-or-current indexed
    // paragraph falls inside it.
    let anchor = sec.paraRange[0]
    for (const el of elems) {
      const elemAnchor = updateAnchor(el, anchor)
      anchor = elemAnchor
      if (range && !elementInRange(el, range, elemAnchor)) continue
      lines.push(...renderElement(el, "", range))
      renderedAnything = true
    }
    lines.push("")
  }
  if (range && !renderedAnything) {
    lines.push(`(no paragraphs in range #${range.from}-#${range.to})`)
  }
  return lines
}

function updateAnchor(el: DocumentElement, prev: number): number {
  if (el.kind === "paragraph") return el.paragraph.index
  if (el.kind === "emptyRun") return el.firstIndex
  if (el.kind === "table" && el.classification === "layout" && el.paragraphs.length > 0) {
    return el.paragraphs[0]!.index
  }
  return prev
}

function elementInRange(
  el: DocumentElement,
  range: { from: number; to: number },
  anchor: number,
): boolean {
  if (el.kind === "paragraph") {
    const i = el.paragraph.index
    return i >= range.from && i <= range.to
  }
  if (el.kind === "emptyRun") {
    // A run of empty paragraphs spans firstIndex..firstIndex+count-1; include
    // it whenever any of its paragraphs overlaps the slice.
    const lastIndex = el.firstIndex + el.count - 1
    return lastIndex >= range.from && el.firstIndex <= range.to
  }
  if (el.kind === "table" && el.classification === "layout") {
    if (el.paragraphs.length === 0) return false
    return el.paragraphs.some((p) => p.index >= range.from && p.index <= range.to)
  }
  // Non-indexed elements (data/form tables, image, pageBreak, equation,
  // sectionBreak): treat as attached to their positional anchor.
  return anchor >= range.from && anchor <= range.to
}

function elementSectionIndex(el: DocumentElement): number {
  if (el.kind === "paragraph") return el.paragraph.context.sectionIndex
  return (el as any).sectionIndex ?? 0
}

function renderElement(
  el: DocumentElement,
  indent: string,
  range: { from: number; to: number } | null,
): string[] {
  const lines: string[] = []
  if (el.kind === "paragraph") {
    const p = el.paragraph
    lines.push(`${indent}  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`)
  } else if (el.kind === "table") {
    if (el.classification === "layout") {
      lines.push(`${indent}--- LAYOUT TABLE ---`)
      for (const p of el.paragraphs) {
        if (range && (p.index < range.from || p.index > range.to)) continue
        lines.push(`${indent}    #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`)
      }
      lines.push(`${indent}--- END LAYOUT TABLE ---`)
    } else if (el.classification === "data") {
      const headers = el.headers
        .map((h) => `"${truncate(h, 12)}"`)
        .slice(0, el.cols)
        .join(",")
      lines.push(`${indent}--- TABLE (${el.rows}×${el.cols}) headers:[${headers}] ---`)
    } else {
      lines.push(`${indent}--- FORM TABLE (${el.rows}×${el.cols}) ---`)
    }
  } else if (el.kind === "image") {
    lines.push(`${indent}--- IMAGE (${el.widthCm.toFixed(1)}cm × ${el.heightCm.toFixed(1)}cm) ---`)
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

function formatSize(n: number): string {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`
  return `${(n / 1024 / 1024).toFixed(2)} MB`
}

main()
