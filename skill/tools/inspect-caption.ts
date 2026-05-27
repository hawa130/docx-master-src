/**
 * inspect-caption: per-identifier view of SEQ-based captions in a doc.
 *
 * Usage:
 *   inspect_caption <docx-path> [identifier]
 *
 * Without identifier: lists all SEQ identifiers found in body with
 * occurrence counts.
 * With identifier: per-paragraph details — counter value, location
 * (indexed `#NNN` or data-cell `T<n>R<n>C<n> K<n>`), anchor name (if any),
 * referencing-REF count.
 */

import { loadDocx } from "@lib/xml/load.ts"
import { NS } from "@lib/parse/types.ts"
import { summarizeTable } from "@lib/parse/table-classifier.ts"
import {
  firstChildNS,
  getChildren,
  getChildrenNS,
  paragraphRuns,
  paragraphStyleId,
  wAttr,
} from "@lib/xml/xml-utils.ts"
import { walkIndexedParagraphs } from "@lib/edit/locator.ts"
import { parseFieldRuns, seqFields } from "@lib/edit/fields/field-parse.ts"

const w = NS.w

/** Location of a caption paragraph: either a global 1-based index (body /
 * layout-table cells) or a set of data-cell coords. */
type ParagraphLocation =
  | { kind: "indexed"; index: number }
  | { kind: "cell"; table: number; row: number; col: number; paragraph: number }

interface Occurrence {
  location: ParagraphLocation
  parentSeqValue: string
  subSeqValue: string | undefined
  styleId: string | undefined
  anchorName: string | undefined
}

interface IdentifierSummary {
  identifier: string
  format: string | undefined
  restartAtOutlineLevel: number | undefined
  occurrences: Occurrence[]
  referencingRefs: number
}

async function main(): Promise<void> {
  const file = process.argv[2]
  const target = process.argv[3]
  if (!file) {
    console.error("Usage: inspect_caption <docx-path> [identifier]")
    process.exit(1)
  }
  const doc = await loadDocx(file)
  const documentDoc = doc.documentDoc
  if (!documentDoc) {
    console.error("inspect_caption: failed to read word/document.xml")
    process.exit(1)
  }
  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) {
    console.error("inspect_caption: document has no body")
    process.exit(1)
  }

  // Build index → element map for indexed paragraphs (body + layout cells).
  const indexedParas = walkIndexedParagraphs(documentDoc)
  const elementToIndex = new Map<Element, number>()
  for (const p of indexedParas) elementToIndex.set(p.element, p.index)

  const byId = new Map<string, IdentifierSummary>()
  const refTargets = new Map<string, number>()

  /** Process a single paragraph element at a given location. */
  const processPara = (para: Element, location: ParagraphLocation): void => {
    const runs = paragraphRuns(para)
    if (runs.length === 0) return
    const parsed = parseFieldRuns(runs)

    // Capture REFs (counted later against captions)
    for (const entry of parsed) {
      if (entry.kind === "field" && entry.fieldType === "REF") {
        const name = entry.details.bookmarkName
        if (name) refTargets.set(name, (refTargets.get(name) ?? 0) + 1)
      }
    }

    // Collect advancing SEQ fields. `seqFields(..., skipRepeat: true)` drops
    // `\c` SEQs (engine-injected chapter prefixes) — they read the counter
    // without advancing and shouldn't appear as caption occurrences.
    const advancingSeqs = seqFields(para, { skipRepeat: true })
    if (advancingSeqs.length === 0) return

    const styleId = paragraphStyleId(para)
    const anchorName = firstBookmarkName(para)
    for (const [i, seq] of advancingSeqs.entries()) {
      if (!seq.identifier) continue
      let summary = byId.get(seq.identifier)
      if (!summary) {
        summary = {
          identifier: seq.identifier,
          format: seq.format,
          restartAtOutlineLevel: seq.restartAtOutlineLevel,
          occurrences: [],
          referencingRefs: 0,
        }
        byId.set(seq.identifier, summary)
      }
      const parentSeqValue = i === 0 ? (rawResultText(parsed, "SEQ", seq.identifier) ?? "") : ""
      const subSeqValue =
        i === 0 && advancingSeqs.length > 1
          ? (rawResultText(parsed, "SEQ", advancingSeqs[1]!.identifier ?? "") ?? "")
          : undefined
      summary.occurrences.push({
        location,
        parentSeqValue,
        subSeqValue,
        styleId,
        anchorName,
      })
    }
  }

  // Phase 1: indexed paragraphs (body + layout-table cells).
  for (const p of indexedParas) {
    processPara(p.element, { kind: "indexed", index: p.index })
  }

  // Phase 2: data-table-cell paragraphs.
  let tableIdx = 0
  for (const child of getChildren(body)) {
    if (child.namespaceURI !== NS.w || child.localName !== "tbl") continue
    tableIdx++
    const summary = summarizeTable(child)
    if (summary.classification !== "data") continue
    const rows = getChildrenNS(child, NS.w, "tr")
    for (let ri = 0; ri < rows.length; ri++) {
      const cells = getChildrenNS(rows[ri]!, NS.w, "tc")
      for (let ci = 0; ci < cells.length; ci++) {
        const paras = getChildrenNS(cells[ci]!, NS.w, "p")
        for (let pi = 0; pi < paras.length; pi++) {
          processPara(paras[pi]!, {
            kind: "cell",
            table: tableIdx,
            row: ri + 1,
            col: ci + 1,
            paragraph: pi + 1,
          })
        }
      }
    }
  }

  // Compute referencing-REF counts per identifier (by anchor lookup).
  for (const summary of byId.values()) {
    for (const occ of summary.occurrences) {
      if (occ.anchorName) {
        summary.referencingRefs += refTargets.get(occ.anchorName) ?? 0
      }
    }
  }

  if (target) {
    const summary = byId.get(target)
    if (!summary) {
      console.error(`inspect_caption: no SEQ identifier "${target}" found in document body.`)
      console.error(
        `Known identifiers: ${[...byId.keys()].map((k) => `"${k}"`).join(", ") || "(none)"}`,
      )
      process.exit(1)
    }
    printDetail(summary)
    return
  }

  printSummary(byId)
}

function formatLocation(loc: ParagraphLocation): string {
  if (loc.kind === "indexed") return `#${String(loc.index).padStart(4)}`
  return `T${loc.table}R${loc.row}C${loc.col} K${loc.paragraph}`
}

function printSummary(byId: Map<string, IdentifierSummary>): void {
  if (byId.size === 0) {
    console.log("No SEQ-based captions detected in this document.")
    return
  }
  // Count how many occurrences are in data cells for each identifier.
  const cellCounts = new Map<string, number>()
  for (const s of byId.values()) {
    const n = s.occurrences.filter((o) => o.location.kind === "cell").length
    if (n > 0) cellCounts.set(s.identifier, n)
  }
  console.log(`SEQ-based captions detected (${byId.size}):`)
  for (const s of byId.values()) {
    const fmt = s.format ?? "(unknown)"
    const restart = s.restartAtOutlineLevel ? ` restart=${s.restartAtOutlineLevel}` : " global"
    const cellNote = cellCounts.has(s.identifier) ? ` (${cellCounts.get(s.identifier)} in cells)` : ""
    console.log(
      `  ${s.identifier.padEnd(20)} format=${fmt.padEnd(10)}${restart.padEnd(14)} occurrences=${s.occurrences.length}  refs=${s.referencingRefs}${cellNote}`,
    )
  }
}

function printDetail(s: IdentifierSummary): void {
  const fmt = s.format ?? "(unknown)"
  const restart = s.restartAtOutlineLevel
    ? `outline level ${s.restartAtOutlineLevel}`
    : "global (no restart)"
  const cellCount = s.occurrences.filter((o) => o.location.kind === "cell").length
  console.log(`Caption: ${s.identifier}`)
  console.log(`  format:           ${fmt}`)
  console.log(`  restart:          ${restart}`)
  console.log(`  occurrences:      ${s.occurrences.length}${cellCount > 0 ? ` (of which ${cellCount} inside data-table cells)` : ""}`)
  console.log(`  citations (REFs): ${s.referencingRefs}`)
  console.log("")
  console.log("  Occurrences:")
  for (const occ of s.occurrences) {
    const sub = occ.subSeqValue ? `${occ.parentSeqValue}${occ.subSeqValue}` : occ.parentSeqValue
    const anchor = occ.anchorName ?? "(none)"
    const styleId = occ.styleId ?? "(unstyled)"
    const loc = formatLocation(occ.location).padEnd(16)
    console.log(
      `    para ${loc}  counter ${sub.padEnd(8)}  anchor: ${anchor.padEnd(20)}  style: ${styleId}`,
    )
  }
}

function firstBookmarkName(paragraph: Element): string | undefined {
  for (const c of getChildren(paragraph)) {
    if (c.namespaceURI === w && c.localName === "bookmarkStart") {
      const name = wAttr(c, "name")
      return name ?? undefined
    }
  }
  return undefined
}

function rawResultText(
  parsed: ReturnType<typeof parseFieldRuns>,
  fieldType: "SEQ",
  identifier: string,
): string | undefined {
  for (const entry of parsed) {
    if (
      entry.kind === "field" &&
      entry.fieldType === fieldType &&
      entry.details.identifier === identifier
    ) {
      return entry.result
    }
  }
  return undefined
}

await main()
