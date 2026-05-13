/**
 * inspect-caption: per-identifier view of SEQ-based captions in a doc.
 *
 * Usage:
 *   inspect_caption <docx-path> [identifier]
 *
 * Without identifier: lists all SEQ identifiers found in body with
 * occurrence counts.
 * With identifier: per-paragraph details — counter value, paragraph
 * index, anchor name (if any), referencing-REF count.
 */

import { loadDocx } from "@lib/xml/load.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, wAttr } from "@lib/xml/xml-utils.ts"
import { parseFieldRuns, type FieldDetails } from "@lib/edit/fields/field-parse.ts"

const w = NS.w

interface Occurrence {
  paragraphIndex: number
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

  const byId = new Map<string, IdentifierSummary>()
  const refTargets = new Map<string, number>()
  let paragraphIndex = 0

  for (const para of walkParagraphs(body)) {
    paragraphIndex++
    const runs = paragraphRuns(para)
    if (runs.length === 0) continue
    const parsed = parseFieldRuns(runs)

    // Capture REFs (counted later against captions)
    for (const entry of parsed) {
      if (entry.kind === "field" && entry.fieldType === "REF") {
        const name = entry.details.bookmarkName
        if (name) refTargets.set(name, (refTargets.get(name) ?? 0) + 1)
      }
    }

    // Find SEQ fields in this paragraph (one or two: parent + optional sub)
    const seqFields: FieldDetails[] = []
    for (const entry of parsed) {
      if (entry.kind === "field" && entry.fieldType === "SEQ") seqFields.push(entry.details)
    }
    if (seqFields.length === 0) continue
    const parent = seqFields[0]!
    if (!parent.identifier) continue

    let summary = byId.get(parent.identifier)
    if (!summary) {
      summary = {
        identifier: parent.identifier,
        format: parent.format,
        restartAtOutlineLevel: parent.restartAtOutlineLevel,
        occurrences: [],
        referencingRefs: 0,
      }
      byId.set(parent.identifier, summary)
    }

    summary.occurrences.push({
      paragraphIndex,
      parentSeqValue: rawResultText(parsed, "SEQ", parent.identifier) ?? "",
      subSeqValue:
        seqFields.length > 1
          ? (rawResultText(parsed, "SEQ", seqFields[1]!.identifier ?? "") ?? "")
          : undefined,
      styleId: paragraphStyleId(para),
      anchorName: firstBookmarkName(para),
    })
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

function printSummary(byId: Map<string, IdentifierSummary>): void {
  if (byId.size === 0) {
    console.log("No SEQ-based captions detected in this document.")
    return
  }
  console.log(`SEQ-based captions detected (${byId.size}):`)
  for (const s of byId.values()) {
    const fmt = s.format ?? "(unknown)"
    const restart = s.restartAtOutlineLevel ? ` restart=${s.restartAtOutlineLevel}` : " global"
    console.log(
      `  ${s.identifier.padEnd(20)} format=${fmt.padEnd(10)}${restart.padEnd(14)} occurrences=${s.occurrences.length}  refs=${s.referencingRefs}`,
    )
  }
}

function printDetail(s: IdentifierSummary): void {
  const fmt = s.format ?? "(unknown)"
  const restart = s.restartAtOutlineLevel
    ? `outline level ${s.restartAtOutlineLevel}`
    : "global (no restart)"
  console.log(`Caption: ${s.identifier}`)
  console.log(`  format:           ${fmt}`)
  console.log(`  restart:          ${restart}`)
  console.log(`  occurrences:      ${s.occurrences.length}`)
  console.log(`  citations (REFs): ${s.referencingRefs}`)
  console.log("")
  console.log("  Occurrences:")
  for (const occ of s.occurrences) {
    const sub = occ.subSeqValue ? `${occ.parentSeqValue}${occ.subSeqValue}` : occ.parentSeqValue
    const anchor = occ.anchorName ?? "(none)"
    const styleId = occ.styleId ?? "(unstyled)"
    console.log(
      `    para ${String(occ.paragraphIndex).padStart(4)}  counter ${sub.padEnd(8)}  anchor: ${anchor.padEnd(20)}  style: ${styleId}`,
    )
  }
}

function paragraphRuns(paragraph: Element): Element[] {
  const out: Element[] = []
  for (const c of getChildren(paragraph)) {
    if (c.namespaceURI === w && c.localName === "r") out.push(c)
  }
  return out
}

function paragraphStyleId(paragraph: Element): string | undefined {
  const pPr = firstChildNS(paragraph, w, "pPr")
  if (!pPr) return undefined
  const pStyle = firstChildNS(pPr, w, "pStyle")
  if (!pStyle) return undefined
  return wAttr(pStyle, "val") ?? undefined
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

function* walkParagraphs(root: Element): Generator<Element> {
  for (const child of getChildren(root)) {
    if (child.namespaceURI !== w) continue
    if (child.localName === "p") {
      yield child
    } else if (child.localName === "tbl" || child.localName === "tr" || child.localName === "tc") {
      yield* walkParagraphs(child)
    }
  }
}

await main()
