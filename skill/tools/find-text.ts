/**
 * find_text <docx> <pattern> [--regex] [--paragraph N | --range A-B] [--limit N] [--context N]
 *
 * Character-level text search returning per-match coordinates: paragraph
 * index, run index (when match is in a direct-child run), char offset,
 * length, and a context preview with the match bracketed.
 *
 * Read-only primitive: locate text by literal or regex; what the agent
 * does next (replace via set-run, deeper inspection, coverage validation,
 * just browsing) is up to them.
 */

import { loadDocx } from "@lib/load.ts"
import { searchDocument, describeRegion, type MatchHit } from "@lib/text-search.ts"
import { pad } from "@lib/format.ts"

async function main() {
  const argv = process.argv.slice(2)
  let isRegex = false
  let paragraph: number | undefined
  let range: { from: number; to: number } | undefined
  let limit = 50
  let contextChars = 25
  const positional: string[] = []
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!
    if (a === "--regex") isRegex = true
    else if (a === "--paragraph") {
      const v = argv[++i]
      if (!v || isNaN(parseInt(v, 10))) {
        console.error("--paragraph requires a positive integer")
        process.exit(1)
      }
      paragraph = parseInt(v, 10)
    } else if (a === "--range") {
      const v = argv[++i]
      const m = v?.match(/^(\d+)-(\d+)$/)
      if (!m) {
        console.error('--range requires the form "A-B" (e.g. "30-50")')
        process.exit(1)
      }
      range = { from: parseInt(m[1]!, 10), to: parseInt(m[2]!, 10) }
    } else if (a === "--limit") {
      const v = argv[++i]
      if (!v || isNaN(parseInt(v, 10))) {
        console.error("--limit requires a non-negative integer (0 = no cap)")
        process.exit(1)
      }
      limit = parseInt(v, 10)
    } else if (a === "--context") {
      const v = argv[++i]
      if (!v || isNaN(parseInt(v, 10))) {
        console.error("--context requires a non-negative integer")
        process.exit(1)
      }
      contextChars = parseInt(v, 10)
    } else if (a.startsWith("--")) {
      console.error(`unknown flag: ${a}`)
      process.exit(1)
    } else {
      positional.push(a)
    }
  }

  const file = positional[0]
  const pattern = positional[1]
  if (!file || pattern === undefined) {
    console.error(
      'Usage: node scripts/find_text.js <docx-path> <pattern> [--regex] [--paragraph N | --range A-B] [--limit N] [--context N]',
    )
    process.exit(1)
  }

  if (isRegex) {
    try {
      void new RegExp(pattern)
    } catch (err) {
      console.error(`Invalid regex /${pattern}/: ${(err as Error).message}`)
      process.exit(1)
    }
  }

  try {
    const doc = await loadDocx(file)
    const hits = searchDocument(doc.documentDoc, {
      pattern,
      regex: isRegex,
      paraIndex: paragraph,
      paraRange: range,
      limit: limit === 0 ? Infinity : limit,
      contextChars,
    })

    const lines: string[] = []
    const scopeStr = paragraph !== undefined
      ? `paragraph #${paragraph}`
      : range
        ? `paragraphs #${range.from}–#${range.to}`
        : "whole body"
    lines.push(`Pattern: ${isRegex ? `/${pattern}/` : `literal "${pattern}"`}`)
    lines.push(`Scope:   ${scopeStr}`)
    lines.push("")

    if (hits.length === 0) {
      lines.push("No matches.")
      console.log(lines.join("\n"))
      return
    }

    const paraSet = new Set(hits.map((h) => h.paragraphIndex))
    const counts = summarize(hits)
    const annotation = formatAnnotations(counts)
    lines.push(
      `${hits.length} matches across ${paraSet.size} paragraphs${annotation ? ` (${annotation})` : ""}:`,
    )
    lines.push("")

    for (const h of hits) {
      lines.push(formatMatchLine(h))
    }
    if (limit !== 0 && hits.length === limit) {
      lines.push("")
      lines.push(`(hit limit ${limit}; pass --limit 0 to see all matches)`)
    }
    console.log(lines.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

function formatMatchLine(h: MatchHit): string {
  const runCol = formatRunColumn(h)
  const ch = `ch=${String(h.ch).padStart(3, " ")}`
  const len = `len=${h.len}`
  const tags: string[] = []
  if (h.crossRun) tags.push("cross-run")
  if (h.region) tags.push(describeRegion(h.region))
  const tagStr = tags.length > 0 ? `   (${tags.join("; ")})` : ""
  return `  #${pad(h.paragraphIndex)}  ${runCol}  ${ch}  ${len}   ${h.context}${tagStr}`
}

function formatRunColumn(h: MatchHit): string {
  // Pad to a fixed width so columns line up across nested-vs-direct rows.
  const PAD = 12
  if (h.runIndex === null) return " ".repeat(PAD)
  const inner =
    h.crossRun && h.runIndexEnd !== null
      ? `run[${h.runIndex}..${h.runIndexEnd}]`
      : `run[${h.runIndex}]`
  return inner.padEnd(PAD, " ")
}

function summarize(hits: MatchHit[]): Record<string, number> {
  const out: Record<string, number> = {}
  for (const h of hits) {
    if (h.crossRun) out["cross-run"] = (out["cross-run"] ?? 0) + 1
    if (h.region) {
      const k = describeRegion(h.region)
      out[k] = (out[k] ?? 0) + 1
    }
  }
  return out
}

function formatAnnotations(counts: Record<string, number>): string {
  const keys = Object.keys(counts)
  if (keys.length === 0) return ""
  return keys.map((k) => `${counts[k]} ${k}`).join("; ")
}

main()
