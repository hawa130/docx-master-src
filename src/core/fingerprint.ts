import type { ParsedParagraph } from "./types.ts"

export interface FingerprintSummary {
  label: string
  description: string
  count: number
  rawFingerprint: string
}

export class Fingerprinter {
  assign(paragraphs: ParsedParagraph[]): {
    labels: Map<string, string>
    summary: FingerprintSummary[]
  } {
    const counts = new Map<string, number>()
    const samples = new Map<string, ParsedParagraph>()
    for (const p of paragraphs) {
      const hash = makeHash(p)
      counts.set(hash, (counts.get(hash) || 0) + 1)
      if (!samples.has(hash)) samples.set(hash, p)
    }

    // sort by frequency desc, then by raw hash for stability
    const sorted = Array.from(counts.entries()).sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1]
      return a[0].localeCompare(b[0])
    })

    const labels = new Map<string, string>()
    const summary: FingerprintSummary[] = []
    for (let i = 0; i < sorted.length; i++) {
      const [hash, count] = sorted[i]!
      const label = letterLabel(i)
      labels.set(hash, label)
      const sample = samples.get(hash)!
      summary.push({
        label,
        description: describe(sample),
        count,
        rawFingerprint: hash,
      })
    }

    for (const p of paragraphs) {
      const h = makeHash(p)
      p.fingerprint = labels.get(h) || "?"
    }

    return { labels, summary }
  }
}

function makeHash(p: ParsedParagraph): string {
  const r = p.rPr
  const pp = p.pPr
  const font = r.fontAscii || r.fontHAnsi || r.fontEastAsia || "?"
  const size = r.size !== undefined ? String(r.size) : "?"
  const flags =
    (r.bold ? "B" : "") +
    (r.italic ? "I" : "") +
    (r.underline ? "U" : "") +
    (r.caps ? "C" : "")
  const color = r.color && r.color !== "auto" ? r.color : ""
  const alignment = pp.alignment || ""
  const indent =
    pp.firstLineIndent || pp.firstLineIndentChars ? "1stInd" : ""
  // Include numbering presence so list items split out from visually-identical
  // body paragraphs. Without this, two paragraphs that share the same rPr
  // (e.g. 11pt non-bold body text vs. 11pt non-bold list item) would collapse
  // into one fingerprint and bulk_rules couldn't target lists separately.
  const list = pp.numId ? "L" : ""
  return `${font}|${size}|${flags}|${color}|${alignment}|${indent}|${list}`
}

function describe(p: ParsedParagraph): string {
  const r = p.rPr
  const pp = p.pPr
  const parts: string[] = []
  const font = r.fontEastAsia || r.fontAscii || r.fontHAnsi
  if (font) parts.push(font)
  if (r.size !== undefined) parts.push(`${r.size / 2}pt`)
  if (r.bold) parts.push("Bold")
  if (r.italic) parts.push("Italic")
  if (r.underline) parts.push("Underline")
  if (r.caps) parts.push("Caps")
  if (r.color && r.color !== "auto") parts.push(`#${r.color}`)
  if (pp.alignment) parts.push(capitalize(pp.alignment))
  if (pp.firstLineIndent || pp.firstLineIndentChars) parts.push("1stIndent")
  if (pp.numId) parts.push("List")
  return parts.join(" ") || "(no formatting)"
}

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1)
}

function letterLabel(i: number): string {
  // A, B, ..., Z, AA, AB, ...
  let n = i
  let out = ""
  do {
    out = String.fromCharCode(65 + (n % 26)) + out
    n = Math.floor(n / 26) - 1
  } while (n >= 0)
  return out
}
