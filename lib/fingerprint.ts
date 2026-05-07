import { createHash } from "node:crypto"
import type { ParsedParagraph } from "./types.ts"
import type { StyleResolver } from "./style-resolver.ts"

export interface FingerprintSummary {
  /** Display label, sorted by frequency. Volatile across doc edits — fine for in-session interactive use. */
  label: string
  /** 6-char content hash derived from the rawFingerprint. Stable across runs and edits — use this in persisted configs. */
  hash: string
  description: string
  count: number
  rawFingerprint: string
  /** Average non-whitespace text length across paragraphs sharing this fingerprint.
   * Cheap signal that distinguishes content (long) from chrome / form labels (short). */
  avgTextLength: number
  /** Dominant pStyle binding (id + display name) when ≥80% of paragraphs with
   * this fingerprint reference the same styleId; undefined when split. Lets the
   * agent see "this fingerprint comes from style X" without an extra inspect_style_def call. */
  boundStyleId?: string
  boundStyleName?: string
}

export class Fingerprinter {
  assign(
    paragraphs: ParsedParagraph[],
    styleResolver?: StyleResolver,
  ): {
    labels: Map<string, string>
    hashes: Map<string, string>
    summary: FingerprintSummary[]
  } {
    const counts = new Map<string, number>()
    const samples = new Map<string, ParsedParagraph>()
    const totalTextLen = new Map<string, number>()
    // Tracks explicit pStyle bindings only; paragraphs without a styleId
    // simply don't contribute, so pickDominantStyle's percentage check
    // naturally accounts for "mostly unstyled" cases.
    const styleIdCounts = new Map<string, Map<string, number>>()
    for (const p of paragraphs) {
      const hash = makeHash(p)
      counts.set(hash, (counts.get(hash) || 0) + 1)
      if (!samples.has(hash)) samples.set(hash, p)
      totalTextLen.set(hash, (totalTextLen.get(hash) ?? 0) + p.text.trim().length)
      if (p.styleId) {
        let inner = styleIdCounts.get(hash)
        if (!inner) {
          inner = new Map()
          styleIdCounts.set(hash, inner)
        }
        inner.set(p.styleId, (inner.get(p.styleId) ?? 0) + 1)
      }
    }

    // sort by frequency desc, then by raw hash for stability
    const sorted = Array.from(counts.entries()).sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1]
      return a[0].localeCompare(b[0])
    })

    const labels = new Map<string, string>()
    const hashes = new Map<string, string>()
    const summary: FingerprintSummary[] = []
    for (let i = 0; i < sorted.length; i++) {
      const [hash, count] = sorted[i]!
      const label = letterLabel(i)
      const contentHash = shortHash(hash)
      labels.set(hash, label)
      hashes.set(hash, contentHash)
      const sample = samples.get(hash)!
      const avgTextLength = Math.round((totalTextLen.get(hash) ?? 0) / count)
      const dominant = pickDominantStyle(styleIdCounts.get(hash) ?? new Map(), count)
      const boundStyleName =
        dominant && styleResolver
          ? styleResolver.getStyleDefinition(dominant)?.name
          : undefined
      summary.push({
        label,
        hash: contentHash,
        description: describe(sample),
        count,
        rawFingerprint: hash,
        avgTextLength,
        boundStyleId: dominant,
        boundStyleName,
      })
    }

    for (const p of paragraphs) {
      const h = makeHash(p)
      p.fingerprint = labels.get(h) || "?"
    }

    return { labels, hashes, summary }
  }
}

/** Return the styleId used by ≥80% of paragraphs sharing this fingerprint, or
 * undefined when the binding is split — including the "mostly unstyled"
 * case (because paragraphs without an explicit pStyle don't contribute to
 * `styleCounts`, so their share is implicitly counted against the total).
 * Skips "Normal" because every Word doc inherits it by default; surfacing
 * it adds no signal beyond "no custom binding". */
function pickDominantStyle(
  styleCounts: Map<string, number>,
  total: number,
): string | undefined {
  let bestId: string | undefined
  let bestCount = 0
  for (const [id, n] of styleCounts) {
    if (n > bestCount) {
      bestId = id
      bestCount = n
    }
  }
  if (bestId === undefined) return undefined
  if (bestId === "Normal") return undefined
  if (bestCount / total < 0.8) return undefined
  return bestId
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
  // Structural label takes precedence over generic "List": after standardization,
  // headings carry both numId AND outlineLevel (because the heading style binds
  // numbering AND sets outline level for TOC). Showing "List" on a Heading is
  // technically true but reads as wrong — surface "Heading-N" so the visual
  // summary matches the agent's mental model. Plain list items (numId, no
  // outline level) still get "List". Non-numbered headings (outline level
  // without numId — e.g. abstract / acknowledgments under HeadingNoNum) also
  // get "Heading-N" so structure is visible regardless of numbering.
  if (pp.outlineLevel !== undefined) {
    parts.push(`Heading-${pp.outlineLevel + 1}`)
  } else if (pp.numId) {
    parts.push("List")
  }
  return parts.join(" ") || "(no formatting)"
}

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1)
}

/**
 * Content-derived stable identifier — first 6 hex chars of SHA-256 of the
 * rawFingerprint string. Same visual fingerprint across docs / edits maps
 * to the same hash, so configs that reference fingerprints by hash survive
 * doc changes that would shuffle the frequency-sorted letter labels.
 */
function shortHash(raw: string): string {
  return createHash("sha256").update(raw).digest("hex").slice(0, 6)
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
