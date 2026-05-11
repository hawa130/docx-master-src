/**
 * Best-effort detection of typed/manual numbering prefixes inside `edits[]`
 * Block paragraphs. Surfaced as a dry-run warning so the agent can fix
 * before commit. Two failure modes covered, distinguished in the report:
 *
 *   - `bound`: styleId IS bound to an auto-numbering scheme. The typed
 *     prefix will double-print with the scheme's emitted marker. Fix:
 *     drop the typed prefix from `text`.
 *   - `unbound`: styleId is NOT bound to numbering, but the paragraph
 *     starts with an enumeration-shape prefix. Likely a block enumeration
 *     that should use a list-bound style (e.g. ListNumber) with a
 *     numbering scheme; or intentional typed prefix (bibliography) that
 *     agent should affirm.
 *
 * Only inspects the config (`edits[]`) — chrome paragraphs that ride
 * through the rule-routing path are already covered by `unstrippedByStyle`.
 *
 * Heuristic by design: regex patterns cover the most common CJK-academic
 * shapes (`一、`, `（一）`, `1.`, `1.1`, `第N章`, ...). Roman numerals and
 * alphabetic markers are deliberately excluded — false-positive risk in
 * prose is too high.
 */

import type { EditOp, Block, RichText } from "@lib/config/edit-types.ts"

const COMMON_TYPED_PREFIX_PATTERNS: RegExp[] = [
  /^\s*\d+(\.\d+)+\s/,
  /^\s*\d+\.\s/,
  /^\s*\d+、/,
  /^\s*第[一二三四五六七八九十百千零〇0-9]+[章节篇部条款]/,
  /^\s*[一二三四五六七八九十百千零〇]+、/,
  /^\s*[（(][一二三四五六七八九十百千零〇0-9]+[)）]/,
  /^\s*[①-⑳⒈-⒛]/,
  /^\s*Chapter\s+\d+/i,
]

export interface ManualNumberingHit {
  count: number
  samples: string[]
  /** `bound`: styleId has an auto-numbering scheme attached (double-print
   *  risk). `unbound`: styleId has no numbering binding (potential
   *  block-enumeration miscategorisation). */
  kind: "bound" | "unbound"
}

export function detectManualNumbering(
  edits: EditOp[] | undefined,
  numberedStyleIds: Set<string>,
): Map<string, ManualNumberingHit> {
  const out = new Map<string, ManualNumberingHit>()
  if (!edits) return out
  for (const edit of edits) {
    const blocks = blocksOf(edit)
    if (!blocks) continue
    for (const block of blocks) {
      if (block.type !== "paragraph") continue
      if (!block.styleId) continue
      const text = richTextToPlain(block.text)
      if (!text) continue
      if (!COMMON_TYPED_PREFIX_PATTERNS.some((rx) => rx.test(text))) continue
      const kind: ManualNumberingHit["kind"] = numberedStyleIds.has(block.styleId)
        ? "bound"
        : "unbound"
      let hit = out.get(block.styleId)
      if (!hit) {
        hit = { count: 0, samples: [], kind }
        out.set(block.styleId, hit)
      }
      hit.count += 1
      if (hit.samples.length < 3) {
        const trimmed = text.trim().slice(0, 40)
        hit.samples.push(trimmed + (text.trim().length > 40 ? "…" : ""))
      }
    }
  }
  return out
}

function blocksOf(edit: EditOp): Block[] | null {
  if (edit.op === "replace") return edit.with
  if (edit.op === "insert-before" || edit.op === "insert-after") return edit.content
  return null
}

function richTextToPlain(rich: RichText): string {
  if (typeof rich === "string") return rich
  return rich.map((r) => r.text).join("")
}
