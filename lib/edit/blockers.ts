/**
 * Detect paragraphs that aren't safe to edit. The edit engine consults this
 * before applying any op; if any resolved target paragraph is blocked, the
 * op is refused with a specific reason. Inspect tools surface the same data
 * so the agent can plan around blockers before composing edits.
 *
 * What counts as a blocker:
 *   - tracked-change: paragraph sits inside an existing <w:ins> / <w:del>.
 *     Mutating it would nest tracked changes; Word's review UI handles the
 *     nesting unevenly. Refuse.
 *   - field: paragraph is inside a complex field region (between
 *     <w:fldChar begin> and <w:fldChar end>). Editing field internals
 *     breaks STYLEREF / TOC / REF / cross-reference. Refuse.
 *   - sdt: paragraph is inside a <w:sdt> content control (block- or
 *     inline-level). Content controls have their own update semantics
 *     (see Phase 2 SDT-specialized path).
 *
 * Field-region detection runs a pseudo-state-machine over the body in
 * document order: we increment depth on `begin` and decrement on `end`,
 * marking every paragraph encountered while depth > 0. Field regions can
 * span paragraph boundaries — `STYLEREF` in particular often does.
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"

export type BlockerReason = "tracked-change" | "field" | "sdt"

export interface BlockerScan {
  byElement: Map<Element, BlockerReason>
  /** Same data keyed by 1-based paragraph index (only for indexed paragraphs;
   * data-table cell paragraphs are absent). The engine prefers element-
   * keyed lookup; the index map is for inspect_blockers' agent-facing report. */
  byIndex: Map<number, BlockerReason>
}

const w = NS.w

export function detectBlockers(
  documentDoc: Document,
  indexByElement: Map<Element, number>,
): BlockerScan {
  const byElement = new Map<Element, BlockerReason>()
  const byIndex = new Map<number, BlockerReason>()
  const root = documentDoc.documentElement
  if (!root) return { byElement, byIndex }
  const body = firstChildNS(root, NS.w, "body")
  if (!body) return { byElement, byIndex }

  // State across the entire body walk. Field nesting can span paragraphs;
  // ins/del/sdt are tracked per ancestor stack.
  let fieldDepth = 0

  const recordBlocker = (pEl: Element, reason: BlockerReason): void => {
    // First reason wins — once a paragraph is marked, additional reasons
    // don't refine the message. The engine refuses on any blocker, so the
    // specific reason is just for the report.
    if (byElement.has(pEl)) return
    byElement.set(pEl, reason)
    const idx = indexByElement.get(pEl)
    if (idx !== undefined) byIndex.set(idx, reason)
  }

  /** Walk the descendants of an element, updating fieldDepth on fldChar
   * tokens and marking any <w:p> ancestor of a tracked-change / sdt /
   * field-region run. The `inheritedReason` is whatever ancestor blocker
   * already applies to the current subtree. */
  const walk = (
    node: Element,
    currentParagraph: Element | null,
    inheritedReason: BlockerReason | null,
  ): void => {
    // If this is a paragraph and an ancestor blocker applies, record it.
    if (node.namespaceURI === w && node.localName === "p") {
      currentParagraph = node
      if (inheritedReason) recordBlocker(node, inheritedReason)
      // Field depth carries across paragraphs; if currently inside a field
      // region, mark this paragraph too.
      if (fieldDepth > 0) recordBlocker(node, "field")
    }

    // Tracked-change / sdt scopes: child paragraphs inside should be blocked.
    let scopeReason: BlockerReason | null = inheritedReason
    if (node.namespaceURI === w) {
      if (node.localName === "ins" || node.localName === "del") {
        scopeReason = "tracked-change"
        // If a paragraph already exists as ancestor, mark it now.
        if (currentParagraph) recordBlocker(currentParagraph, "tracked-change")
      } else if (node.localName === "sdt") {
        scopeReason = "sdt"
        if (currentParagraph) recordBlocker(currentParagraph, "sdt")
      } else if (node.localName === "fldChar") {
        const ftype = wAttr(node, "fldCharType")
        if (ftype === "begin") {
          fieldDepth++
          if (currentParagraph) recordBlocker(currentParagraph, "field")
        } else if (ftype === "end") {
          if (fieldDepth > 0) fieldDepth--
        }
        // separate token: inside region, no depth change
      }
    }

    for (const child of getChildren(node)) {
      walk(child, currentParagraph, scopeReason)
    }
  }

  walk(body, null, null)
  return { byElement, byIndex }
}

/** Reason → human-readable explanation for the agent. */
export function explainBlockerReason(reason: BlockerReason): string {
  switch (reason) {
    case "tracked-change":
      return "paragraph already contains tracked changes (existing <w:ins>/<w:del>) — edit refused to avoid nested revision markup"
    case "field":
      return "paragraph is inside a complex field region (TOC / STYLEREF / cross-reference) — editing field internals breaks the field"
    case "sdt":
      return "paragraph is inside a content control (<w:sdt>) — use the SDT specialized path (Phase 2; not yet available)"
  }
}

/** Tally counts per reason for inspect_blockers' summary line. */
export function summarizeBlockers(scan: BlockerScan): Record<BlockerReason, number> {
  const out: Record<BlockerReason, number> = {
    "tracked-change": 0,
    field: 0,
    sdt: 0,
  }
  for (const reason of scan.byElement.values()) out[reason]++
  return out
}

/** Re-export the namespace lookup helper used by callers that walk on top
 * of detectBlockers' output (e.g. checking adjacency of blocked regions). */
export { getChildrenNS }
