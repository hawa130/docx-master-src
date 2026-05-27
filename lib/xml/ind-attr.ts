/**
 * Shared <w:ind> attribute emission. The same per-slot write logic is
 * needed by both style-level field-merge (lib/apply/style-mutation.ts —
 * mutates existing <w:ind> in a style's <w:pPr>) and direct-format
 * emission (lib/edit/fragment-emit.ts — builds fresh <w:ind> for an
 * inserted / replaced / format-op paragraph). Centralizing here:
 *
 *   - keeps the char-vs-twip attribute-name choice in one place
 *     (`Nchar` → `w:<slot>Chars`, twips → `w:<slot>`)
 *   - guarantees explicit zero is written, not skipped — `parseIndent(0)`
 *     returns `{ kind: "twip", value: 0 }`, distinct from `null` (which
 *     means "field absent, inherit"). Skipping `value === 0` silently
 *     drops a meaningful override (config-schema.md §indents semantics).
 */

import type { IndentParsed } from "@lib/shared/units.ts"
import { NS } from "@lib/parse/types.ts"

const w = NS.w

export type IndentSlot = "firstLine" | "hanging" | "left" | "right"

const FIRST_LINE_GROUP = ["firstLine", "firstLineChars", "hanging", "hangingChars"] as const

/** Write a parsed indent value onto a <w:ind> element.
 *  Emits even when `parsed.value === 0` — explicit zero overrides the
 *  style cascade. Callers gate on `parsed !== null` themselves
 *  (a `null` parse result means the source field was absent). */
export function setIndentAttr(ind: Element, slot: IndentSlot, parsed: IndentParsed): void {
  const attr = parsed.kind === "char" ? `w:${slot}Chars` : `w:${slot}`
  ind.setAttributeNS(w, attr, String(parsed.value))
}

/** Remove all attrs in the firstLine / hanging mutually-exclusive group.
 *  In OOXML these four attrs are pairwise exclusive (same offset, different
 *  units; opposing directions). Field-merge into an existing <w:ind> must
 *  clear the group before writing the new value, otherwise stale attrs
 *  from a prior mutation produce ambiguous Word behavior. */
export function clearFirstLineHangingGroup(ind: Element): void {
  for (const a of FIRST_LINE_GROUP) ind.removeAttributeNS(w, a)
}
