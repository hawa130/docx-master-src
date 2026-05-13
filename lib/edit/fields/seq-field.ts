/**
 * SEQ field emitter — caption-class auto-numbering.
 *
 *   { SEQ <identifier> \* <FORMAT> [\s <N>] [\r <N>] [\c] [\h] [\* MERGEFORMAT] }
 *
 * Each `SEQ <id>` occurrence increments the document-local counter for
 * that identifier. Different identifiers track independent counters —
 * `SEQ Figure` and `SEQ Equation` don't interfere. This is Word's native
 * caption mechanism (References → Insert Caption produces SEQ).
 *
 * Switches:
 *   `\* FORMAT`     numeric format for the rendered counter (ARABIC,
 *                   alphabetic, ALPHABETIC, roman, ROMAN, CHINESENUM2,
 *                   CHINESENUM3). Always present.
 *   `\s N`          restart the counter at each occurrence of a paragraph
 *                   at outline level N (1-indexed, matches Word UI's
 *                   "Heading 1" = level 1). Paired with a STYLEREF chapter
 *                   prefix in the caption's run sequence.
 *   `\r N`          reset the counter to N at this field (used by
 *                   CaptionCounterReset). Mutually exclusive with `\c`.
 *   `\c`            repeat the previous value of this identifier without
 *                   incrementing (used by subequation "continue" members
 *                   to keep the parent number stable across (1a)(1b)).
 *   `\h`            hide the field result — Word renders nothing. Used
 *                   for CaptionCounterReset's marker field whose only
 *                   purpose is to advance counter state.
 *
 * Format preservation (rPr replication + MERGEFORMAT) handled by the
 * shared skeleton — see `complex-field.ts`.
 */

import { emitComplexField } from "@lib/edit/fields/complex-field.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"

export type SeqFormat =
  | "arabic"
  | "alphabetic"
  | "ALPHABETIC"
  | "roman"
  | "ROMAN"
  | "chinese"
  | "chinese-formal"

/** Mapping from agent-facing format string to Word's `\*` switch token.
 * Engine emits these verbatim in instrText. */
export const FORMAT_SWITCH: Record<SeqFormat, string> = {
  arabic: "ARABIC",
  alphabetic: "alphabetic",
  ALPHABETIC: "ALPHABETIC",
  roman: "roman",
  ROMAN: "ROMAN",
  chinese: "CHINESENUM2",
  "chinese-formal": "CHINESENUM3",
}

export interface SeqFieldSpec {
  /** Counter identifier. Free string — Word uses it as the document-local
   * counter key. Conventional values: "Equation", "Figure", "Table",
   * "Theorem", etc. */
  identifier: string
  /** Numeric format for the rendered counter. */
  format: SeqFormat
  /** Outline level (1-9, 1-indexed) at which the counter resets. Omitted
   * → counter doesn't reset. Derived from the caption's chapterPrefix
   * last-entry outline level. */
  restartAtOutlineLevel?: number
  /** Reset the counter to this value at this field. Mutually exclusive
   * with `repeat`. Used by CaptionCounterReset. */
  resetTo?: number
  /** Repeat the previous counter value without incrementing (`\c`
   * switch). Mutually exclusive with `resetTo`. Used by subequation
   * "continue" members. */
  repeat?: boolean
  /** Hide the field result (`\h` switch). Word renders nothing — used
   * for CaptionCounterReset's marker field. */
  hidden?: boolean
  /** Initial result text shown before Word updates fields. Empty by
   * default — counter sim backfills the rendered value later. */
  initialResult?: string
  /** Optional rPr for the field runs (format-bearing captions). See
   * `complex-field.ts` for replication semantics. */
  formatRPr?: RunFormat
}

/** Emit the 5-run SEQ sequence. Returns the runs plus the result text
 * element for counter-sim backfill. */
export function emitSeqField(
  ownerDoc: Document,
  spec: SeqFieldSpec,
): { runs: Element[]; resultTextEl: Element } {
  if (spec.repeat && spec.resetTo !== undefined) {
    throw new Error(
      `SEQ field "${spec.identifier}": \\c (repeat) and \\r (reset) are mutually exclusive; pass one or the other.`,
    )
  }

  const parts: string[] = [`SEQ ${spec.identifier}`, `\\* ${FORMAT_SWITCH[spec.format]}`]
  if (spec.restartAtOutlineLevel !== undefined) {
    parts.push(`\\s ${spec.restartAtOutlineLevel}`)
  }
  if (spec.resetTo !== undefined) {
    parts.push(`\\r ${spec.resetTo}`)
  }
  if (spec.repeat) {
    parts.push("\\c")
  }
  if (spec.hidden) {
    parts.push("\\h")
  }

  const { runs, resultTextEl } = emitComplexField(ownerDoc, {
    instrCode: parts.join(" "),
    initialResult: spec.initialResult,
    format: spec.formatRPr,
  })
  return { runs, resultTextEl }
}
