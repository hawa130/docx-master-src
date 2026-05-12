/**
 * STYLEREF field emitter — chapter-prefix lookup.
 *
 *   { STYLEREF "<styleName>" \n [\* MERGEFORMAT] }
 *
 * Resolves to the rendered content of the nearest preceding paragraph
 * with the given style. Paired with SEQ in caption layouts to produce
 * chapter-prefixed numbers like "2.3" — STYLEREF "Heading 1" \n returns
 * the H1's paragraph number ("2"); SEQ Equation \s 1 returns the
 * equation counter within that chapter ("3"); a literal "." run joins
 * them.
 *
 * The `\n` switch (paragraph number) is the canonical caption switch —
 * returns just the numbered prefix without the heading text. Other
 * switches are available for non-caption uses (`\l` lower formatting,
 * `\p` position, etc.) but the caption pipeline only uses `\n`.
 *
 * styleName: Word's display name (the `<w:name w:val="..."/>` value in
 * styles.xml), NOT the styleId. The engine resolves styleId → name at
 * resolution time. STYLEREF requires double-quoted style names —
 * required for names with spaces ("Heading 1"), harmless for names
 * without. The emitter always quotes.
 *
 * Format preservation (rPr replication + MERGEFORMAT) handled by the
 * shared skeleton — see `complex-field.ts`.
 */

import { emitComplexField } from "@lib/edit/fields/complex-field.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"

export interface StyleRefFieldSpec {
  /** Style display name (NOT styleId). Resolved upstream from styleId
   * via styles.xml's `<w:name w:val="..."/>`. */
  styleName: string
  /** Switch list. Each entry already includes the leading backslash, e.g.
   * `"\\n"`. Typically `["\\n"]` for caption chapter prefix. */
  switches: string[]
  /** Initial result text shown before Word updates fields. Empty by
   * default — counter sim backfills the resolved chapter number. */
  initialResult?: string
  /** Optional rPr for the field runs (format-bearing captions). */
  formatRPr?: RunFormat
}

/** Emit the 5-run STYLEREF sequence. Returns the runs plus the result
 * text element for counter-sim backfill. */
export function emitStyleRefField(
  ownerDoc: Document,
  spec: StyleRefFieldSpec,
): { runs: Element[]; resultTextEl: Element } {
  const switchPart = spec.switches.length > 0 ? ` ${spec.switches.join(" ")}` : ""
  const { runs, resultTextEl } = emitComplexField(ownerDoc, {
    instrCode: `STYLEREF "${spec.styleName}"${switchPart}`,
    initialResult: spec.initialResult,
    format: spec.formatRPr,
  })
  return { runs, resultTextEl }
}
