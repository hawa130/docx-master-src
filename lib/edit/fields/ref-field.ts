/**
 * REF field emitter — thin layer over the shared complex-field skeleton.
 *
 * REF resolves a bookmark at field-update time and renders the result
 * per its switches (`\n` lvlText / `\r` paragraph number / `\h` hyperlink
 * / etc.). The placeholder run carries the user-visible text Word
 * displays BEFORE update; after F9 (or auto-update on open via
 * `updateFields=true` in settings.xml) Word replaces it with the
 * bookmark's resolved content. Computing the placeholder accurately at
 * emit time means the document reads correctly even if a user opens it
 * without triggering field update.
 *
 * Format preservation (rPr replication + MERGEFORMAT) is handled by the
 * shared skeleton — see `complex-field.ts`.
 */

import { emitComplexField } from "@lib/edit/fields/complex-field.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"

export interface RefFieldSpec {
  /** Bookmark name. Pass the canonical name produced by
   * BookmarkAllocator (no leading underscore decoration). */
  bookmarkName: string
  /** Switch list. Each entry already includes the leading backslash, e.g.
   * `"\\n"`, `"\\h"`. Order matches Word output convention but doesn't
   * affect resolution. */
  switches: string[]
  /** Text shown until Word updates fields. Empty string is legal — Word
   * still resolves at update time. */
  placeholder: string
  /** Optional rPr for the field runs, replicated across every run so
   * Word preserves the formatting across updates. */
  format?: RunFormat
}

/** Emit the 5-run REF sequence. Returns the runs plus the result text
 * element so callers can backfill the placeholder after counter
 * simulation. */
export function emitRefField(
  ownerDoc: Document,
  spec: RefFieldSpec,
): { runs: Element[]; resultTextEl: Element } {
  const switchPart = spec.switches.length > 0 ? ` ${spec.switches.join(" ")}` : ""
  const { runs, resultTextEl } = emitComplexField(ownerDoc, {
    instrCode: `REF ${spec.bookmarkName}${switchPart}`,
    initialResult: spec.placeholder,
    format: spec.format,
  })
  return { runs, resultTextEl }
}

/** Map a `display` value to the corresponding REF switch list.
 *   - "label"  → `\n` (full numbered paragraph text from lvlText,
 *                e.g. "图 1" or "1.2.3") + `\h` (hyperlink)
 *   - "number" → `\r` (paragraph number in relative context — for single-
 *                level schemes equals the counter "1"; for multi-level
 *                returns relative position, not the leaf digit) + `\h`
 *   - "full"   → no number switches (bookmark text content, e.g. caption
 *                title) + `\h`
 *
 * `\h` is always added so the rendered text is clickable in Word — matches
 * what Insert → Cross-reference produces. */
export function switchesForDisplay(display: "full" | "label" | "number"): string[] {
  switch (display) {
    case "label":
      return ["\\n", "\\h"]
    case "number":
      return ["\\r", "\\h"]
    case "full":
      return ["\\h"]
  }
}

/** A pending REF whose placeholder text must be backfilled after numbering
 * counters are simulated. The engine stages these during edits and the
 * apply orchestrator commits them once the post-edit counter pass yields
 * resolved label/number/full strings per target paragraph. */
export interface PendingRefBackfill {
  /** The `<w:t>` element inside the placeholder run. textContent is set
   * during commit. */
  placeholderTextEl: Element
  /** Bookmark name. The backfill consumer resolves this against the
   * allocator to recover the target element for the counter-simulator
   * lookup. Pre-scan guarantees the name is registered before emit; a
   * resolution failure here is an engine bug, not an agent input error. */
  targetName: string
  /** Which rendered field to use as the placeholder. */
  display: "full" | "label" | "number"
}
