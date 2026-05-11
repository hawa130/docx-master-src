/**
 * REF field emitter.
 *
 * Produces the 5-run sequence that OOXML uses for a complex `REF` field:
 *
 *   <w:r><w:fldChar fldCharType="begin"/></w:r>
 *   <w:r><w:instrText xml:space="preserve"> REF _RefXXX \\n \\h </w:instrText></w:r>
 *   <w:r><w:fldChar fldCharType="separate"/></w:r>
 *   <w:r>[ rPr from runFormat ]<w:t>placeholder</w:t></w:r>
 *   <w:r><w:fldChar fldCharType="end"/></w:r>
 *
 * The placeholder run carries the user-visible text Word displays BEFORE
 * field update; after F9 (or auto-update on open via `updateFields=true`
 * in settings.xml) Word replaces it with the bookmark's resolved content.
 * Computing the placeholder accurately at emit time means the doc reads
 * correctly even if a user opens it without triggering field update.
 *
 * The placeholder run accepts an optional `RunFormat` so the visible text
 * inherits color / italic / size etc. — same field path as inline text
 * runs.
 */

import { NS } from "@lib/parse/types.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"
import { buildRPrChildren } from "@lib/edit/fragment-emit.ts"

const w = NS.w

export interface RefFieldSpec {
  /** Bookmark name (without leading underscore prefix decoration — pass
   * the canonical name produced by BookmarkAllocator). */
  bookmarkName: string
  /** Switch list. Each entry already includes the leading backslash, e.g.
   * `"\\n"`, `"\\h"`. Order matches Word output convention but doesn't
   * affect resolution. */
  switches: string[]
  /** Text shown until Word updates fields. Empty string is legal — Word
   * still resolves at update time — but a value here means the document
   * reads correctly on first open even if field update is declined. */
  placeholder: string
  /** Optional rPr for the placeholder run. */
  format?: RunFormat
}

/** Emit the 5-run REF sequence. Returns the runs ready to insert into a
 * paragraph element (or any container that accepts `<w:r>` children). */
export function emitRefField(ownerDoc: Document, spec: RefFieldSpec): Element[] {
  const begin = ownerDoc.createElementNS(w, "w:r")
  const beginFld = ownerDoc.createElementNS(w, "w:fldChar")
  beginFld.setAttributeNS(w, "w:fldCharType", "begin")
  begin.appendChild(beginFld)

  const instr = ownerDoc.createElementNS(w, "w:r")
  const instrText = ownerDoc.createElementNS(w, "w:instrText")
  // `xml:space="preserve"` so the leading + trailing single spaces around
  // the field code don't get collapsed during XML serialization. Without
  // this, some Word builds (and stricter validators) fail to parse the
  // field code or render the literal text.
  instrText.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
  const switchPart = spec.switches.length > 0 ? ` ${spec.switches.join(" ")}` : ""
  instrText.textContent = ` REF ${spec.bookmarkName}${switchPart} `
  instr.appendChild(instrText)

  const separate = ownerDoc.createElementNS(w, "w:r")
  const sepFld = ownerDoc.createElementNS(w, "w:fldChar")
  sepFld.setAttributeNS(w, "w:fldCharType", "separate")
  separate.appendChild(sepFld)

  const placeholderRun = ownerDoc.createElementNS(w, "w:r")
  if (spec.format) {
    const rPr = ownerDoc.createElementNS(w, "w:rPr")
    for (const c of buildRPrChildren(spec.format, ownerDoc)) rPr.appendChild(c)
    if (rPr.childNodes.length > 0) placeholderRun.appendChild(rPr)
  }
  const placeholderT = ownerDoc.createElementNS(w, "w:t")
  placeholderT.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
  placeholderT.textContent = spec.placeholder
  placeholderRun.appendChild(placeholderT)

  const end = ownerDoc.createElementNS(w, "w:r")
  const endFld = ownerDoc.createElementNS(w, "w:fldChar")
  endFld.setAttributeNS(w, "w:fldCharType", "end")
  end.appendChild(endFld)

  return [begin, instr, separate, placeholderRun, end]
}

/** Map a `display` value to the corresponding REF switch list.
 *   - "label"  → `\n` (full numbered paragraph text resolved from lvlText,
 *                e.g. "图 1" or "1.2.3") + `\h` (hyperlink)
 *   - "number" → `\r` (paragraph number in relative context — for single-
 *                level schemes equals the counter "1"; for multi-level
 *                schemes returns relative position, not the leaf digit) + `\h`
 *   - "full"   → no switches (bookmark text content, e.g. caption title) + `\h`
 *
 * `\h` is always added so the rendered text is clickable in Word — matches
 * what the UI's Insert → Cross-reference produces. */
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
  /** Element ref to the target paragraph — looked up against the counter
   * simulator's output map. */
  targetParagraph: Element
  /** Which rendered field to use as the placeholder. */
  display: "full" | "label" | "number"
}
