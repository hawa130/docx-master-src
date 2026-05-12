/**
 * REF field emitter.
 *
 * Produces the 5-run sequence that OOXML uses for a complex `REF` field:
 *
 *   <w:r>[ rPr? ]<w:fldChar fldCharType="begin"/></w:r>
 *   <w:r>[ rPr? ]<w:instrText xml:space="preserve"> REF _RefXXX \\n \\h [\\* MERGEFORMAT] </w:instrText></w:r>
 *   <w:r>[ rPr? ]<w:fldChar fldCharType="separate"/></w:r>
 *   <w:r>[ rPr? ]<w:t>placeholder</w:t></w:r>
 *   <w:r>[ rPr? ]<w:fldChar fldCharType="end"/></w:r>
 *
 * The placeholder run carries the user-visible text Word displays BEFORE
 * field update; after F9 (or auto-update on open via `updateFields=true`
 * in settings.xml) Word replaces it with the bookmark's resolved content.
 * Computing the placeholder accurately at emit time means the doc reads
 * correctly even if a user opens it without triggering field update.
 *
 * The placeholder run accepts an optional `RunFormat`. When that format
 * produces non-empty rPr, the helper `applyFieldFormat` (1) replicates
 * the same rPr onto every field run and (2) appends `\\* MERGEFORMAT` to
 * the field code. Both are required: Word's field-update logic rewrites
 * the result run on F9, and the new run inherits rPr from the surrounding
 * field runs — not from the prior placeholder. Without both ingredients
 * the format silently disappears after the first update. See
 * `cross-references.md` "Verifying format-bearing refs".
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
  /** Optional rPr for the placeholder run, replicated across every field
   * run so Word preserves the formatting across updates. */
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
  instr.appendChild(instrText)

  const separate = ownerDoc.createElementNS(w, "w:r")
  const sepFld = ownerDoc.createElementNS(w, "w:fldChar")
  sepFld.setAttributeNS(w, "w:fldCharType", "separate")
  separate.appendChild(sepFld)

  const placeholderRun = ownerDoc.createElementNS(w, "w:r")
  const placeholderT = ownerDoc.createElementNS(w, "w:t")
  placeholderT.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
  placeholderT.textContent = spec.placeholder
  placeholderRun.appendChild(placeholderT)

  const end = ownerDoc.createElementNS(w, "w:r")
  const endFld = ownerDoc.createElementNS(w, "w:fldChar")
  endFld.setAttributeNS(w, "w:fldCharType", "end")
  end.appendChild(endFld)

  // Format preservation: rPr replication + MERGEFORMAT switch (bundled —
  // both are required, see module header). Run *before* setting
  // instrText.textContent so the suffix can land on the same instrText
  // run; rPr is inserted as the FIRST child of each run per CT_R schema
  // order (rPr must precede content children).
  const { instrTextSuffix } = applyFieldFormat(
    [begin, instr, separate, placeholderRun, end],
    spec.format,
    ownerDoc,
  )
  const switchPart = spec.switches.length > 0 ? ` ${spec.switches.join(" ")}` : ""
  instrText.textContent = ` REF ${spec.bookmarkName}${switchPart}${instrTextSuffix} `

  return [begin, instr, separate, placeholderRun, end]
}

/** Replicate the agent-declared rPr across every run in a complex field's
 * begin / instrText / separate / result / end sequence, and emit the
 * `\\* MERGEFORMAT` switch when there's actually formatting to preserve.
 *
 * Why both must move together: Word's field-update logic rewrites the
 * result run on F9 / updateFields-on-open. The new run inherits rPr from
 * the surrounding field runs (not from the prior result run we wrote);
 * MERGEFORMAT then tells Word to keep that rPr across subsequent updates.
 * Either ingredient alone is insufficient — together they make agent-
 * declared formatting survive Word's update cycle.
 *
 * Returns the field-code suffix to append (empty when no format produces
 * children — preserves the no-format path exactly).
 *
 * This function is private to field-ref.ts today but the contract is
 * field-type-agnostic: any complex field that uses
 * fldChar=begin/separate/end + instrText will need the same treatment
 * to preserve format across updates. When a second field type emitter
 * appears (PAGEREF / HYPERLINK / STYLEREF / ...), promote this to a
 * shared module — the function shape is already correct for that. */
function applyFieldFormat(
  fieldRuns: readonly Element[],
  format: RunFormat | undefined,
  ownerDoc: Document,
): { instrTextSuffix: string } {
  if (!format) return { instrTextSuffix: "" }
  const rPr = ownerDoc.createElementNS(w, "w:rPr")
  for (const c of buildRPrChildren(format, ownerDoc)) rPr.appendChild(c)
  // `format: {}` (truthy but empty) and `format: { color: undefined }`
  // both end up here with no rPr children — treat as no-format. Matches
  // existing "ref.format ?? defaultFormat" inheritance semantics in the
  // edit engine: declaring an empty object shouldn't trigger
  // MERGEFORMAT, since there's nothing to merge.
  if (rPr.childNodes.length === 0) return { instrTextSuffix: "" }
  for (const r of fieldRuns) {
    // rPr must be the FIRST child of <w:r> per CT_R schema order.
    // `insertBefore(..., r.firstChild)` puts it at index 0 regardless of
    // whether the run already has content (it does — fldChar / instrText /
    // t are appended earlier in emitRefField).
    r.insertBefore(rPr.cloneNode(true), r.firstChild)
  }
  return { instrTextSuffix: " \\* MERGEFORMAT" }
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
   * simulator's output map. `null` when the ref was forward (target's
   * paragraph hadn't emitted yet at the time we wrote the placeholder);
   * the backfill consumer resolves via `targetName` against the now-fully-
   * populated allocator. */
  targetParagraph: Element | null
  /** Anchor name used for late-resolution when `targetParagraph` is null.
   * Always set so the consumer can produce an actionable error if the
   * anchor disappeared somehow (it shouldn't — pre-scan guards that). */
  targetName: string
  /** Which rendered field to use as the placeholder. */
  display: "full" | "label" | "number"
}
