/**
 * Shared 5-run skeleton for OOXML complex fields.
 *
 * Every complex field (REF / SEQ / STYLEREF / ...) renders as:
 *
 *   <w:r>[ rPr? ]<w:fldChar fldCharType="begin"/></w:r>
 *   <w:r>[ rPr? ]<w:instrText xml:space="preserve"> <code> </w:instrText></w:r>
 *   <w:r>[ rPr? ]<w:fldChar fldCharType="separate"/></w:r>
 *   <w:r>[ rPr? ]<w:t>[result]</w:t></w:r>
 *   <w:r>[ rPr? ]<w:fldChar fldCharType="end"/></w:r>
 *
 * The result run carries the user-visible text Word displays BEFORE field
 * update; after F9 (or auto-update on open via `updateFields=true` in
 * settings.xml) Word replaces it with the resolved value. Field-type-
 * specific emitters supply `instrCode` (without surrounding spaces) and an
 * optional `initialResult`; the skeleton handles the run layout, the
 * `xml:space="preserve"` wrapping, and the format-preservation contract.
 *
 * `format`: optional `RunFormat`. When non-empty, `applyFieldFormat` (1)
 * replicates the same rPr onto every field run and (2) appends
 * `\* MERGEFORMAT` to the field code. Both are required together — Word's
 * field-update logic rewrites the result run on F9 and the new run
 * inherits rPr from the surrounding field runs, not from the prior result
 * run we wrote; MERGEFORMAT tells Word to preserve that rPr across
 * subsequent updates. Either alone is insufficient.
 */

import { NS } from "@lib/parse/types.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"
import { buildRPrChildren } from "@lib/edit/fragment-emit.ts"

const w = NS.w
const XML_NS = "http://www.w3.org/XML/1998/namespace"

export interface ComplexFieldSpec {
  /** Field code body without surrounding spaces and without the
   * `\* MERGEFORMAT` suffix — the skeleton appends both. Example:
   * `"REF _Ref12345 \\n \\h"`, `"SEQ Equation \\* ARABIC \\s 1"`. */
  instrCode: string
  /** Text shown until Word updates fields. Empty string is legal — Word
   * still resolves at update time — but a value here means the document
   * reads correctly on first open even if field update is declined. */
  initialResult?: string
  /** Optional rPr for the field runs. When non-empty, the skeleton
   * replicates the rPr across every run AND appends `\\* MERGEFORMAT` to
   * the instrText. Empty / undefined → no format preservation, no
   * MERGEFORMAT. */
  format?: RunFormat
}

export interface EmittedComplexField {
  /** The 5 runs in document order: begin, instr, separate, result, end.
   * Insert this array into a paragraph (or any container accepting
   * `<w:r>` children). */
  runs: Element[]
  /** The result run (index 3 in `runs`). */
  resultRun: Element
  /** The `<w:t>` inside the result run. textContent is set during emit;
   * callers that backfill placeholder text after counter simulation
   * mutate this element directly. */
  resultTextEl: Element
}

/** Emit a complex field's 5-run sequence. Returns the runs plus direct
 * references to the result run and its text element for callers that
 * need to backfill the placeholder later (e.g. REF after counter sim). */
export function emitComplexField(ownerDoc: Document, spec: ComplexFieldSpec): EmittedComplexField {
  const begin = ownerDoc.createElementNS(w, "w:r")
  const beginFld = ownerDoc.createElementNS(w, "w:fldChar")
  beginFld.setAttributeNS(w, "w:fldCharType", "begin")
  begin.appendChild(beginFld)

  const instr = ownerDoc.createElementNS(w, "w:r")
  const instrText = ownerDoc.createElementNS(w, "w:instrText")
  // `xml:space="preserve"` keeps the leading + trailing single spaces
  // around the field code from collapsing during XML serialization.
  // Without this, some Word builds and stricter validators fail to parse
  // the field code or render the literal text.
  instrText.setAttributeNS(XML_NS, "xml:space", "preserve")
  instr.appendChild(instrText)

  const separate = ownerDoc.createElementNS(w, "w:r")
  const sepFld = ownerDoc.createElementNS(w, "w:fldChar")
  sepFld.setAttributeNS(w, "w:fldCharType", "separate")
  separate.appendChild(sepFld)

  const resultRun = ownerDoc.createElementNS(w, "w:r")
  const resultTextEl = ownerDoc.createElementNS(w, "w:t")
  resultTextEl.setAttributeNS(XML_NS, "xml:space", "preserve")
  resultTextEl.textContent = spec.initialResult ?? ""
  resultRun.appendChild(resultTextEl)

  const end = ownerDoc.createElementNS(w, "w:r")
  const endFld = ownerDoc.createElementNS(w, "w:fldChar")
  endFld.setAttributeNS(w, "w:fldCharType", "end")
  end.appendChild(endFld)

  // Format preservation: rPr replication + MERGEFORMAT switch (bundled).
  // Run BEFORE setting instrText.textContent so the suffix can land on
  // the same instrText run; rPr is inserted as the FIRST child of each
  // run per CT_R schema order (rPr must precede content children).
  const { instrTextSuffix } = applyFieldFormat(
    [begin, instr, separate, resultRun, end],
    spec.format,
    ownerDoc,
  )
  instrText.textContent = ` ${spec.instrCode}${instrTextSuffix} `

  return { runs: [begin, instr, separate, resultRun, end], resultRun, resultTextEl }
}

/** Replicate agent-declared rPr across every run + emit `\\* MERGEFORMAT`
 * iff the format produces non-empty rPr. Empty format (`{}` /
 * `{ color: undefined }`) treated as no-format — matches inheritance
 * semantics in the edit engine.
 *
 * Returns the field-code suffix to append (empty when no format). */
function applyFieldFormat(
  fieldRuns: readonly Element[],
  format: RunFormat | undefined,
  ownerDoc: Document,
): { instrTextSuffix: string } {
  if (!format) return { instrTextSuffix: "" }
  const rPr = ownerDoc.createElementNS(w, "w:rPr")
  for (const c of buildRPrChildren(format, ownerDoc)) rPr.appendChild(c)
  if (rPr.childNodes.length === 0) return { instrTextSuffix: "" }
  for (const r of fieldRuns) {
    r.insertBefore(rPr.cloneNode(true), r.firstChild)
  }
  return { instrTextSuffix: " \\* MERGEFORMAT" }
}
