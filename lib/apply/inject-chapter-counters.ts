/**
 * Inject a hidden auto-chapter SEQ field at the start of each paragraph
 * whose style is referenced under `chapterPrefix` with a `format`
 * override. Caption-emit's chapter prefix `SEQ _chap_<styleId> \c \*
 * <FORMAT>` reads the counter via Word's `\c` (repeat) switch; the
 * injected field is what advances the counter.
 *
 * Why a parallel hidden counter instead of `STYLEREF "Heading 1" \n
 * \* ARABIC`: STYLEREF \n returns the heading's full rendered lvlText
 * ("第一章" for chineseCounting), and `\* ARABIC` does NOT extract the
 * numeric portion — Word renders "第一章.1" instead of "1.1". The
 * hidden SEQ maintains a parallel Arabic counter that Word's F9 keeps
 * live on H1 add / remove.
 *
 * Idempotent: skips paragraphs that already carry an `_chap_<styleId>`
 * SEQ field at the start (re-applying doesn't double-inject).
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, wAttr, walkBodyParagraphs } from "@lib/xml/xml-utils.ts"
import { parseFieldRuns } from "@lib/edit/fields/field-parse.ts"
import { emitSeqField, type SeqFormat } from "@lib/edit/fields/seq-field.ts"
import { chapterCounterIdentifier } from "@lib/edit/caption-emit.ts"
import type { ResolvedCaptionConfig } from "@lib/edit/caption-counter.ts"

const w = NS.w

export function injectChapterCounters(
  documentDoc: Document,
  captions: Map<string, ResolvedCaptionConfig>,
): void {
  // styleId → format. Only entries with a format override participate;
  // bare-string chapterPrefix entries use STYLEREF + heading's native
  // numbering and don't need a parallel counter.
  const tracked = new Map<string, SeqFormat>()
  for (const config of captions.values()) {
    for (const entry of config.chapterPrefix) {
      if (entry.format !== undefined) {
        const existing = tracked.get(entry.styleId)
        if (existing !== undefined && existing !== entry.format) {
          throw new Error(
            `captions conflict: styleId "${entry.styleId}" appears under chapterPrefix with both format "${existing}" and "${entry.format}". The hidden auto-chapter counter for one style must use a single format — split into two styleIds or unify the format.`,
          )
        }
        tracked.set(entry.styleId, entry.format)
      }
    }
  }
  if (tracked.size === 0) return

  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return

  for (const para of walkBodyParagraphs(body)) {
    const styleId = paragraphStyleId(para)
    if (!styleId) continue
    const format = tracked.get(styleId)
    if (format === undefined) continue
    if (paragraphHasChapterSeq(para, styleId)) continue
    injectHiddenSeq(para, styleId, format, documentDoc)
  }
}

function injectHiddenSeq(
  paragraph: Element,
  styleId: string,
  format: SeqFormat,
  ownerDoc: Document,
): void {
  // Emit the SEQ field WITHOUT \h. Word's documented \h-on-SEQ
  // semantics: "the `\h` switch does NOT hide the result when a `\*`
  // format switch is also present" (verified in Microsoft Q&A — both
  // `\*` and `\#` unhide a `\h` SEQ). Since we always emit `\* ARABIC`
  // (or whatever the format is), `\h` would silently fail and the
  // chapter number would render inside the H1.
  //
  // Use character-level `<w:vanish/>` rPr instead — robust against the
  // format-switch interaction, and Word's counter still advances since
  // the SEQ field itself runs (only its rendered text is hidden).
  const { runs } = emitSeqField(ownerDoc, {
    identifier: chapterCounterIdentifier(styleId),
    format,
  })
  for (const r of runs) addVanishRPr(r, ownerDoc)

  // Insert after pPr if present, otherwise at paragraph start. Each
  // insertBefore call adds the new node immediately before `anchorNode`;
  // walking forward keeps the field's begin / instr / separate /
  // result / end order intact in the DOM. (Reverse iteration here
  // would silently swap the order — fldChar end would land at the
  // front, Word reads it before begin, the whole field breaks.)
  const pPr = firstChildNS(paragraph, w, "pPr")
  const anchorNode = pPr ? pPr.nextSibling : paragraph.firstChild
  if (anchorNode) {
    for (const r of runs) paragraph.insertBefore(r, anchorNode)
  } else {
    for (const r of runs) paragraph.appendChild(r)
  }
}

/** Prepend `<w:vanish/>` to a run's rPr (creating rPr if absent) so
 * Word renders the run as hidden text — counter still advances, no
 * visible artifact in the paragraph. rPr must be the first child of
 * the run per CT_R schema order. */
function addVanishRPr(run: Element, ownerDoc: Document): void {
  let rPr = firstChildNS(run, w, "rPr")
  if (!rPr) {
    rPr = ownerDoc.createElementNS(w, "w:rPr")
    run.insertBefore(rPr, run.firstChild)
  }
  const existing = firstChildNS(rPr, w, "vanish")
  if (!existing) {
    rPr.insertBefore(ownerDoc.createElementNS(w, "w:vanish"), rPr.firstChild)
  }
}

function paragraphStyleId(paragraph: Element): string | undefined {
  const pPr = firstChildNS(paragraph, w, "pPr")
  if (!pPr) return undefined
  const pStyle = firstChildNS(pPr, w, "pStyle")
  if (!pStyle) return undefined
  return wAttr(pStyle, "val") ?? undefined
}

function paragraphHasChapterSeq(paragraph: Element, styleId: string): boolean {
  const expectedId = chapterCounterIdentifier(styleId)
  const runs: Element[] = []
  for (const c of getChildren(paragraph)) {
    if (c.namespaceURI === w && c.localName === "r") runs.push(c)
  }
  const parsed = parseFieldRuns(runs)
  for (const entry of parsed) {
    if (
      entry.kind === "field" &&
      entry.fieldType === "SEQ" &&
      entry.details.identifier === expectedId
    ) {
      return true
    }
  }
  return false
}
