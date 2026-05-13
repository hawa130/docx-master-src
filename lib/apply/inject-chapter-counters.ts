/**
 * Inject a hidden auto-chapter SEQ field into each paragraph whose
 * style is referenced under `chapterPrefix` with a `format` override.
 * Caption-emit's chapter prefix `SEQ _chap_<styleId> \c \* <FORMAT>`
 * reads the counter via Word's `\c` (repeat) switch; the injected
 * field is what advances the counter on each heading.
 *
 * Why a parallel hidden counter instead of `STYLEREF "Heading 1" \n
 * \* ARABIC`: STYLEREF \n returns the heading's full rendered lvlText
 * ("第一章" for chineseCounting), and `\* ARABIC` does NOT extract the
 * numeric portion — Word renders "第一章.1" instead of "1.1". The
 * hidden SEQ maintains a parallel Arabic counter that Word's F9 keeps
 * live on H1 add / remove.
 *
 * Field placement: appended at the END of the heading paragraph. The
 * heading's own auto-numbering counter (e.g. lvlText "第%1章") renders
 * at the paragraph's visual start; a SEQ field at that position
 * conflicts with the lvlText layout slot in some Word versions and
 * suppresses the heading's prefix. Tail placement still advances the
 * counter (Word evaluates fields document-wide on F9) and the cross-
 * paragraph SEQ \c ordering is preserved (headings precede captions
 * in body order).
 *
 * Hiding: character-level `<w:vanish/>` rPr on each run (via
 * `addVanishRPr`). Word's SEQ `\h` switch is silently overridden by a
 * `\*` format switch in the same field — and we always emit `\*` —
 * so the rPr-level vanish is the reliable hide mechanism.
 *
 * Idempotent: skips paragraphs that already carry the matching
 * `_chap_<styleId>` SEQ field anywhere in their run sequence.
 */

import { NS } from "@lib/parse/types.ts"
import {
  addVanishRPr,
  firstChildNS,
  getChildren,
  wAttr,
  walkBodyParagraphs,
} from "@lib/xml/xml-utils.ts"
import { parseFieldRuns } from "@lib/edit/fields/field-parse.ts"
import { emitSeqField, type SeqFormat } from "@lib/edit/fields/seq-field.ts"
import { chapterCounterIdentifier } from "@lib/edit/caption-emit.ts"
import type { ResolvedCaptionConfig } from "@lib/edit/caption-counter.ts"

const w = NS.w

/** Returns the count of paragraphs that received a fresh hidden SEQ
 * (paragraphs that already carried one are skipped). The count surfaces
 * in the dry-run change report so the agent can preview chapter-counter
 * coverage before committing. */
export function injectChapterCounters(
  documentDoc: Document,
  captions: Map<string, ResolvedCaptionConfig>,
): number {
  // styleId → format. Only entries with a format override participate;
  // bare-string chapterPrefix entries use STYLEREF + heading's native
  // numbering and don't need a parallel counter. (caption-resolver pre-
  // checks cross-entry consistency so the throw below is a defensive
  // engine-layer guard — every caller routes through resolveCaptions.)
  const tracked = new Map<string, { format: SeqFormat; captionIdentifier: string }>()
  for (const [captionIdentifier, config] of captions) {
    for (const entry of config.chapterPrefix) {
      if (entry.format !== undefined) {
        const existing = tracked.get(entry.styleId)
        if (existing !== undefined && existing.format !== entry.format) {
          throw new Error(
            `captions conflict: chapterPrefix styleId "${entry.styleId}" appears with format "${existing.format}" (captions["${existing.captionIdentifier}"]) and "${entry.format}" (captions["${captionIdentifier}"]). The hidden auto-chapter counter for one style must use a single format — split into two styleIds or unify the format.`,
          )
        }
        tracked.set(entry.styleId, { format: entry.format, captionIdentifier })
      }
    }
  }
  if (tracked.size === 0) return 0

  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return 0

  let injected = 0
  for (const para of walkBodyParagraphs(body)) {
    const styleId = paragraphStyleId(para)
    if (!styleId) continue
    const entry = tracked.get(styleId)
    if (entry === undefined) continue
    if (paragraphHasChapterSeq(para, styleId)) continue
    injectHiddenSeq(para, styleId, entry.format, documentDoc)
    injected++
  }
  return injected
}

function injectHiddenSeq(
  paragraph: Element,
  styleId: string,
  format: SeqFormat,
  ownerDoc: Document,
): void {
  const { runs } = emitSeqField(ownerDoc, {
    identifier: chapterCounterIdentifier(styleId),
    format,
  })
  for (const r of runs) {
    addVanishRPr(r, ownerDoc)
    paragraph.appendChild(r)
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
