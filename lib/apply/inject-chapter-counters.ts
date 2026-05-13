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
  const { runs } = emitSeqField(ownerDoc, {
    identifier: chapterCounterIdentifier(styleId),
    format,
    hidden: true,
  })
  // Insert after pPr if present, otherwise at paragraph start. Order
  // among runs within the paragraph matters for SEQ — the hidden field
  // must appear BEFORE any caption that reads it via \c (\c shows the
  // most recently advanced value), but within the same paragraph the
  // injection order doesn't matter for that ordering — captions are in
  // later paragraphs anyway.
  const pPr = firstChildNS(paragraph, w, "pPr")
  const insertBefore = pPr ? pPr.nextSibling : paragraph.firstChild
  for (let i = runs.length - 1; i >= 0; i--) {
    paragraph.insertBefore(runs[i]!, insertBefore)
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
