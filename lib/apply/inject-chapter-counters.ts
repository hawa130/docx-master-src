/**
 * For each paragraph whose style is referenced under `chapterPrefix`
 * with a `format` override, emit a SIBLING engine-managed paragraph
 * immediately after it. The sibling carries the hidden auto-chapter
 * `SEQ _chap_<styleId> \* <FORMAT>` field; caption-emit's chapter prefix
 * `SEQ _chap_<styleId> \c \* <FORMAT>` reads the counter via Word's `\c`
 * (repeat) switch.
 *
 * Why a parallel hidden counter instead of `STYLEREF "Heading 1" \n
 * \* ARABIC`: STYLEREF \n returns the heading's full rendered lvlText
 * ("第一章" for chineseCounting), and `\* ARABIC` does NOT extract the
 * numeric portion — Word renders "第一章.1" instead of "1.1". The
 * hidden SEQ maintains a parallel Arabic counter that Word's F9 keeps
 * live on H1 add / remove.
 *
 * Why a sibling paragraph instead of inline at the heading's tail:
 * STYLEREF "heading 1" \* MERGEFORMAT reads the H1 paragraph's full
 * text content (including hidden runs) and reformats with the field's
 * own rPr — stripping vanish. A trailing hidden SEQ inside H1 then
 * surfaces as visible numeric pollution in any header that styleref-
 * references the heading. Moving the SEQ to a separate paragraph keeps
 * H1's text content clean for STYLEREF / TOC / REF / nav-pane extraction
 * while preserving document-order SEQ increment for captions.
 *
 * Sibling layout: a new `<w:p>` with `_HiddenChapterCounter` pStyle,
 * paragraph-mark vanish (pPr.rPr.vanish) so the paragraph collapses to
 * zero height when hidden text isn't shown, and the same 5-run SEQ
 * skeleton with vanish on each run. The `_`-prefix styleId is the
 * engine-managed-paragraph convention — `walkIndexedParagraphs` and the
 * `DocumentParser` skip these so agent-facing indices match user-visible
 * content order. See CLAUDE.md "Mechanical correctness".
 */

import { NS } from "@lib/parse/types.ts"
import {
  addVanishRPr,
  firstChildNS,
  getChildrenNS,
  paragraphStyleId,
  wAttr,
  walkBodyParagraphs,
} from "@lib/xml/xml-utils.ts"
import { emitSeqField, type SeqFormat } from "@lib/edit/fields/seq-field.ts"
import { chapterCounterIdentifier } from "@lib/edit/caption-emit.ts"
import type { ResolvedCaptionConfig } from "@lib/edit/caption-counter.ts"

const w = NS.w

/** styleId used for engine-managed sibling counter paragraphs. The
 *  leading underscore flags it as engine-internal — walkers and the
 *  parser skip `_`-prefix styleIds so the agent-visible paragraph index
 *  reflects content order, not raw DOM order. */
export const HIDDEN_CHAPTER_COUNTER_STYLE_ID = "_HiddenChapterCounter"

/** Returns the count of sibling counter paragraphs newly emitted. The
 * count surfaces in the dry-run change report so the agent can preview
 * chapter-counter coverage before committing. */
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

  // Collect targets first; mutating the DOM (inserting siblings) during
  // iteration confuses the recursive walker.
  const targets: { para: Element; styleId: string; format: SeqFormat }[] = []
  for (const para of walkBodyParagraphs(body)) {
    const styleId = paragraphStyleId(para)
    if (!styleId) continue
    const entry = tracked.get(styleId)
    if (entry === undefined) continue
    targets.push({ para, styleId, format: entry.format })
  }

  for (const { para, styleId, format } of targets) {
    emitCounterSibling(para, styleId, format, documentDoc)
  }
  return targets.length
}

/** Insert an engine-managed paragraph holding the hidden SEQ field
 *  immediately after `heading`. */
function emitCounterSibling(
  heading: Element,
  styleId: string,
  format: SeqFormat,
  ownerDoc: Document,
): void {
  const sibling = ownerDoc.createElementNS(w, "w:p")

  // pPr with pStyle=_HiddenChapterCounter + paragraph-mark vanish.
  // Paragraph-mark vanish (pPr.rPr.vanish) is what collapses the
  // paragraph to zero height; without it, the paragraph would still
  // occupy a line even when its runs are all hidden.
  const pPr = ownerDoc.createElementNS(w, "w:pPr")
  const pStyle = ownerDoc.createElementNS(w, "w:pStyle")
  pStyle.setAttributeNS(w, "w:val", HIDDEN_CHAPTER_COUNTER_STYLE_ID)
  pPr.appendChild(pStyle)
  const pRPr = ownerDoc.createElementNS(w, "w:rPr")
  pRPr.appendChild(ownerDoc.createElementNS(w, "w:vanish"))
  pPr.appendChild(pRPr)
  sibling.appendChild(pPr)

  // SEQ field runs, each with vanish (belt-and-braces with the style's
  // own vanish — covers Word versions that ignore style-level vanish
  // when computing visibility).
  const { runs } = emitSeqField(ownerDoc, {
    identifier: chapterCounterIdentifier(styleId),
    format,
  })
  for (const r of runs) {
    addVanishRPr(r, ownerDoc)
    sibling.appendChild(r)
  }

  // Insert immediately after the heading paragraph in its container.
  const parent = heading.parentNode!
  parent.insertBefore(sibling, heading.nextSibling)
}

/** Inject the `_HiddenChapterCounter` style definition into `stylesDoc`
 *  when not already present. Called from the apply pipeline whenever
 *  `injectChapterCounters` emits at least one sibling — without the
 *  matching style entry the `<w:pStyle>` reference is dangling.
 *
 *  Style shape:
 *   - `<w:hidden/>` + `<w:semiHidden/>` + `<w:unhideWhenUsed/>` so it
 *     doesn't clutter the agent / user style panels
 *   - `<w:pPr>` with zero spacing and paragraph-mark vanish (defense
 *     in depth — paragraphs also set the same vanish on their pPr.rPr)
 *   - `<w:rPr>` with `<w:vanish/>` and 2-half-point size so the runs
 *     are character-hidden and minimal even on Word versions that
 *     ignore the pPr vanish
 *
 *  Returns true when the style was newly injected. */
export function ensureHiddenChapterCounterStyle(stylesDoc: Document): boolean {
  const root = stylesDoc.documentElement
  if (!root) return false
  for (const s of getChildrenNS(root, w, "style")) {
    if (wAttr(s, "styleId") === HIDDEN_CHAPTER_COUNTER_STYLE_ID) return false
  }
  const style = stylesDoc.createElementNS(w, "w:style")
  style.setAttributeNS(w, "w:type", "paragraph")
  style.setAttributeNS(w, "w:styleId", HIDDEN_CHAPTER_COUNTER_STYLE_ID)

  const name = stylesDoc.createElementNS(w, "w:name")
  name.setAttributeNS(w, "w:val", "_Hidden Chapter Counter")
  style.appendChild(name)

  const basedOn = stylesDoc.createElementNS(w, "w:basedOn")
  basedOn.setAttributeNS(w, "w:val", "Normal")
  style.appendChild(basedOn)

  style.appendChild(stylesDoc.createElementNS(w, "w:hidden"))
  style.appendChild(stylesDoc.createElementNS(w, "w:semiHidden"))
  style.appendChild(stylesDoc.createElementNS(w, "w:unhideWhenUsed"))

  // Style-level pPr is CT_PPrGeneral — rPr is NOT a valid child here
  // (paragraph-mark vanish goes on each paragraph's own pPr.rPr in the
  // emit path, not in the style def). Style's top-level rPr below
  // carries the run-level vanish that covers paragraph mark too via
  // Word's "paragraph mark inherits style run defaults" rule.
  const pPr = stylesDoc.createElementNS(w, "w:pPr")
  const spacing = stylesDoc.createElementNS(w, "w:spacing")
  spacing.setAttributeNS(w, "w:before", "0")
  spacing.setAttributeNS(w, "w:after", "0")
  spacing.setAttributeNS(w, "w:line", "240")
  spacing.setAttributeNS(w, "w:lineRule", "auto")
  pPr.appendChild(spacing)
  style.appendChild(pPr)

  const rPr = stylesDoc.createElementNS(w, "w:rPr")
  rPr.appendChild(stylesDoc.createElementNS(w, "w:vanish"))
  const sz = stylesDoc.createElementNS(w, "w:sz")
  sz.setAttributeNS(w, "w:val", "2")
  rPr.appendChild(sz)
  style.appendChild(rPr)

  root.appendChild(style)
  return true
}
