/**
 * `edit-caption` op handler.
 *
 * Replaces caption body text — runs after the primary anchor's
 * `<w:bookmarkEnd>`, preserving the fields (SEQ / STYLEREF) and bookmark
 * pair so cross-references continue to resolve.
 *
 * Target resolution:
 *   - `{ anchor: name }`: lookup via BookmarkAllocator. The bookmark must
 *     have been allocated by an earlier caption emit (range bookmark).
 *   - `{ captionId, index }`: walk body in order; count paragraphs that
 *     (a) have paragraphStyleId matching captions[<id>].styleId AND
 *     (b) contain a SEQ field with identifier matching `<id>`. Pick the
 *     1-based Nth.
 *
 * Throws when:
 *   - target paragraph has no `bookmarkEnd` (caption was emitted without
 *     an anchor — engine can't locate the body boundary)
 *   - resolution fails (unknown anchor / out-of-range index)
 */

import { NS } from "@lib/parse/types.ts"
import {
  buildPlainTextRun,
  firstChildNS,
  getChildren,
  wAttr,
  walkBodyParagraphs,
} from "@lib/xml/xml-utils.ts"
import { parseFieldRuns } from "@lib/edit/fields/field-parse.ts"
import type { BookmarkAllocator } from "@lib/edit/bookmark.ts"
import type { ResolvedCaptionConfig } from "@lib/edit/caption-counter.ts"

const w = NS.w

export type EditCaptionTarget = { anchor: string } | { captionId: string; index: number }

export interface EditCaptionInput {
  documentDoc: Document
  target: EditCaptionTarget
  text: string
  bookmarkAllocator: BookmarkAllocator
  captionsConfigs: Map<string, ResolvedCaptionConfig>
}

/** Resolve target to a caption paragraph element. Throws on miss. */
export function resolveEditCaptionTarget(input: EditCaptionInput): Element {
  if ("anchor" in input.target) {
    const rec = input.bookmarkAllocator.resolveByName(input.target.anchor)
    if (!rec) {
      throw new Error(
        `edit-caption: anchor "${input.target.anchor}" was not found. ` +
          `Declare it on the caption block before editing, or pick a different target.`,
      )
    }
    return rec.element
  }
  const { captionId, index } = input.target
  const config = input.captionsConfigs.get(captionId)
  if (!config) {
    throw new Error(`edit-caption: captionId "${captionId}" is not declared in captions table.`)
  }
  const matches: Element[] = []
  const body = firstChildNS(input.documentDoc.documentElement, w, "body")
  if (!body) {
    throw new Error("edit-caption: document has no body")
  }
  for (const para of walkBodyParagraphs(body)) {
    if (paragraphStyleId(para) !== config.paragraphStyleId) continue
    if (!paragraphContainsSeq(para, captionId)) continue
    matches.push(para)
  }
  if (index < 1 || index > matches.length) {
    throw new Error(
      `edit-caption: captionId "${captionId}" index ${index} out of range. ` +
        `Document has ${matches.length} caption paragraph(s) of this identifier.`,
    )
  }
  return matches[index - 1]!
}

/** Replace the caption paragraph's body text — runs after the primary
 * `<w:bookmarkEnd>`. Existing body runs are removed; a new bodySeparator
 * + text pair is appended. Preserves bookmark and field structure. */
export function applyEditCaption(
  paragraph: Element,
  text: string,
  config: ResolvedCaptionConfig,
  ownerDoc: Document,
): void {
  // Locate the last bookmarkEnd in the paragraph. Caption emits a
  // single primary bookmark wrapping prefix-to-suffix; everything after
  // is body content (bodySeparator + text). When no bookmark exists
  // (caption emitted without anchor), fall back to the last fldChar
  // end as the boundary.
  const children = getChildren(paragraph)
  let boundaryIdx = -1
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i]!
    if (c.namespaceURI !== w) continue
    if (c.localName === "bookmarkEnd") {
      boundaryIdx = i
      break
    }
  }
  if (boundaryIdx < 0) {
    // No bookmark; find the last run carrying fldChar end.
    for (let i = children.length - 1; i >= 0; i--) {
      const c = children[i]!
      if (c.namespaceURI !== w || c.localName !== "r") continue
      const fld = firstChildNS(c, w, "fldChar")
      if (fld && wAttr(fld, "fldCharType") === "end") {
        boundaryIdx = i
        break
      }
    }
  }
  if (boundaryIdx < 0) {
    throw new Error(
      "edit-caption: target paragraph has no field or bookmark boundary — can't locate the body region. " +
        "Was this paragraph emitted via the caption pipeline?",
    )
  }
  // Remove every child after the boundary.
  for (let i = children.length - 1; i > boundaryIdx; i--) {
    paragraph.removeChild(children[i]!)
  }
  // Append bodySeparator + text runs.
  if (text !== "") {
    paragraph.appendChild(buildPlainTextRun(ownerDoc, config.bodySeparator))
    paragraph.appendChild(buildPlainTextRun(ownerDoc, text))
  }
}

function paragraphStyleId(paragraph: Element): string | undefined {
  const pPr = firstChildNS(paragraph, w, "pPr")
  if (!pPr) return undefined
  const pStyle = firstChildNS(pPr, w, "pStyle")
  if (!pStyle) return undefined
  return wAttr(pStyle, "val") ?? undefined
}

function paragraphContainsSeq(paragraph: Element, identifier: string): boolean {
  const runs: Element[] = []
  for (const c of getChildren(paragraph)) {
    if (c.namespaceURI === w && c.localName === "r") runs.push(c)
  }
  const parsed = parseFieldRuns(runs)
  for (const entry of parsed) {
    if (entry.kind === "field" && entry.fieldType === "SEQ") {
      if (entry.details.identifier === identifier) return true
    }
  }
  return false
}
