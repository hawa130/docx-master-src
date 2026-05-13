/**
 * Standardize re-emit for caption paragraphs already in the document.
 *
 * Scans body for paragraphs where:
 *   (a) paragraph style ∈ captions.styleIds AND
 *   (b) paragraph contains a SEQ field with identifier matching one of
 *       the declared captions
 *
 * For each match (excluding paragraphs freshly emitted in this apply
 * pass), rebuilds the pre-body run sequence in place using the current
 * captions config. Preserves:
 *   - paragraph element identity (REF backfills' resolved targets stay
 *     valid)
 *   - bookmark id + name (so REFs keep resolving)
 *   - SEQ identifier (Word's running counter continues)
 *   - body text (everything after the primary bookmarkEnd / last fldChar
 *     end)
 *
 * Identifier mismatch (SEQ exists but its identifier isn't in the
 * captions config): pass through unchanged + warn.
 *
 * Produces fresh PendingCaptionFill entries pointing at the rebuilt
 * result text elements; caller appends to the apply pipeline's
 * pendingCaptionFills list so the counter sim renders these too.
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, wAttr, walkBodyParagraphs } from "@lib/xml/xml-utils.ts"
import { parseFieldRuns } from "@lib/edit/fields/field-parse.ts"
import { buildCaptionRunSequence } from "@lib/edit/caption-emit.ts"
import type { BookmarkAllocator } from "@lib/edit/bookmark.ts"
import type { PendingCaptionFill, ResolvedCaptionConfig } from "@lib/edit/caption-counter.ts"

const w = NS.w

export interface StandardizeCaptionsResult {
  fills: PendingCaptionFill[]
  warnings: string[]
}

export function standardizeCaptions(
  documentDoc: Document,
  captions: Map<string, ResolvedCaptionConfig>,
  bookmarkAllocator: BookmarkAllocator,
  freshlyEmitted: ReadonlySet<Element>,
): StandardizeCaptionsResult {
  const fills: PendingCaptionFill[] = []
  const warnings: string[] = []
  if (captions.size === 0) return { fills, warnings }

  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return { fills, warnings }

  // styleId → captionId lookup for paragraph filtering.
  const styleToIdentifier = new Map<string, string>()
  for (const [id, config] of captions) {
    styleToIdentifier.set(config.paragraphStyleId, id)
  }

  // Track paragraph index alongside the canonical walk — cheap counter that
  // mirrors the body-paragraph order an agent sees from `find_paragraphs` /
  // `overview`, so warnings can name the exact paragraph to fix.
  let paragraphIndex = 0
  for (const para of walkBodyParagraphs(body)) {
    paragraphIndex++
    if (freshlyEmitted.has(para)) continue
    const styleId = paragraphStyleId(para)
    if (!styleId) continue
    const expectedIdentifier = styleToIdentifier.get(styleId)
    if (!expectedIdentifier) continue

    const seqIdentifier = findSeqIdentifier(para)
    if (!seqIdentifier) continue // not a caption-shaped paragraph
    if (seqIdentifier !== expectedIdentifier) {
      warnings.push(
        `standardize-captions: paragraph #${paragraphIndex} styled as "${styleId}" carries SEQ "${seqIdentifier}" but the captions table expects "${expectedIdentifier}". Passed through unchanged — add the identifier to the captions table to bring it under standardize.`,
      )
      continue
    }

    const config = captions.get(expectedIdentifier)!
    const { bookmarkId, bookmarkName, bodyText, subGroup } = extractCaptionParts(
      para,
      expectedIdentifier,
      config.bodySeparator,
    )
    const fill = rebuildCaptionParagraphInPlace(
      para,
      config,
      bookmarkId,
      bookmarkName,
      bodyText,
      subGroup,
      documentDoc,
    )
    if (fill) fills.push(fill)
    // Maintain bookmark allocator binding for re-emitted captions so
    // REF backfills can resolve. The source-doc bookmark is already
    // indexed (allocator.nameIndex was populated at construction); we
    // just confirm it points at the right paragraph.
    if (bookmarkName) {
      bookmarkAllocator.bindRangeBookmark(bookmarkName, para)
    }
  }

  return { fills, warnings }
}

/** Replace runs after pPr in the paragraph with a freshly-built caption
 * sequence (using current config). Returns a PendingCaptionFill for the
 * new SEQ/STYLEREF result elements. */
function rebuildCaptionParagraphInPlace(
  paragraph: Element,
  config: ResolvedCaptionConfig,
  bookmarkId: number | undefined,
  bookmarkName: string | undefined,
  bodyText: string,
  subGroup: "start" | "continue" | undefined,
  ownerDoc: Document,
): PendingCaptionFill | undefined {
  const bookmark =
    bookmarkId !== undefined && bookmarkName !== undefined
      ? { id: bookmarkId, name: bookmarkName }
      : undefined
  const seq = buildCaptionRunSequence(ownerDoc, {
    captionConfig: config,
    subGroup,
    bookmark,
    body: bodyText === "" ? undefined : bodyText,
  })

  // Remove existing children except pPr, then append fresh runs.
  const existingChildren = getChildren(paragraph)
  const existingPPr = existingChildren.find((c) => c.namespaceURI === w && c.localName === "pPr")
  for (const c of existingChildren) {
    if (c !== existingPPr) paragraph.removeChild(c)
  }
  for (const r of seq.runs) paragraph.appendChild(r)

  return {
    paragraph,
    identifier: config.identifier,
    subGroup,
    chapterPrefixResults: seq.chapterPrefixResults,
    parentSeqResult: seq.parentSeqResult,
    subSeqResult: seq.subSeqResult,
  }
}

interface CaptionParts {
  bookmarkId: number | undefined
  bookmarkName: string | undefined
  bodyText: string
  subGroup: "start" | "continue" | undefined
}

function extractCaptionParts(
  paragraph: Element,
  identifier: string,
  bodySeparator: string,
): CaptionParts {
  let bookmarkId: number | undefined
  let bookmarkName: string | undefined
  const children = getChildren(paragraph)

  // Find the primary bookmarkStart (first in paragraph) — caption emits
  // one and only one in this position.
  for (const c of children) {
    if (c.namespaceURI === w && c.localName === "bookmarkStart") {
      const id = wAttr(c, "id")
      const name = wAttr(c, "name")
      if (id !== null && name !== null) {
        const n = parseInt(id, 10)
        if (Number.isFinite(n)) {
          bookmarkId = n
          bookmarkName = name
        }
      }
      break
    }
  }

  // Body text = everything in <w:t> elements after the boundary.
  let boundaryIdx = -1
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i]!
    if (c.namespaceURI === w && c.localName === "bookmarkEnd") {
      boundaryIdx = i
      break
    }
  }
  if (boundaryIdx < 0) {
    // No bookmark — fall back to last fldChar end.
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
  let bodyText = ""
  if (boundaryIdx >= 0) {
    // Reconstruct body text from runs after the boundary. Skip the
    // first run if it's the bodySeparator (single text run with leading
    // separator chars) — we'll re-emit it. Conservative: gather all <w:t>
    // text from runs after boundary, then trim leading separator
    // characters at the emit side. Simpler: just gather everything and
    // let the agent's bodySeparator config drive separator on re-emit.
    let collected = ""
    for (let i = boundaryIdx + 1; i < children.length; i++) {
      const c = children[i]!
      if (c.namespaceURI !== w || c.localName !== "r") continue
      for (const t of getChildren(c)) {
        if (t.namespaceURI === w && t.localName === "t") {
          collected += t.textContent ?? ""
        }
      }
    }
    // Strip the exact bodySeparator prefix so re-emit's bodySeparator
    // doesn't double up. Matches the literal configured separator
    // rather than a character class — body text legitimately starting
    // with "1:23 时刻..." stays intact.
    bodyText = collected.startsWith(bodySeparator)
      ? collected.slice(bodySeparator.length)
      : collected
  }

  // subGroup detection from the existing fields:
  //   - parent SEQ carries \c (repeat)         → "continue"
  //   - sub-counter SEQ "identifierSub" present
  //     without parent \c                      → "start"
  //   - otherwise                              → undefined (standalone)
  const subGroup = detectSubGroup(paragraph, identifier)

  return { bookmarkId, bookmarkName, bodyText, subGroup }
}

function detectSubGroup(paragraph: Element, identifier: string): "start" | "continue" | undefined {
  const runs: Element[] = []
  for (const c of getChildren(paragraph)) {
    if (c.namespaceURI === w && c.localName === "r") runs.push(c)
  }
  const parsed = parseFieldRuns(runs)
  let parentRepeat = false
  let subSeqPresent = false
  const subId = `${identifier}Sub`
  for (const entry of parsed) {
    if (entry.kind !== "field" || entry.fieldType !== "SEQ") continue
    if (entry.details.identifier === identifier && entry.details.repeat) parentRepeat = true
    if (entry.details.identifier === subId) subSeqPresent = true
  }
  if (parentRepeat) return "continue"
  if (subSeqPresent) return "start"
  return undefined
}

function paragraphStyleId(paragraph: Element): string | undefined {
  const pPr = firstChildNS(paragraph, w, "pPr")
  if (!pPr) return undefined
  const pStyle = firstChildNS(pPr, w, "pStyle")
  if (!pStyle) return undefined
  return wAttr(pStyle, "val") ?? undefined
}

function findSeqIdentifier(paragraph: Element): string | undefined {
  const runs: Element[] = []
  for (const c of getChildren(paragraph)) {
    if (c.namespaceURI === w && c.localName === "r") runs.push(c)
  }
  const parsed = parseFieldRuns(runs)
  for (const entry of parsed) {
    if (entry.kind !== "field" || entry.fieldType !== "SEQ") continue
    // Skip `\c` (repeat) SEQs — engine-injected chapter prefixes read the
    // current counter value without advancing and aren't this paragraph's
    // own identifier. Same fix shape as inspect-caption / overview.
    if (entry.details.repeat) continue
    return entry.details.identifier
  }
  return undefined
}
