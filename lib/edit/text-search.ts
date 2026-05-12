/**
 * Character-level text search across paragraphs. Atomic find primitive: locate
 * a literal or regex pattern in the document body, return per-match metadata
 * (paragraph index, run index, offset, length, context preview, and any
 * structural region the match falls inside). The agent decides what to do
 * with the coordinates — replace via `set-run`, deeper inspection via
 * `inspect_runs`, coverage validation, or just browsing.
 *
 * Design choices:
 *   - Paragraph is the search unit. Matches don't span paragraphs.
 *   - <w:tab/> and <w:br/> become sentinel chars (\t / \n) in the projection;
 *     matches that would cross those boundaries are rejected (text flow does
 *     not naturally span them).
 *   - <w:drawing> / <w:object> / <m:oMath> become ￼ (object replacement
 *     character) in the projection — invisible to literal/regex matches that
 *     don't deliberately include that codepoint.
 *   - Run-index in the output refers to direct `<w:r>` children of `<w:p>` —
 *     this matches `set-run`'s runIndex semantics. Matches in nested runs
 *     (inside <w:hyperlink> / <w:ins> / <w:del> / <w:sdt>) report no runIndex;
 *     the structural region annotation tells the agent where they are.
 *   - Field-region detection uses the same fldChar depth state-machine as
 *     `lib/blockers.ts`; a match between <w:fldChar begin> and the matching
 *     <w:fldChar end> is reported as `field`.
 */

import { walkIndexedParagraphs } from "@lib/edit/locator.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, wAttr } from "@lib/xml/xml-utils.ts"

const w = NS.w

const TAB_SENTINEL = "\t"
const BREAK_SENTINEL = "\n"
const OBJECT_PLACEHOLDER = "￼"

export type RegionKind = "tracked-change" | "field" | "sdt" | "hyperlink"

interface ProjectionSegment {
  start: number
  end: number
  /** 1-based direct-child run index in the paragraph. Null for nested runs
   * (inside hyperlink/ins/del/sdt) or sentinel positions (tab/break/object). */
  runIndex: number | null
  region: RegionKind | null
}

interface ParagraphProjection {
  text: string
  segments: ProjectionSegment[]
}

export interface MatchHit {
  paragraphIndex: number
  ch: number
  len: number
  matched: string
  runIndex: number | null
  runIndexEnd: number | null
  crossRun: boolean
  region: RegionKind | null
  context: string
}

export interface SearchOptions {
  pattern: string
  regex?: boolean
  paraIndex?: number
  paraRange?: { from: number; to: number }
  contextChars?: number
  bracket?: [string, string]
  limit?: number
}

export function searchDocument(documentDoc: Document, opts: SearchOptions): MatchHit[] {
  if (!opts.pattern || opts.pattern.length === 0) {
    throw new Error("search pattern is empty")
  }
  const ctxN = opts.contextChars ?? 25
  const [bL, bR] = opts.bracket ?? ["‹", "›"]
  const limit = opts.limit ?? Infinity
  const re = opts.regex ? new RegExp(opts.pattern, "g") : null

  const hits: MatchHit[] = []
  const fieldState = { depth: 0 }

  for (const p of walkIndexedParagraphs(documentDoc)) {
    if (opts.paraIndex !== undefined && p.index !== opts.paraIndex) continue
    if (opts.paraRange) {
      if (p.index < opts.paraRange.from || p.index > opts.paraRange.to) continue
    }
    const proj = buildParagraphProjection(p.element, fieldState)
    if (proj.text.length === 0) continue

    const matches = findInProjection(proj.text, opts.pattern, re)
    for (const m of matches) {
      const slice = proj.text.slice(m.start, m.start + m.length)
      if (slice.includes(TAB_SENTINEL) || slice.includes(BREAK_SENTINEL)) continue

      const overlapping = proj.segments.filter(
        (seg) => seg.start < m.start + m.length && seg.end > m.start,
      )
      if (overlapping.length === 0) continue

      let firstRunIdx: number | null = null
      let lastRunIdx: number | null = null
      let region: RegionKind | null = null
      for (const seg of overlapping) {
        if (seg.runIndex !== null) {
          if (firstRunIdx === null) firstRunIdx = seg.runIndex
          lastRunIdx = seg.runIndex
        }
        if (region === null && seg.region !== null) region = seg.region
      }
      const crossRun = firstRunIdx !== null && lastRunIdx !== null && firstRunIdx !== lastRunIdx

      const matched = slice
      const beforeRaw = proj.text.slice(Math.max(0, m.start - ctxN), m.start)
      const afterRaw = proj.text.slice(m.start + m.length, m.start + m.length + ctxN)
      const before = beforeRaw.replace(/[\t\n￼]/g, " ")
      const after = afterRaw.replace(/[\t\n￼]/g, " ")
      const elidedL = m.start > ctxN ? "..." : ""
      const elidedR = m.start + m.length + ctxN < proj.text.length ? "..." : ""
      const context = `${elidedL}${before}${bL}${matched}${bR}${after}${elidedR}`

      hits.push({
        paragraphIndex: p.index,
        ch: m.start,
        len: m.length,
        matched,
        runIndex: firstRunIdx,
        runIndexEnd: lastRunIdx,
        crossRun,
        region,
        context,
      })
      if (hits.length >= limit) return hits
    }
  }
  return hits
}

/* ------------- projection ------------- */

function buildParagraphProjection(
  pEl: Element,
  fieldState: { depth: number },
): ParagraphProjection {
  const segments: ProjectionSegment[] = []
  let text = ""
  // 1-based: the next direct <w:r> child gets this index, then increments.
  let directRunIdx = 1

  function appendNode(node: Element, scopeRegion: RegionKind | null): void {
    if (node.namespaceURI !== w) return
    const local = node.localName
    let effectiveRegion = scopeRegion
    if (local === "ins" || local === "del") effectiveRegion ??= "tracked-change"
    else if (local === "sdt") effectiveRegion ??= "sdt"
    else if (local === "hyperlink") effectiveRegion ??= "hyperlink"

    if (local === "r") {
      const isDirect = node.parentNode === pEl
      const idx = isDirect ? directRunIdx : null
      if (isDirect) directRunIdx++
      appendRun(node, idx, effectiveRegion)
      return
    }

    if (local === "hyperlink" || local === "ins" || local === "del" || local === "sdt") {
      const recurseRoot = local === "sdt" ? (firstChildNS(node, w, "sdtContent") ?? node) : node
      for (const child of getChildren(recurseRoot)) {
        appendNode(child, effectiveRegion)
      }
      return
    }

    if (local === "pPr") return
  }

  function appendRun(rEl: Element, runIndex: number | null, region: RegionKind | null): void {
    for (const c of getChildren(rEl)) {
      if (c.namespaceURI !== w) {
        addSegment(OBJECT_PLACEHOLDER, runIndex, region)
        continue
      }
      const ln = c.localName
      if (ln === "t") {
        const s = c.textContent ?? ""
        if (s) addSegment(s, runIndex, region)
      } else if (ln === "tab") {
        addSegment(TAB_SENTINEL, runIndex, region)
      } else if (ln === "br" || ln === "cr") {
        addSegment(BREAK_SENTINEL, runIndex, region)
      } else if (ln === "fldChar") {
        const ftype = wAttr(c, "fldCharType")
        if (ftype === "begin") fieldState.depth++
        else if (ftype === "end") {
          if (fieldState.depth > 0) fieldState.depth--
        }
      } else if (ln === "instrText") {
        // field code; not displayed text
      } else if (ln === "drawing" || ln === "object" || ln === "pict") {
        addSegment(OBJECT_PLACEHOLDER, runIndex, region)
      } else if (ln === "sym") {
        addSegment(OBJECT_PLACEHOLDER, runIndex, region)
      }
    }
  }

  function addSegment(s: string, runIndex: number | null, region: RegionKind | null): void {
    const start = text.length
    text += s
    const end = text.length
    const finalRegion: RegionKind | null = fieldState.depth > 0 ? "field" : region
    segments.push({ start, end, runIndex, region: finalRegion })
  }

  for (const node of getChildren(pEl)) {
    if (node.namespaceURI !== w) continue
    appendNode(node, null)
  }

  return { text, segments }
}

/* ------------- match ------------- */

interface RawMatch {
  start: number
  length: number
}

function findInProjection(text: string, pattern: string, re: RegExp | null): RawMatch[] {
  const out: RawMatch[] = []
  if (re) {
    for (const m of text.matchAll(re)) {
      if (m[0].length === 0) continue
      out.push({ start: m.index ?? 0, length: m[0].length })
    }
  } else {
    let from = 0
    while (true) {
      const idx = text.indexOf(pattern, from)
      if (idx < 0) break
      out.push({ start: idx, length: pattern.length })
      from = idx + pattern.length
    }
  }
  return out
}

/* ------------- description helpers ------------- */

export function describeRegion(region: RegionKind): string {
  switch (region) {
    case "tracked-change":
      return "in tracked change"
    case "field":
      return "in field result"
    case "sdt":
      return "in content control"
    case "hyperlink":
      return "in hyperlink"
  }
}
