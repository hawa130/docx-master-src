/**
 * Pangu-spacing static lint for `edits[]` text content.
 *
 * "Pangu spacing" = an ASCII space typed between a CJK character and a
 * Latin / digit character (or vice versa) in Chinese prose. Word's own
 * `autoSpace` setting already inserts the visual gap between adjacent
 * CJK and Latin glyphs at render time; a typed ASCII space stacks on
 * top of that automatic gap and renders too wide. LLM-generated Chinese
 * prose hits this pattern routinely because the training corpus has
 * many CJK ↔ Latin pairs separated by a literal space.
 *
 * The lint walks every string-bearing field of every `edits[]` op
 * (Block.text on ParagraphBlock / CaptionBlock; CellContent strings on
 * cells; inline run text), runs two narrow regexes (CJK→Latin/digit and
 * the mirror), and returns a non-empty warning entry for each hit.
 * Wired into the dry-run and real-apply surfaces so the agent can scrub
 * the source. Not fatal — agents may need a literal space deliberately
 * (e.g. inside a quoted code identifier), so the engine just flags and
 * continues.
 *
 * Scope: only ASCII space `U+0020`. Full-width spaces (`U+3000`),
 * non-breaking spaces (`U+00A0`), and other whitespace are intentional
 * markers and never flagged. CJK detection uses the Unified Ideographs
 * block `U+4E00..U+9FFF` — extension planes don't appear in the prose
 * we target and would broaden false positives.
 */

import type { EditOp } from "@lib/config/edit-types.ts"

export interface PanguWarning {
  editIndex: number
  /** Up to ~30 characters of context around the hit. */
  snippet: string
  /** The matched 3-character window (CJK + space + Latin/digit, or
   *  the mirror). Useful for the agent to grep the source. */
  hit: string
}

// CJK Unified Ideographs block. Codepoint range avoids `\p{Script=Han}`
// which would pull in CJK Extension A/B/etc. — those aren't seen in
// normal CN/JP/KR prose and would broaden false-positive surface (rare
// radicals + private use overlap with technical glyphs that may
// legitimately want a literal space). U+4E00..U+9FFF is what Word's
// `autoSpace` itself targets.
const CJK = "[\\u4e00-\\u9fff]"
const ASCII_LATIN_OR_DIGIT = "[A-Za-z0-9]"
// Single ASCII space only — full-width and non-breaking spaces are
// intentional, never flagged. Matches one or more spaces so authors
// who typed two spaces still get caught.
const ASCII_SPACE = " +"
const PANGU_AFTER_CJK = new RegExp(`(${CJK})(${ASCII_SPACE})(${ASCII_LATIN_OR_DIGIT})`, "g")
const PANGU_BEFORE_CJK = new RegExp(`(${ASCII_LATIN_OR_DIGIT})(${ASCII_SPACE})(${CJK})`, "g")

/** Run the lint over a list of edit ops. Returns warnings in op order
 * with up to 5 entries per op; the report-side caps the displayed
 * count separately. Empty list when no hits. */
export function lintPanguInEdits(edits: readonly EditOp[]): PanguWarning[] {
  const warnings: PanguWarning[] = []
  for (const [i, op] of edits.entries()) {
    const texts = collectStringsFromOp(op)
    for (const text of texts) scanString(text, i, warnings)
  }
  return warnings
}

function scanString(text: string, editIndex: number, out: PanguWarning[]): void {
  // Per-op cap of 5 keeps reports legible when an agent pastes a long
  // bilingual paragraph riddled with the pattern — we don't need every
  // hit, the agent re-scans after fixing the first batch.
  const PER_OP_CAP = 5
  let perOpHits = 0
  const recordHit = (match: RegExpExecArray): void => {
    if (perOpHits >= PER_OP_CAP) return
    perOpHits++
    const start = Math.max(0, match.index - 12)
    const end = Math.min(text.length, match.index + match[0].length + 12)
    out.push({
      editIndex,
      snippet: text.slice(start, end),
      hit: match[0],
    })
  }
  PANGU_AFTER_CJK.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = PANGU_AFTER_CJK.exec(text)) !== null) recordHit(m)
  PANGU_BEFORE_CJK.lastIndex = 0
  while ((m = PANGU_BEFORE_CJK.exec(text)) !== null) recordHit(m)
}

/** Pull every author-supplied string out of one op's fragment. Walks
 * ParagraphBlock / CaptionBlock text (string shorthand + InlineRun
 * arrays), ImageBlock alt text, TableBlock cell content (string or
 * Block[]), and inline run text inside cells. Skips InlineRef +
 * InlineEquation — refs carry no prose, equations are LaTeX and have
 * their own conventions. */
function collectStringsFromOp(op: EditOp): string[] {
  const out: string[] = []
  const fragment =
    op.op === "replace"
      ? op.with
      : op.op === "insert-before" || op.op === "insert-after"
        ? op.content
        : null
  if (!fragment) return out
  for (const block of fragment) collectStringsFromBlock(block, out)
  return out
}

function collectStringsFromBlock(block: unknown, out: string[]): void {
  if (block === null || typeof block !== "object") return
  const b = block as Record<string, unknown>
  switch (b.type) {
    case "paragraph":
    case "caption":
      pushRichText(b.text, out)
      return
    case "image":
      if (typeof b.alt === "string") out.push(b.alt)
      return
    case "table":
      pushTableRows(b.rows, out)
      return
    // page-break / horizontal-rule / equation — no prose text fields.
    default:
      return
  }
}

function pushRichText(text: unknown, out: string[]): void {
  if (typeof text === "string") {
    out.push(text)
    return
  }
  if (!Array.isArray(text)) return
  for (const node of text) {
    if (node && typeof node === "object" && "text" in node) {
      const t = (node as { text: unknown }).text
      if (typeof t === "string") out.push(t)
    }
  }
}

function pushTableRows(rows: unknown, out: string[]): void {
  if (!Array.isArray(rows)) return
  for (const row of rows) {
    if (!row || typeof row !== "object") continue
    const cells = (row as { cells?: unknown }).cells
    if (!Array.isArray(cells)) continue
    for (const cell of cells) pushCellContent(cell, out)
  }
}

function pushCellContent(cell: unknown, out: string[]): void {
  if (typeof cell === "string") {
    out.push(cell)
    return
  }
  if (!cell || typeof cell !== "object") return
  // Cell object: `{ content: RichText | Block[], ... }` — recurse.
  const content = (cell as { content?: unknown }).content
  if (typeof content === "string") {
    out.push(content)
    return
  }
  if (!Array.isArray(content)) return
  // Heuristic: if the array members look like Blocks (have `type`),
  // treat as Block[]; otherwise treat as RichText InlineNodes.
  const first = content[0]
  if (first && typeof first === "object" && "type" in first) {
    for (const block of content) collectStringsFromBlock(block, out)
  } else {
    pushRichText(content, out)
  }
}
