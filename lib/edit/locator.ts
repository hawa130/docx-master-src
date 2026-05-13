/**
 * Locator → concrete <w:p> Element references.
 *
 * The resolver collapses every Locator kind into the same `ResolvedTarget`
 * shape (paragraphs[] + container). Op execution downstream is locator-
 * agnostic — the only place locator semantics live is here.
 *
 * Container uniformity rule: `paragraph` and `range` locators only reach
 * indexed paragraphs (body + layout-table cells, per DocumentParser). A
 * `range` whose endpoints sit in different containers (one in body, one in a
 * cell) is rejected — replace/insert/delete across structural boundaries
 * has no clean OOXML semantics. Data/form table cells are unindexed and
 * thus only reachable via a `cell` locator.
 */

import type { ParsedParagraph } from "@lib/parse/types.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, textContent, wVal } from "@lib/xml/xml-utils.ts"
import { summarizeTable } from "@lib/parse/table-classifier.ts"
import { assertNever, type Locator, type ResolvedTarget } from "@lib/config/edit-types.ts"

/* ------------- indexed walk ------------- */

interface IndexedPara {
  /** 1-based, matches DocumentParser. */
  index: number
  element: Element
  /** Body element or <w:tc>. */
  container: Element
}

/** Walk the document in DocumentParser order and return every indexed
 * paragraph with its element + container. Skips paragraphs inside data/form
 * tables (they're unindexed; reachable only via cell locator). */
export function walkIndexedParagraphs(documentDoc: Document): IndexedPara[] {
  const out: IndexedPara[] = []
  const root = documentDoc.documentElement
  if (!root) return out
  const body = firstChildNS(root, NS.w, "body")
  if (!body) return out
  let nextIndex = 1
  const recur = (parent: Element): void => {
    for (const child of getChildren(parent)) {
      if (child.namespaceURI !== NS.w) continue
      if (child.localName === "p") {
        out.push({ index: nextIndex++, element: child, container: parent })
      } else if (child.localName === "tbl") {
        const summary = summarizeTable(child)
        if (summary.classification === "layout") {
          for (const tr of getChildrenNS(child, NS.w, "tr")) {
            for (const tc of getChildrenNS(tr, NS.w, "tc")) {
              recur(tc)
            }
          }
        }
        // data/form: leave untouched — unindexed.
      }
    }
  }
  recur(body)
  return out
}

/* ------------- table walk (for cell locator) ------------- */

interface TableRef {
  /** 1-based top-level table position (document order). Includes layout, data,
   * form alike — cell locator addresses any table by position so the
   * agent can target data/form tables that paragraph indices skip. */
  tableIndex: number
  element: Element
}

export function walkTopLevelTables(documentDoc: Document): TableRef[] {
  const out: TableRef[] = []
  const root = documentDoc.documentElement
  if (!root) return out
  const body = firstChildNS(root, NS.w, "body")
  if (!body) return out
  let i = 1
  for (const child of getChildren(body)) {
    if (child.namespaceURI === NS.w && child.localName === "tbl") {
      out.push({ tableIndex: i++, element: child })
    }
  }
  return out
}

/* ------------- resolver context ------------- */

export interface ResolverContext {
  documentDoc: Document
  body: Element
  indexed: IndexedPara[]
  indexByElement: Map<Element, number>
  tables: TableRef[]
  /** parsedDoc.paragraphs aligned with `indexed[i].index`. Resolver consults
   * this for outline level (heading locator) and styleId. */
  parsed: ParsedParagraph[]
}

export function buildResolverContext(
  documentDoc: Document,
  parsed: ParsedParagraph[],
): ResolverContext {
  const root = documentDoc.documentElement
  if (!root) throw new Error("document has no root element")
  const body = firstChildNS(root, NS.w, "body")
  if (!body) throw new Error("document has no body")
  const indexed = walkIndexedParagraphs(documentDoc)
  const indexByElement = new Map<Element, number>()
  for (const p of indexed) indexByElement.set(p.element, p.index)
  return {
    documentDoc,
    body,
    indexed,
    indexByElement,
    tables: walkTopLevelTables(documentDoc),
    parsed,
  }
}

/* ------------- resolve ------------- */

export function resolveLocator(loc: Locator, ctx: ResolverContext): ResolvedTarget {
  switch (loc.type) {
    case "paragraph":
      return resolveParagraph(loc.index, ctx)
    case "range":
      return resolveRange(loc.from, loc.to, ctx)
    case "cell":
      return resolveCell(loc.table, loc.row, loc.col, ctx)
    case "heading":
      return resolveHeading(loc.text, loc.level, ctx)
    case "whole-body":
      return resolveWholeBody(ctx)
    default:
      return assertNever(loc)
  }
}

function resolveParagraph(index: number, ctx: ResolverContext): ResolvedTarget {
  const hit = ctx.indexed[index - 1]
  if (!hit || hit.index !== index) {
    const max = ctx.indexed.length
    throw new Error(
      `paragraph #${index} not found. Document has ${max} indexed paragraph(s) (range: #1${max ? `–#${max}` : ""}).` +
        ` Paragraphs inside data/form tables are unindexed and only reachable via a cell locator.`,
    )
  }
  return { paragraphs: [hit.element], container: hit.container }
}

function resolveRange(from: number, to: number, ctx: ResolverContext): ResolvedTarget {
  const fromHit = ctx.indexed[from - 1]
  const toHit = ctx.indexed[to - 1]
  if (!fromHit || fromHit.index !== from) {
    throw new Error(`range.from: paragraph #${from} not found (max #${ctx.indexed.length}).`)
  }
  if (!toHit || toHit.index !== to) {
    throw new Error(`range.to: paragraph #${to} not found (max #${ctx.indexed.length}).`)
  }
  if (fromHit.container !== toHit.container) {
    throw new Error(
      `range #${from}–#${to} crosses a structural boundary (one endpoint in body, the other inside a layout-table cell). ` +
        `Splitting cross-boundary ranges has no clean OOXML semantics; split into two separate edits.`,
    )
  }
  const paragraphs: Element[] = []
  for (let i = from - 1; i <= to - 1; i++) {
    const p = ctx.indexed[i]!
    if (p.container !== fromHit.container) {
      throw new Error(
        `range #${from}–#${to}: paragraph #${p.index} sits in a different container — range cannot span structural boundaries.`,
      )
    }
    paragraphs.push(p.element)
  }
  return { paragraphs, container: fromHit.container }
}

function resolveCell(
  table: number,
  row: number,
  col: number,
  ctx: ResolverContext,
): ResolvedTarget {
  // All three coords are 1-based agent-facing; convert to 0-based for array
  // access here. Error messages keep the 1-based form so they line up with
  // what the agent wrote in their config.
  if (table < 1 || table > ctx.tables.length) {
    throw new Error(
      `cell.table: index ${table} out of range. Document has ${ctx.tables.length} top-level table(s); valid 1..${ctx.tables.length}.`,
    )
  }
  const tbl = ctx.tables[table - 1]!.element
  const rows = getChildrenNS(tbl, NS.w, "tr")
  if (row < 1 || row > rows.length) {
    throw new Error(
      `cell.row: index ${row} out of range. Table ${table} has ${rows.length} row(s); valid 1..${rows.length}.`,
    )
  }
  const cells = getChildrenNS(rows[row - 1]!, NS.w, "tc")
  if (col < 1 || col > cells.length) {
    throw new Error(
      `cell.col: index ${col} out of range. Table ${table} row ${row} has ${cells.length} cell(s); valid 1..${cells.length}.`,
    )
  }
  const tc = cells[col - 1]!
  const paragraphs = getChildrenNS(tc, NS.w, "p")
  return { paragraphs, container: tc }
}

function resolveHeading(
  text: string,
  level: number | undefined,
  ctx: ResolverContext,
): ResolvedTarget {
  const target = text.trim()
  // Compare against rendered text trimmed; outlineLevel from parsed pPr (which
  // already factors in the styleId cascade). Returns the FIRST hit — agents
  // wanting a specific occurrence should switch to a paragraph locator after
  // running find_paragraphs / overview to disambiguate.
  for (const p of ctx.parsed) {
    if (p.text.trim() !== target) continue
    if (level !== undefined && p.pPr.outlineLevel !== level) continue
    if (level === undefined && p.pPr.outlineLevel === undefined) continue // require it to be a heading at all
    const hit = ctx.indexed[p.index - 1]
    if (!hit) continue
    return { paragraphs: [hit.element], container: hit.container }
  }
  const lvlPart = level !== undefined ? ` at outline level ${level}` : ""
  throw new Error(
    `heading "${text}"${lvlPart} not found. Use find_paragraphs to locate by regex, ` +
      `then refer by paragraph index instead.`,
  )
}

function resolveWholeBody(ctx: ResolverContext): ResolvedTarget {
  // Iterate ctx.indexed so the resolved set matches DocumentParser's scope —
  // body + layout-table cells. A naive `getChildren(ctx.body)` would only
  // see body-level paragraphs and silently drop everything inside layout
  // tables (the half of the doc that form-style templates put text into).
  // Data/form table paragraphs stay out of scope here by design — they
  // require an explicit `cell` locator.
  const paragraphs = ctx.indexed.map((p) => p.element)
  return { paragraphs, container: ctx.body }
}

/** Resolve a `set-run` op's RunLocator. Returns the target paragraph element
 * and the specific run element to mutate. The run is selected by:
 *   - `runIndex: M`     → 1-based run index in the paragraph
 *   - `blank: K`        → Kth (1-based) run that's a "blank" placeholder
 *                         (whitespace-only text + rPr containing `<w:u/>`)
 *   - neither given     → first blank run (= `blank: 1`)
 *
 * Throws with a clear message when the paragraph isn't found, the index is
 * out of range, or the requested blank doesn't exist (lists what blanks ARE
 * present so the agent can adjust). */
export function resolveRunLocator(
  loc: { paragraph: number; blank?: number; runIndex?: number },
  ctx: ResolverContext,
): { paragraph: Element; run: Element } {
  const hit = ctx.indexed[loc.paragraph - 1]
  if (!hit || hit.index !== loc.paragraph) {
    const max = ctx.indexed.length
    throw new Error(
      `paragraph #${loc.paragraph} not found. Document has ${max} indexed paragraph(s).`,
    )
  }
  const runs = getChildrenNS(hit.element, NS.w, "r")
  if (runs.length === 0) {
    throw new Error(`paragraph #${loc.paragraph} has no runs to target.`)
  }
  if (loc.runIndex !== undefined) {
    if (loc.runIndex < 1 || loc.runIndex > runs.length) {
      throw new Error(
        `paragraph #${loc.paragraph}: runIndex ${loc.runIndex} out of range (paragraph has ${runs.length} run(s); valid 1..${runs.length}).`,
      )
    }
    return { paragraph: hit.element, run: runs[loc.runIndex - 1]! }
  }
  // blank-run mode (default when neither field given)
  const blankK = loc.blank ?? 1
  const blanks = runs.filter(isBlankRun)
  if (blanks.length === 0) {
    throw new Error(
      `paragraph #${loc.paragraph}: no blank runs found. ` +
        `A blank run is one whose text is whitespace-only and rPr carries <w:u/> ` +
        `(typical form-fill placeholder). Use \`runIndex\` to target a specific run by 1-based index instead.`,
    )
  }
  if (blankK < 1 || blankK > blanks.length) {
    throw new Error(
      `paragraph #${loc.paragraph}: blank ${blankK} out of range (paragraph has ${blanks.length} blank run(s); valid 1..${blanks.length}).`,
    )
  }
  return { paragraph: hit.element, run: blanks[blankK - 1]! }
}

/** Heuristic: a "blank" run is one whose text content is whitespace-only and
 * whose rPr declares underline (`<w:u/>`). Captures form-fill placeholder
 * runs without false-positive on legitimate empty runs that happen to lack
 * underline (those are usually inter-text spacers, not fillable slots). */
function isBlankRun(r: Element): boolean {
  const rPr = firstChildNS(r, NS.w, "rPr")
  if (!rPr) return false
  if (!firstChildNS(rPr, NS.w, "u")) return false
  let text = ""
  for (const c of getChildren(r)) {
    if (c.namespaceURI === NS.w && c.localName === "t") text += textContent(c)
  }
  return text.trim() === ""
}

/* ------------- helpers usable downstream ------------- */

/** Read the styleId of a paragraph element (for blocker logging / heading
 * disambiguation). Returns "Normal" when no explicit pStyle is set. */
export function paragraphStyleId(pEl: Element): string {
  const pPr = firstChildNS(pEl, NS.w, "pPr")
  const pStyle = pPr ? firstChildNS(pPr, NS.w, "pStyle") : null
  return (pStyle && wVal(pStyle)) || "Normal"
}

/** Plain-text content of a paragraph (concatenates <w:t> textContent across
 * all runs). Used by blockers / debug logging. */
export function paragraphText(pEl: Element): string {
  let out = ""
  for (const r of getChildrenNS(pEl, NS.w, "r")) {
    for (const c of getChildren(r)) {
      if (c.namespaceURI === NS.w && c.localName === "t") out += textContent(c)
    }
  }
  return out
}

/** Read the trailing sectPr at body level (if any). Used for "append to
 * body" semantics so we don't insert after the section descriptor. */
export function trailingBodySectPr(body: Element): Element | null {
  // body may end with a <w:sectPr> sibling; if absent, the last paragraph's
  // pPr.sectPr serves the role and stays where it is (we don't move it).
  const children = getChildren(body)
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i]!
    if (c.namespaceURI !== NS.w) continue
    if (c.localName === "sectPr") return c
    if (c.localName === "p" || c.localName === "tbl") return null
  }
  return null
}
