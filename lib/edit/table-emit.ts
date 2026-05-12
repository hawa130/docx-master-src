/**
 * TableBlock → OOXML emission.
 *
 * Translates the JSON TableBlock into `<w:tbl>` + `<w:tblPr>` + `<w:tblGrid>`
 * + `<w:tr>` + `<w:tc>` with all merge / borders / cell-property emit. The
 * agent declares only "restart" cells (those with content); engine computes
 * a grid occupancy map and injects vMerge continuation cells where rowspans
 * claim cells in subsequent rows.
 *
 * Three OOXML correctness invariants enforced here:
 *
 *   1. CT_TblPr / CT_TrPr / CT_TcPr child order — Word loaders silently
 *      mis-render or refuse mis-ordered properties. Child-order helpers
 *      from `xml-order.ts` enforce it.
 *
 *   2. `<w:tblGrid>` matches effective column count — Word relies on this
 *      to lay out columns. Engine auto-generates from row structure if
 *      agent omits `cols`; if `cols` is provided, length must match.
 *
 *   3. Every `<w:tc>` contains at least one block-level child — Word
 *      rejects an empty cell with "needs repair". Empty content (empty
 *      string, empty Block[], empty InlineNode[]) is replaced with a
 *      single empty `<w:p/>`.
 *
 * Recursive cell content (TableBlock inside cells) is NOT supported in v1
 * by schema — CellBlockSchema excludes TableBlock. The grid occupancy
 * algorithm and continuation logic still applies to one-level tables.
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren } from "@lib/xml/xml-utils.ts"
import {
  TBL_PR_CHILD_ORDER,
  TC_PR_CHILD_ORDER,
  TR_PR_CHILD_ORDER,
  insertChildInOrder,
} from "@lib/xml/xml-order.ts"
import type { Block } from "@lib/config/edit-types.ts"
import { emitBlock, emitRichText, type EmitContext } from "@lib/edit/fragment-emit.ts"

const w = NS.w

/* ------------- types: normalized internal model -------------
 *
 * These mirror the schema definitions in edit-config-schema.ts but stay
 * independent (no z.infer) — the schema's recursive ZodMiniType loses
 * field-level inference. Schema validation guarantees the shapes upstream;
 * here we work with hand-typed structures.
 */

type BorderStyleString = "none" | "single" | "thick" | "double" | "dotted" | "dashed"
interface BorderEdgeObject {
  style: "single" | "thick" | "double" | "dotted" | "dashed"
  size?: number
  color?: string
}
type BorderEdge = BorderStyleString | BorderEdgeObject

interface BordersCustom {
  top?: BorderEdge
  bottom?: BorderEdge
  left?: BorderEdge
  right?: BorderEdge
  insideH?: BorderEdge
  insideV?: BorderEdge
}
type BordersPreset = "all" | "none" | "outer" | "three-line"
type Borders = BordersPreset | BordersCustom

type ColWidth = "auto" | number

interface CellObj {
  content: string | ReadonlyArray<unknown> // schema-validated; emit by shape inspection
  colspan?: number
  rowspan?: number
  vAlign?: "top" | "center" | "bottom"
  borders?: BordersCustom
  shading?: string
}

interface TableBlock {
  type: "table"
  rows: ReadonlyArray<ReadonlyArray<TableCell>>
  headerRows?: number
  headerStyle?: string
  cols?: ReadonlyArray<{ width: ColWidth }>
  borders?: Borders
  alignment?: "left" | "center" | "right"
  vAlign?: "top" | "center" | "bottom"
  layout?: "fixed" | "autofit"
}

type TableCell = string | ReadonlyArray<unknown> | CellObj

interface GridSlot {
  kind: "restart" | "vmerge-continue"
  cell: CellObj // normalized to object form
  colspan: number
  rowspan: number
  isHeaderRow: boolean
  /** Row index (0-based) where this slot is emitted. Used by the
   * three-line preset to inject the header-bottom line on the LAST
   * header row only (not all of them). */
  rowIndex: number
}

/* ------------- public entry ------------- */

export function emitTableBlock(block: TableBlock, ownerDoc: Document, ctx: EmitContext): Element {
  const headerRows = block.headerRows ?? 0
  const { grid, effectiveCols } = buildGrid(block.rows, headerRows)

  if (block.cols !== undefined && block.cols.length !== effectiveCols) {
    throw new Error(
      `table.cols length (${block.cols.length}) does not match effective column count (${effectiveCols}). ` +
        `Effective columns = max(declared cells per row + ongoing rowspan claims), expanded by colspans.`,
    )
  }

  const tbl = ownerDoc.createElementNS(w, "w:tbl")
  const tblPr = buildTblPr(block, ownerDoc)
  tbl.appendChild(tblPr)
  tbl.appendChild(buildTblGrid(block.cols, effectiveCols, ownerDoc))

  for (let r = 0; r < grid.length; r++) {
    const rowEl = ownerDoc.createElementNS(w, "w:tr")
    if (r < headerRows) {
      const trPr = ownerDoc.createElementNS(w, "w:trPr")
      const tblHeader = ownerDoc.createElementNS(w, "w:tblHeader")
      insertChildInOrder(trPr, tblHeader, TR_PR_CHILD_ORDER)
      rowEl.appendChild(trPr)
    }
    for (const slot of grid[r]!) {
      rowEl.appendChild(buildTc(slot, block, ownerDoc, ctx))
    }
    tbl.appendChild(rowEl)
  }
  return tbl
}

/* ------------- grid occupancy ------------- */

/** Walk declared cells per row; weave in continuation slots for ongoing
 * rowspans. Validates span bounds. Returns a fully resolved grid. */
function buildGrid(
  rows: ReadonlyArray<ReadonlyArray<TableCell>>,
  headerRows: number,
): { grid: GridSlot[][]; effectiveCols: number } {
  const grid: GridSlot[][] = rows.map(() => [])
  // Active rowspans: column index → { rowsRemaining, restartSlot }.
  const ongoing = new Map<number, { rowsRemaining: number; source: GridSlot }>()
  let effectiveCols: number | undefined

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!
    let col = 0
    let declaredIdx = 0

    while (declaredIdx < row.length || ongoing.has(col)) {
      // 1. Place any continuation cells for ongoing rowspans at this col.
      if (ongoing.has(col)) {
        const active = ongoing.get(col)!
        const cont: GridSlot = {
          kind: "vmerge-continue",
          cell: active.source.cell,
          colspan: active.source.colspan,
          rowspan: active.source.rowspan,
          isHeaderRow: r < headerRows,
          rowIndex: r,
        }
        grid[r]!.push(cont)
        active.rowsRemaining -= 1
        if (active.rowsRemaining <= 0) ongoing.delete(col)
        col += active.source.colspan
        continue
      }
      // 2. Place a declared cell.
      if (declaredIdx >= row.length) break
      const cellObj = normalizeCell(row[declaredIdx]!)
      const colspan = cellObj.colspan ?? 1
      const rowspan = cellObj.rowspan ?? 1
      if (rowspan > rows.length - r) {
        throw new Error(
          `table row ${r + 1} col ${col + 1}: rowspan ${rowspan} exceeds rows remaining (${rows.length - r}).`,
        )
      }
      const slot: GridSlot = {
        kind: "restart",
        cell: cellObj,
        colspan,
        rowspan,
        isHeaderRow: r < headerRows,
        rowIndex: r,
      }
      grid[r]!.push(slot)
      // Claim subsequent rows in each spanned column.
      if (rowspan > 1) {
        ongoing.set(col, { rowsRemaining: rowspan - 1, source: slot })
      }
      declaredIdx += 1
      col += colspan
    }
    // Drain any remaining ongoing claims that sit AFTER all declared cells
    // in this row (e.g. rowspan in the last column).
    while (ongoing.has(col)) {
      const active = ongoing.get(col)!
      const cont: GridSlot = {
        kind: "vmerge-continue",
        cell: active.source.cell,
        colspan: active.source.colspan,
        rowspan: active.source.rowspan,
        isHeaderRow: r < headerRows,
        rowIndex: r,
      }
      grid[r]!.push(cont)
      active.rowsRemaining -= 1
      if (active.rowsRemaining <= 0) ongoing.delete(col)
      col += active.source.colspan
    }
    if (effectiveCols === undefined) effectiveCols = col
    else if (col !== effectiveCols) {
      throw new Error(
        `table row ${r + 1}: declared cells + ongoing rowspans span ${col} columns; row 1 spans ${effectiveCols}. All rows must total the same column count.`,
      )
    }
  }
  if (ongoing.size > 0) {
    throw new Error(
      `table: rowspan extends past the last row at column(s) [${[...ongoing.keys()].join(", ")}].`,
    )
  }
  return { grid, effectiveCols: effectiveCols ?? 0 }
}

function normalizeCell(cell: TableCell): CellObj {
  if (typeof cell === "string") return { content: cell }
  if (Array.isArray(cell)) return { content: cell }
  return cell as CellObj
}

/* ------------- tblPr ------------- */

function buildTblPr(block: TableBlock, ownerDoc: Document): Element {
  const tblPr = ownerDoc.createElementNS(w, "w:tblPr")

  if (block.alignment) {
    const jc = ownerDoc.createElementNS(w, "w:jc")
    jc.setAttributeNS(w, "w:val", block.alignment)
    insertChildInOrder(tblPr, jc, TBL_PR_CHILD_ORDER)
  }

  const borders = resolveTableBorders(block)
  if (borders) {
    insertChildInOrder(tblPr, buildTblBorders(borders, ownerDoc), TBL_PR_CHILD_ORDER)
  }

  // Layout: default "autofit" (no element). Only emit when fixed.
  if (block.layout === "fixed") {
    const tblLayout = ownerDoc.createElementNS(w, "w:tblLayout")
    tblLayout.setAttributeNS(w, "w:type", "fixed")
    insertChildInOrder(tblPr, tblLayout, TBL_PR_CHILD_ORDER)
  }
  return tblPr
}

/** Resolve preset → BordersCustom + apply default. Returns null when borders
 * should be omitted entirely (only when `borders: "none"` is requested AND
 * we want to suppress the element entirely — but it's safer to always emit
 * tblBorders to override theme defaults). */
function resolveTableBorders(block: TableBlock): BordersCustom | null {
  const b = block.borders ?? "all"
  if (typeof b === "string") {
    switch (b) {
      case "all":
        return {
          top: "single",
          bottom: "single",
          left: "single",
          right: "single",
          insideH: "single",
          insideV: "single",
        }
      case "none":
        return {
          top: "none",
          bottom: "none",
          left: "none",
          right: "none",
          insideH: "none",
          insideV: "none",
        }
      case "outer":
        return {
          top: "single",
          bottom: "single",
          left: "single",
          right: "single",
          insideH: "none",
          insideV: "none",
        }
      case "three-line":
        // Top + bottom thick; sides + insides none. The header-row bottom
        // line is injected per-cell in buildTc when isHeaderRow + last
        // header row.
        return {
          top: "thick",
          bottom: "thick",
          left: "none",
          right: "none",
          insideH: "none",
          insideV: "none",
        }
    }
  }
  // Custom: fields default to "none" when omitted.
  return {
    top: b.top ?? "none",
    bottom: b.bottom ?? "none",
    left: b.left ?? "none",
    right: b.right ?? "none",
    insideH: b.insideH ?? "none",
    insideV: b.insideV ?? "none",
  }
}

function buildTblBorders(b: BordersCustom, ownerDoc: Document): Element {
  const el = ownerDoc.createElementNS(w, "w:tblBorders")
  const order: Array<keyof BordersCustom> = ["top", "left", "bottom", "right", "insideH", "insideV"]
  for (const side of order) {
    const edge = b[side]
    if (!edge) continue
    el.appendChild(buildBorderElement("w:" + side, edge, ownerDoc))
  }
  return el
}

function buildTcBorders(b: BordersCustom, ownerDoc: Document): Element {
  const el = ownerDoc.createElementNS(w, "w:tcBorders")
  // Cell-level borders: top/left/bottom/right + insideH/V are valid here
  // for merged-cell-region inner edges (per spec). Order per CT_TcBorders
  // is start/top/end/bottom/insideH/insideV but Word also accepts the
  // sided naming (top/left/bottom/right) — keep consistent with tblBorders.
  const order: Array<keyof BordersCustom> = ["top", "left", "bottom", "right", "insideH", "insideV"]
  for (const side of order) {
    const edge = b[side]
    if (!edge) continue
    el.appendChild(buildBorderElement("w:" + side, edge, ownerDoc))
  }
  return el
}

function buildBorderElement(qname: string, edge: BorderEdge, ownerDoc: Document): Element {
  const el = ownerDoc.createElementNS(w, qname)
  if (edge === "none") {
    el.setAttributeNS(w, "w:val", "nil")
    return el
  }
  let style: string
  let size: number // in pt
  let color = "auto"
  if (typeof edge === "string") {
    style = edge
    size = edge === "thick" ? 1.5 : 0.5
    if (edge === "double") size = 0.75
  } else {
    style = edge.style
    // "thick" object form: default to 1.5pt unless agent overrides
    size = edge.size ?? (edge.style === "thick" ? 1.5 : 0.5)
    color = edge.color ?? "auto"
  }
  // Map "thick" style to OOXML "single" with larger size — Word's val="thick"
  // is a deprecated alias for "single" with sz>=12.
  if (style === "thick") style = "single"
  el.setAttributeNS(w, "w:val", style)
  el.setAttributeNS(w, "w:sz", String(Math.max(2, Math.round(size * 8))))
  el.setAttributeNS(w, "w:space", "0")
  el.setAttributeNS(w, "w:color", color)
  return el
}

/* ------------- tblGrid ------------- */

function buildTblGrid(
  cols: TableBlock["cols"],
  effectiveCols: number,
  ownerDoc: Document,
): Element {
  const grid = ownerDoc.createElementNS(w, "w:tblGrid")
  if (cols !== undefined) {
    for (const c of cols) grid.appendChild(buildGridCol(c.width, ownerDoc))
  } else {
    for (let i = 0; i < effectiveCols; i++) grid.appendChild(buildGridCol("auto", ownerDoc))
  }
  return grid
}

function buildGridCol(width: ColWidth, ownerDoc: Document): Element {
  const col = ownerDoc.createElementNS(w, "w:gridCol")
  // <w:gridCol w:w="..."/> expects DXA (twentieths of a pt). "auto"
  // emits w:w="0" — Word treats zero-width gridCols as autofit hints.
  if (width === "auto") {
    col.setAttributeNS(w, "w:w", "0")
  } else {
    col.setAttributeNS(w, "w:w", String(Math.round(width * 20)))
  }
  return col
}

/* ------------- tc (cell) ------------- */

function buildTc(slot: GridSlot, block: TableBlock, ownerDoc: Document, ctx: EmitContext): Element {
  const tc = ownerDoc.createElementNS(w, "w:tc")
  const tcPr = buildTcPr(slot, block, ownerDoc)
  tc.appendChild(tcPr)

  // Content: emit declared content for restart cells; continuation cells
  // emit an empty paragraph (vMerge continue cells don't carry content).
  if (slot.kind === "vmerge-continue") {
    tc.appendChild(ownerDoc.createElementNS(w, "w:p"))
    return tc
  }
  const blocks = emitCellContent(slot.cell, block, slot.isHeaderRow, ownerDoc, ctx)
  if (blocks.length === 0) {
    tc.appendChild(ownerDoc.createElementNS(w, "w:p"))
  } else {
    for (const el of blocks) tc.appendChild(el)
    // If cell ends with <w:tbl>, Word requires a trailing <w:p>.
    const last = blocks[blocks.length - 1]!
    if (last.namespaceURI === w && last.localName === "tbl") {
      tc.appendChild(ownerDoc.createElementNS(w, "w:p"))
    }
  }
  return tc
}

function buildTcPr(slot: GridSlot, block: TableBlock, ownerDoc: Document): Element {
  const tcPr = ownerDoc.createElementNS(w, "w:tcPr")
  if (slot.colspan > 1) {
    const gridSpan = ownerDoc.createElementNS(w, "w:gridSpan")
    gridSpan.setAttributeNS(w, "w:val", String(slot.colspan))
    insertChildInOrder(tcPr, gridSpan, TC_PR_CHILD_ORDER)
  }
  if (slot.kind === "vmerge-continue") {
    const vMerge = ownerDoc.createElementNS(w, "w:vMerge")
    insertChildInOrder(tcPr, vMerge, TC_PR_CHILD_ORDER)
  } else if (slot.rowspan > 1) {
    const vMerge = ownerDoc.createElementNS(w, "w:vMerge")
    vMerge.setAttributeNS(w, "w:val", "restart")
    insertChildInOrder(tcPr, vMerge, TC_PR_CHILD_ORDER)
  }

  // Cell-level borders override: agent-provided OR three-line preset's
  // header-bottom injection.
  const cellBorders = slot.cell.borders
  const presetHeaderBottom = computePresetHeaderBottom(slot, block)
  if (cellBorders !== undefined || presetHeaderBottom !== undefined) {
    const merged: BordersCustom = { ...(cellBorders ?? {}) }
    if (presetHeaderBottom !== undefined && merged.bottom === undefined) {
      merged.bottom = presetHeaderBottom
    }
    insertChildInOrder(tcPr, buildTcBorders(merged, ownerDoc), TC_PR_CHILD_ORDER)
  }

  if (slot.cell.shading) {
    const shd = ownerDoc.createElementNS(w, "w:shd")
    shd.setAttributeNS(w, "w:val", "clear")
    shd.setAttributeNS(w, "w:color", "auto")
    shd.setAttributeNS(w, "w:fill", slot.cell.shading)
    insertChildInOrder(tcPr, shd, TC_PR_CHILD_ORDER)
  }
  // vAlign is ALWAYS emitted — Word's default (no w:vAlign) renders top,
  // but the skill default is "center" to match academic / formal
  // typography. Resolution order: per-cell explicit > table-level > "center".
  const resolvedVAlign = slot.cell.vAlign ?? block.vAlign ?? "center"
  const vAlign = ownerDoc.createElementNS(w, "w:vAlign")
  vAlign.setAttributeNS(w, "w:val", resolvedVAlign)
  insertChildInOrder(tcPr, vAlign, TC_PR_CHILD_ORDER)
  return tcPr
}

/** For the `"three-line"` preset, inject a thin bottom border on each cell
 * of the LAST header row — that's the "middle" line of the three-line look.
 * Returns undefined when not applicable. */
function computePresetHeaderBottom(slot: GridSlot, block: TableBlock): BorderEdge | undefined {
  if (block.borders !== "three-line") return undefined
  const headerRows = block.headerRows ?? 0
  if (headerRows === 0) return undefined
  if (!slot.isHeaderRow) return undefined
  // Only the bottommost header row gets the middle line — multi-row
  // headers would otherwise stack multiple middle lines, breaking the
  // academic three-line convention.
  if (slot.rowIndex !== headerRows - 1) return undefined
  return { style: "single", size: 0.5, color: "auto" }
}

/** Emit cell content as an array of block-level elements (paragraphs /
 * images / breaks). Handles all four cell-content forms. Empty content
 * returns `[]`; caller injects empty `<w:p/>`. */
function emitCellContent(
  cell: CellObj,
  block: TableBlock,
  isHeaderRow: boolean,
  ownerDoc: Document,
  ctx: EmitContext,
): Element[] {
  const content = cell.content
  const headerStyle = isHeaderRow ? block.headerStyle : undefined

  // Plain string: single paragraph.
  if (typeof content === "string") {
    return [emitParagraphFromText(content, headerStyle, ownerDoc, ctx)]
  }
  if (!Array.isArray(content) || content.length === 0) {
    return []
  }
  // Discriminate by element shape. Block elements carry a `type` literal
  // matching the block-type discriminant. InlineNode has `text` or `refTo`
  // but no `type`. (Schema validation already ensured one or the other.)
  const first = content[0] as { type?: string }
  if (first && typeof first === "object" && typeof first.type === "string") {
    // Block[]
    const out: Element[] = []
    for (const b of content as ReadonlyArray<Block>) {
      const el = emitBlock(b, ownerDoc, ctx)
      // Apply headerStyle as default pStyle on paragraph blocks that
      // don't already declare one.
      if (headerStyle && el.namespaceURI === w && el.localName === "p" && !hasPStyle(el)) {
        applyDefaultPStyle(el, headerStyle, ownerDoc)
      }
      out.push(el)
    }
    return out
  }
  // InlineNode[] — wrap in a single paragraph using emitRichText
  return [emitParagraphFromRichText(content, headerStyle, ownerDoc, ctx)]
}

function emitParagraphFromText(
  text: string,
  headerStyle: string | undefined,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  const p = ownerDoc.createElementNS(w, "w:p")
  if (headerStyle) applyDefaultPStyle(p, headerStyle, ownerDoc)
  for (const r of emitRichText(text, ownerDoc, ctx, undefined)) p.appendChild(r)
  return p
}

function emitParagraphFromRichText(
  // InlineNode[] — schema-typed but erased here. emitRichText accepts.
  inline: unknown,
  headerStyle: string | undefined,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  const p = ownerDoc.createElementNS(w, "w:p")
  if (headerStyle) applyDefaultPStyle(p, headerStyle, ownerDoc)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  for (const r of emitRichText(inline as any, ownerDoc, ctx, undefined)) p.appendChild(r)
  return p
}

function hasPStyle(p: Element): boolean {
  const pPr = firstChildNS(p, w, "pPr")
  if (!pPr) return false
  return firstChildNS(pPr, w, "pStyle") !== null
}

function applyDefaultPStyle(p: Element, styleId: string, ownerDoc: Document): void {
  let pPr = firstChildNS(p, w, "pPr")
  if (!pPr) {
    pPr = ownerDoc.createElementNS(w, "w:pPr")
    p.insertBefore(pPr, p.firstChild)
  }
  if (firstChildNS(pPr, w, "pStyle")) return
  const pStyle = ownerDoc.createElementNS(w, "w:pStyle")
  pStyle.setAttributeNS(w, "w:val", styleId)
  pPr.insertBefore(pStyle, pPr.firstChild)
}

/* ------------- post-emit sequencing (container-agnostic) ------------- */

/**
 * Walk `container`'s direct children and fix two Word-mandatory invariants:
 *
 *   1. Adjacent `<w:tbl>` siblings get an empty `<w:p/>` injected between.
 *   2. If the container's last meaningful child is `<w:tbl>`, append `<w:p/>`.
 *      "Meaningful" excludes a trailing `<w:sectPr>` (body level); a sectPr
 *      after a `<w:tbl>` still violates the rule (Word wants `<w:p>`
 *      between `<w:tbl>` and `<w:sectPr>`), so it's treated as needing a
 *      separator.
 *
 * Both rules are idempotent — repeat calls are no-ops once normalized.
 * Call after inserting any fragment that may include TableBlock output,
 * for both body and cell containers.
 */
export function normalizeTableSequencing(container: Element, ownerDoc: Document): void {
  const isTbl = (el: Element) => el.namespaceURI === w && el.localName === "tbl"
  const isSectPr = (el: Element) => el.namespaceURI === w && el.localName === "sectPr"
  const newP = () => ownerDoc.createElementNS(w, "w:p")

  // Single pass: separators + final trailing.
  const children = getChildren(container)
  for (let i = 0; i < children.length - 1; i++) {
    const cur = children[i]!
    const next = children[i + 1]!
    if (isTbl(cur) && (isTbl(next) || isSectPr(next))) {
      container.insertBefore(newP(), next)
    }
  }
  // Trailing: last child is a table (no sectPr after) → append paragraph.
  const last = container.lastElementChild
  if (last && isTbl(last)) {
    container.appendChild(newP())
  }
}
