/**
 * Page-setup mutation: `<w:sectPr>` children — `pgSz` / `pgMar` / `cols`.
 *
 * Sparse-by-design: only declared fields mutate. Margins are merged per-edge
 * (declaring `top` doesn't disturb `bottom` etc.). Sections inherit the
 * top-level defaults and selectively override via `sections.<selector>`,
 * where the selector is `"all"`, `"N"`, or `"N-M"` (1-based, inclusive).
 *
 * Walking order matches `document-parser.ts`: paragraph-embedded sectPrs in
 * document order (sections 1..N-1), then the body-trailing sectPr (section
 * N). The Nth element of `collectSectPrs(body)` corresponds to section N
 * (1-based), so `SectionInfo[]` indices map directly.
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren } from "@lib/xml/xml-utils.ts"
import { SECT_PR_CHILD_ORDER, insertChildInOrder } from "@lib/xml/xml-order.ts"
import type { ApplyConfig } from "@lib/config/config-types.ts"
import { type Length, toTwips } from "@lib/shared/units.ts"
import { expandSectionSelector } from "@lib/apply/section-selector.ts"

const w = NS.w

type PageSetupConfig = NonNullable<ApplyConfig["pageSetup"]>
type SectionFields = NonNullable<PageSetupConfig["sections"]>[string]
type Orientation = "portrait" | "landscape"
type PaperSize = NonNullable<SectionFields["paperSize"]>
type Margins = NonNullable<SectionFields["margins"]>
type Columns = NonNullable<SectionFields["columns"]>
type PgNumType = NonNullable<SectionFields["pgNumType"]>

/** Paper size constants in twips. Portrait orientation (width < height).
 *  Engine swaps when orientation is landscape. */
const PAPER_SIZE_TWIPS: Record<string, readonly [number, number]> = {
  A3: [16838, 23811],
  A4: [11906, 16838],
  A5: [8392, 11906],
  Letter: [12240, 15840],
  Legal: [12240, 20160],
  B5: [9978, 14173],
  // Chinese 16-kai: 185mm × 260mm (published practice; varies by publisher,
  // this is the most common modern setting and matches WPS / Word zh-CN UI).
  "16K": [10489, 14741],
}

const DEFAULT_COLUMN_SPACE_TWIPS = toTwips("0.5cm", "columns.space")

/** ±1 twip absorbs cm/in round-trip noise (Word stores margins in twips
 *  but UI sets them in cm/mm; the same logical length can round to either
 *  side of a fractional twip). A 1-twip difference is 0.05mm — far below
 *  Word's display granularity. Used by mutatePgSz / mutatePgMar to decide
 *  whether a declared value differs from the source meaningfully. */
const TWIP_EQ_TOLERANCE = 1

interface SectionReportEntry {
  index: number // 1-based
  before: {
    paperSize?: string
    orientation: Orientation
    margins: { top: number; bottom: number; left: number; right: number }
  }
  after: {
    paperSize?: string
    orientation: Orientation
    margins: { top: number; bottom: number; left: number; right: number }
  }
  changed: boolean
}

export interface PageSetupReport {
  sections: SectionReportEntry[]
  /** Sections actually touched (i.e. their sectPr was modified). */
  touchedCount: number
}

/** Walk body in document order, collect every sectPr element. Order matches
 *  the parser's `SectionInfo[]` index (1-based). */
function collectSectPrs(body: Element): Element[] {
  const out: Element[] = []
  for (const child of getChildren(body)) {
    if (child.namespaceURI !== w) continue
    if (child.localName === "p") {
      const pPr = firstChildNS(child, w, "pPr")
      if (pPr) {
        const sectPr = firstChildNS(pPr, w, "sectPr")
        if (sectPr) out.push(sectPr)
      }
    } else if (child.localName === "sectPr") {
      out.push(child)
    }
  }
  return out
}

/** Build the effective per-section config by layering per-section overrides
 *  on top of the top-level defaults. Margins merge field-wise; paperSize /
 *  orientation / columns replace wholesale. Returns one entry per section
 *  (1-based; entry at index 0 is unused). */
function buildEffectiveConfigs(config: PageSetupConfig, sectionCount: number): SectionFields[] {
  const defaults: SectionFields = {
    paperSize: config.paperSize,
    orientation: config.orientation,
    margins: config.margins,
    columns: config.columns,
    pgNumType: config.pgNumType,
  }
  const effective: SectionFields[] = []
  for (let i = 0; i < sectionCount; i++) {
    effective.push({
      ...defaults,
      margins: defaults.margins ? { ...defaults.margins } : undefined,
      pgNumType: defaults.pgNumType ? { ...defaults.pgNumType } : undefined,
    })
  }
  if (!config.sections) return effective
  for (const [key, override] of Object.entries(config.sections)) {
    const indices = expandSectionSelector(key, sectionCount, "pageSetup.sections")
    for (const idx1 of indices) {
      const slot = effective[idx1 - 1]!
      if (override.paperSize !== undefined) slot.paperSize = override.paperSize
      if (override.orientation !== undefined) slot.orientation = override.orientation
      if (override.columns !== undefined) slot.columns = override.columns
      if (override.margins !== undefined) {
        slot.margins = { ...(slot.margins ?? {}), ...override.margins }
      }
      if (override.pgNumType !== undefined) {
        slot.pgNumType = { ...(slot.pgNumType ?? {}), ...override.pgNumType }
      }
    }
  }
  return effective
}

/** Resolve a PaperSize input to its (width, height) in twips, portrait
 *  orientation. Custom `{width, height}` falls through verbatim. */
function paperSizeToTwips(paper: PaperSize): { w: number; h: number } {
  if (typeof paper === "string") {
    const dims = PAPER_SIZE_TWIPS[paper]
    if (!dims) throw new Error(`pageSetup.paperSize: unknown size "${paper}"`)
    return { w: dims[0], h: dims[1] }
  }
  return {
    w: toTwips(paper.width, "paperSize.width"),
    h: toTwips(paper.height, "paperSize.height"),
  }
}

/** Build / replace `<w:pgSz>` on a sectPr. When paperSize is declared the
 *  result is (constants or custom) optionally swapped for landscape; when
 *  only orientation is declared the existing w/h are read and swapped to
 *  match the requested direction. Returns true when anything changed. */
function mutatePgSz(
  sectPr: Element,
  paperSize: PaperSize | undefined,
  orientation: Orientation | undefined,
  doc: Document,
): boolean {
  if (paperSize === undefined && orientation === undefined) return false

  const existing = firstChildNS(sectPr, w, "pgSz")
  const curW = existing ? parseInt(existing.getAttributeNS(w, "w") || "0", 10) : 0
  const curH = existing ? parseInt(existing.getAttributeNS(w, "h") || "0", 10) : 0
  const curOrient: Orientation =
    existing?.getAttributeNS(w, "orient") === "landscape"
      ? "landscape"
      : curW > 0 && curH > 0 && curW > curH
        ? "landscape"
        : "portrait"

  let portraitW: number
  let portraitH: number
  if (paperSize !== undefined) {
    const dims = paperSizeToTwips(paperSize)
    portraitW = dims.w
    portraitH = dims.h
  } else {
    // Only orientation changed: derive portrait dimensions from current w/h.
    // When source had no pgSz the dims are 0 — bail rather than emit invalid
    // <w:pgSz w:w="0" w:h="0"/>; orientation alone needs an existing pgSz to
    // pivot from.
    if (curW === 0 || curH === 0) {
      throw new Error(
        `pageSetup.orientation: section has no existing <w:pgSz> to pivot from; declare paperSize alongside orientation.`,
      )
    }
    portraitW = Math.min(curW, curH)
    portraitH = Math.max(curW, curH)
  }

  const targetOrient: Orientation = orientation ?? curOrient
  const [outW, outH] =
    targetOrient === "landscape" ? [portraitH, portraitW] : [portraitW, portraitH]

  // No-op when DOM already matches target — keeps touchedCount honest and
  // avoids spurious before/after entries in the dry-run report. ±1 twip
  // absorbs cm round-trip noise (see TWIP_EQ_TOLERANCE).
  if (
    existing &&
    Math.abs(outW - curW) <= TWIP_EQ_TOLERANCE &&
    Math.abs(outH - curH) <= TWIP_EQ_TOLERANCE &&
    (targetOrient === "landscape") === (existing.getAttributeNS(w, "orient") === "landscape")
  ) {
    return false
  }

  let pgSz = existing
  if (!pgSz) {
    pgSz = doc.createElementNS(w, "w:pgSz")
    insertChildInOrder(sectPr, pgSz, SECT_PR_CHILD_ORDER)
  }
  pgSz.setAttributeNS(w, "w:w", String(outW))
  pgSz.setAttributeNS(w, "w:h", String(outH))
  if (targetOrient === "landscape") {
    pgSz.setAttributeNS(w, "w:orient", "landscape")
  } else {
    pgSz.removeAttributeNS(w, "orient")
  }
  return true
}

const MARGIN_ATTR_KEYS: Array<{ field: keyof Margins; attr: string }> = [
  { field: "top", attr: "top" },
  { field: "right", attr: "right" },
  { field: "bottom", attr: "bottom" },
  { field: "left", attr: "left" },
  { field: "header", attr: "header" },
  { field: "footer", attr: "footer" },
  { field: "gutter", attr: "gutter" },
]

function mutatePgMar(sectPr: Element, margins: Margins | undefined, doc: Document): boolean {
  if (!margins) return false
  let pgMar = firstChildNS(sectPr, w, "pgMar")
  const fresh = !pgMar
  if (!pgMar) {
    pgMar = doc.createElementNS(w, "w:pgMar")
    insertChildInOrder(sectPr, pgMar, SECT_PR_CHILD_ORDER)
  }
  let changed = fresh
  for (const { field, attr } of MARGIN_ATTR_KEYS) {
    const v = margins[field]
    if (v === undefined) continue
    const next = toTwips(v as Length, `margins.${field}`)
    const cur = parseInt(pgMar.getAttributeNS(w, attr) || "", 10)
    if (Number.isNaN(cur) || Math.abs(cur - next) > TWIP_EQ_TOLERANCE) {
      pgMar.setAttributeNS(w, `w:${attr}`, String(next))
      changed = true
    }
  }
  return changed
}

function mutateCols(sectPr: Element, columns: Columns | undefined, doc: Document): boolean {
  if (columns === undefined) return false

  // Normalize both shorthand and object forms into a single descriptor.
  let count: number
  let space: number | undefined
  let separator: boolean | undefined
  let widths: number[] | undefined
  let spaces: number[] | undefined

  if (typeof columns === "number") {
    count = columns
  } else if (columns.widths !== undefined) {
    widths = columns.widths.map((v, i) => toTwips(v, `columns.widths[${i}]`))
    count = widths.length
    if (columns.spaces !== undefined) {
      spaces = columns.spaces.map((v, i) => toTwips(v, `columns.spaces[${i}]`))
    }
    separator = columns.separator
  } else {
    count = columns.count!
    space = columns.space !== undefined ? toTwips(columns.space, "columns.space") : undefined
    separator = columns.separator
  }

  // Build the target <w:cols> in a detached element, then compare to the
  // existing one before replacing — keeps touchedCount honest when an apply
  // re-declares the same column layout already on the doc.
  const cols = doc.createElementNS(w, "w:cols")
  cols.setAttributeNS(w, "w:num", String(count))
  if (widths === undefined) {
    cols.setAttributeNS(w, "w:space", String(space ?? DEFAULT_COLUMN_SPACE_TWIPS))
    cols.setAttributeNS(w, "w:equalWidth", "true")
  } else {
    cols.setAttributeNS(w, "w:equalWidth", "false")
    // Per ECMA-376 17.6.4, when equalWidth=false each <w:col> declares its
    // own w:w and (except the last column) w:space.
    for (let i = 0; i < widths.length; i++) {
      const col = doc.createElementNS(w, "w:col")
      col.setAttributeNS(w, "w:w", String(widths[i]))
      if (i < widths.length - 1) {
        const sp = spaces ? spaces[i]! : DEFAULT_COLUMN_SPACE_TWIPS
        col.setAttributeNS(w, "w:space", String(sp))
      }
      cols.appendChild(col)
    }
  }
  if (separator !== undefined) {
    cols.setAttributeNS(w, "w:sep", separator ? "true" : "false")
  }
  const existing = firstChildNS(sectPr, w, "cols")
  if (existing && colsElementsEqual(existing, cols)) return false
  if (existing) sectPr.removeChild(existing)
  insertChildInOrder(sectPr, cols, SECT_PR_CHILD_ORDER)
  return true
}

/** Structural comparison for two `<w:cols>` elements (built or DOM). Returns
 *  true when the attribute set (num/space/equalWidth/sep) and `<w:col>`
 *  children match. Skips ECMA-376 attributes we don't author (none today),
 *  so a doc with foreign attrs would falsely diff — acceptable since no-op
 *  detection is an optimization, not a correctness invariant. */
function colsElementsEqual(a: Element, b: Element): boolean {
  const attrs = ["num", "space", "equalWidth", "sep"]
  for (const at of attrs) {
    if (a.getAttributeNS(w, at) !== b.getAttributeNS(w, at)) return false
  }
  const aCols = Array.from(getChildren(a)).filter(
    (c) => c.namespaceURI === w && c.localName === "col",
  )
  const bCols = Array.from(getChildren(b)).filter(
    (c) => c.namespaceURI === w && c.localName === "col",
  )
  if (aCols.length !== bCols.length) return false
  for (let i = 0; i < aCols.length; i++) {
    if (aCols[i]!.getAttributeNS(w, "w") !== bCols[i]!.getAttributeNS(w, "w")) return false
    if (aCols[i]!.getAttributeNS(w, "space") !== bCols[i]!.getAttributeNS(w, "space")) return false
  }
  return true
}

/** Build / replace `<w:pgNumType>` on a sectPr. Idempotent on identical
 *  attribute sets so re-runs don't bump touchedCount. */
function mutatePgNumType(
  sectPr: Element,
  pgNumType: PgNumType | undefined,
  doc: Document,
): boolean {
  if (pgNumType === undefined) return false
  const existing = firstChildNS(sectPr, w, "pgNumType")
  const curFmt = existing?.getAttributeNS(w, "fmt") ?? null
  const curStart = existing?.getAttributeNS(w, "start") ?? null
  const wantFmt = pgNumType.fmt ?? null
  const wantStart = pgNumType.start !== undefined ? String(pgNumType.start) : null
  if (existing && curFmt === wantFmt && curStart === wantStart) return false
  if (existing) sectPr.removeChild(existing)
  const el = doc.createElementNS(w, "w:pgNumType")
  if (wantFmt !== null) el.setAttributeNS(w, "w:fmt", wantFmt)
  if (wantStart !== null) el.setAttributeNS(w, "w:start", wantStart)
  insertChildInOrder(sectPr, el, SECT_PR_CHILD_ORDER)
  return true
}

/** Apply pageSetup to every relevant sectPr in the document. Returns a
 *  per-section before/after snapshot for the dry-run report. */
export function applyPageSetup(documentDoc: Document, config: PageSetupConfig): PageSetupReport {
  const body = firstChildNS(documentDoc.documentElement!, w, "body")
  if (!body) return { sections: [], touchedCount: 0 }

  const sectPrs = collectSectPrs(body)
  const sectionCount = sectPrs.length
  if (sectionCount === 0) return { sections: [], touchedCount: 0 }

  const effective = buildEffectiveConfigs(config, sectionCount)

  const report: SectionReportEntry[] = []
  let touched = 0
  for (let i = 0; i < sectionCount; i++) {
    const sectPr = sectPrs[i]!
    const fields = effective[i]!
    const before = snapshotSection(sectPr)

    const a = mutatePgSz(sectPr, fields.paperSize, fields.orientation, documentDoc)
    const b = mutatePgMar(sectPr, fields.margins, documentDoc)
    const c = mutateCols(sectPr, fields.columns, documentDoc)
    const d = mutatePgNumType(sectPr, fields.pgNumType, documentDoc)
    const changed = a || b || c || d

    const after = snapshotSection(sectPr)
    if (changed) touched++
    report.push({ index: i + 1, before, after, changed })
  }
  return { sections: report, touchedCount: touched }
}

/** Read a section's display-relevant attributes from the current DOM state.
 *  Called twice per section (before + after mutation); margin attrs default
 *  to 0 when no <w:pgMar> is present. */
function snapshotSection(sectPr: Element): SectionReportEntry["before"] {
  const pgSz = firstChildNS(sectPr, w, "pgSz")
  const pgMar = firstChildNS(sectPr, w, "pgMar")
  const widthTwips = pgSz ? parseInt(pgSz.getAttributeNS(w, "w") || "0", 10) : 0
  const heightTwips = pgSz ? parseInt(pgSz.getAttributeNS(w, "h") || "0", 10) : 0
  const orient: Orientation =
    pgSz?.getAttributeNS(w, "orient") === "landscape"
      ? "landscape"
      : widthTwips > heightTwips && widthTwips > 0
        ? "landscape"
        : "portrait"
  const readMar = (attr: string) => (pgMar ? parseInt(pgMar.getAttributeNS(w, attr) || "0", 10) : 0)
  return {
    paperSize: matchPaperLabel(widthTwips, heightTwips),
    orientation: orient,
    margins: {
      top: readMar("top"),
      bottom: readMar("bottom"),
      left: readMar("left"),
      right: readMar("right"),
    },
  }
}

/** Map a twips dimension pair back to a paper size label for display. Both
 *  orientations match; tolerance of 1 twip absorbs rounding. */
function matchPaperLabel(width: number, height: number): string | undefined {
  if (width <= 0 || height <= 0) return undefined
  const portraitW = Math.min(width, height)
  const portraitH = Math.max(width, height)
  for (const [label, [pw, ph]] of Object.entries(PAPER_SIZE_TWIPS)) {
    if (Math.abs(portraitW - pw) <= 1 && Math.abs(portraitH - ph) <= 1) return label
  }
  return undefined
}
