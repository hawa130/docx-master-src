/**
 * Display helpers shared across the inspector CLIs. The arithmetic for
 * size / spacing / indent units is identical across tools; only the
 * field selection / toggle predicate differs, so the rPr / pPr renderers
 * here return a parts array and accept option flags for per-tool variants.
 * Callers wrap the parts with `{ ... }` and append any tool-specific
 * extras (underline / highlight / strike / caps / etc.).
 */

import type { ComputedParaStyle, ComputedRunStyle } from "@lib/parse/types.ts"

/** Pad a 1-based paragraph index to 3 digits — "001", "042", "123". */
export function pad(n: number): string {
  return n.toString().padStart(3, "0")
}

/** Collapse internal whitespace and clip with an ellipsis. */
export function truncate(s: string, n: number): string {
  const collapsed = s.replace(/\s+/g, " ").trim()
  return collapsed.length <= n ? collapsed : collapsed.slice(0, n) + "…"
}

/** Recognize standard paper sizes by their twip dimensions. */
export function paperName(width: number, height: number): string {
  // both portrait and landscape A4 are listed so either orientation matches
  const known: Array<[string, number, number]> = [
    ["A4", 11906, 16838],
    ["A4", 16838, 11906],
    ["A3", 16838, 23811],
    ["A5", 8392, 11906],
    ["Letter", 12240, 15840],
    ["Legal", 12240, 20160],
  ]
  for (const [name, w, h] of known) {
    if (Math.abs(width - w) < 50 && Math.abs(height - h) < 50) return name
  }
  return "Custom"
}

/** Render Word's spacing/@line + spacing/@lineRule pair as human text. */
export function formatLineSpacing(line: number, rule: string | undefined): string {
  const r = rule || "auto"
  if (r === "exact") return `${line / 20}pt fixed`
  if (r === "atLeast") return `${line / 20}pt atLeast`
  return `${parseFloat((line / 240).toFixed(2))}×`
}

/** Twips → millimeters, rounded to 1dp (Word page-setup-friendly). */
export function tw2mm(t: number): number {
  return +(t / 56.6929).toFixed(1)
}

export interface RPrFormatOptions {
  /** Render bold / italic toggles only when truthy. Default: render when defined
   * (including false), matching `inspect_range` / `inspect_style_def`'s convention
   * of showing direct overrides explicitly. */
  truthyToggles?: boolean
  /** Skip color when value is `"auto"` (the cascade default). Default false. */
  filterAutoColor?: boolean
}

/** Render the common subset of `ComputedRunStyle` as a parts array — fonts,
 * size, bold, italic, color, vertAlign. Caller wraps with `{ ... }` and
 * appends tool-specific extras (underline / highlight / strike / caps). */
export function formatComputedRPrParts(r: ComputedRunStyle, opts: RPrFormatOptions = {}): string[] {
  const parts: string[] = []
  const truthy = opts.truthyToggles ?? false
  const pushToggle = (name: string, v: boolean | undefined) => {
    if (v === undefined) return
    if (truthy && !v) return
    parts.push(truthy ? `${name}: true` : `${name}: ${v}`)
  }
  const latin = r.fontAscii ?? r.fontHAnsi
  const setSlots = [latin, r.fontEastAsia].filter((v): v is string => v !== undefined)
  if (setSlots.length >= 2 && setSlots.every((v) => v === setSlots[0])) {
    parts.push(`font: "${setSlots[0]}"`)
  } else {
    if (r.fontEastAsia) parts.push(`fontCJK: "${r.fontEastAsia}"`)
    if (latin) parts.push(`fontLatin: "${latin}"`)
  }
  if (r.size !== undefined) parts.push(`size: ${r.size / 2}pt`)
  pushToggle("bold", r.bold)
  pushToggle("italic", r.italic)
  if (r.color && (!opts.filterAutoColor || r.color !== "auto")) parts.push(`color: ${r.color}`)
  if (r.vertAlign) parts.push(`vertAlign: ${r.vertAlign}`)
  return parts
}

export interface PPrFormatOptions {
  /** Include `indentLeft` / `indentRight` when defined. Default false. */
  includeIndentSides?: boolean
  /** Include `numLevel` when defined. Default false. */
  includeNumLevel?: boolean
  /** Explicit numId display string. When provided, used instead of
   * `pp.numId`. Pass `null` to suppress numId entirely (e.g. when style-level
   * cross-paragraph resolution decided no consistent numId applies). */
  numIdDisplay?: string | null
}

/** Render the common subset of `ComputedParaStyle` as a parts array. */
export function formatComputedPPrParts(
  pp: ComputedParaStyle,
  opts: PPrFormatOptions = {},
): string[] {
  const parts: string[] = []
  if (pp.alignment) parts.push(`alignment: ${pp.alignment}`)
  if (pp.spaceBefore !== undefined) parts.push(`spaceBefore: ${pp.spaceBefore / 20}pt`)
  if (pp.spaceAfter !== undefined) parts.push(`spaceAfter: ${pp.spaceAfter / 20}pt`)
  if (pp.lineSpacing !== undefined) {
    parts.push(`lineSpacing: ${formatLineSpacing(pp.lineSpacing, pp.lineRule)}`)
  }
  if (opts.includeIndentSides) {
    if (pp.indentLeft !== undefined) parts.push(`indentLeft: ${pp.indentLeft / 20}pt`)
    if (pp.indentRight !== undefined) parts.push(`indentRight: ${pp.indentRight / 20}pt`)
  }
  if (pp.firstLineIndentChars !== undefined) {
    parts.push(`firstLineIndent: ${pp.firstLineIndentChars / 100}char`)
  } else if (pp.firstLineIndent !== undefined) {
    parts.push(`firstLineIndent: ${pp.firstLineIndent / 20}pt`)
  }
  if (pp.hangingIndentChars !== undefined) {
    parts.push(`hangingIndent: ${pp.hangingIndentChars / 100}char`)
  } else if (pp.hangingIndent !== undefined) {
    parts.push(`hangingIndent: ${pp.hangingIndent / 20}pt`)
  }
  if (pp.outlineLevel !== undefined) parts.push(`outlineLevel: ${pp.outlineLevel}`)
  const numId =
    opts.numIdDisplay === undefined ? (pp.numId ? String(pp.numId) : null) : opts.numIdDisplay
  if (numId) parts.push(`numId: ${numId}`)
  if (opts.includeNumLevel && pp.numLevel !== undefined) {
    parts.push(`numLevel: ${pp.numLevel}`)
  }
  return parts
}
