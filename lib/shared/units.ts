/**
 * Single source of truth for length-unit parsing and conversion.
 *
 * Agent-facing API exposes pt / cm / mm / in via string suffixes
 * (`"12pt" / "2.54cm" / "5mm" / "1in"`). Bare numbers are pt. Internal
 * OOXML units (twips, half-pt, EMU, eighth-pt) live behind dedicated
 * conversion helpers — callers never multiply by 20 / 2 / 12700 / 8 directly.
 *
 * Two narrower types extend the basic Length:
 *   - `IndentValue` adds `"Nchar"` (round-trips Word's `w:firstLineChars` /
 *     `w:hangingChars`, auto-scales with font size).
 *   - `LineSpacingValue` carries three Word `lineRule` modes via type
 *     discrimination: number = multiplier (auto), Length-string = exact
 *     fixed height, `{ atLeast: Length }` = at-least fixed height.
 *     The previous magnitude heuristic (`value >= 10 → exact`) is gone.
 */

const TWIPS_PER_PT = 20
const HALF_PT_PER_PT = 2
const EIGHTH_PT_PER_PT = 8
const EMU_PER_PT = 12700
const TWENTIETHS_PER_LINE = 240

const PT_PER_CM = 28.3464566929 // 1 cm = 1/2.54 in × 72 pt
const PT_PER_MM = 2.83464566929
const PT_PER_IN = 72

const SUPPORTED_LENGTH_UNITS = "pt, cm, mm, in"

/** Public length input form. Bare number is pt. */
export type Length = number | string

export type IndentInput = number | string | null
export type LineSpacingInput = number | string | { atLeast: Length }

/** CSS-style padding input. 1 / [all] / [v, h] / [t, h, b] / [t, r, b, l]. */
export type PaddingInput =
  | Length
  | [Length]
  | [Length, Length]
  | [Length, Length, Length]
  | [Length, Length, Length, Length]

export type PaddingEdges = { top: number; right: number; bottom: number; left: number }

/** Result of parsing an IndentInput: `char` preserves font-size scaling
 *  (Word's `firstLineChars` × 100); `twip` is a fixed indent. */
export type IndentParsed = { kind: "char"; value: number } | { kind: "twip"; value: number }

/** Result of parsing a LineSpacingInput. `value` is in OOXML's units for
 *  the chosen mode: 240ths-of-line for `auto`, twips for `exact` /
 *  `atLeast`. Emitters set `w:lineRule` from `mode` and `w:line` from
 *  `value`. */
export type LineSpacingParsed = {
  mode: "auto" | "exact" | "atLeast"
  value: number
}

/* ----------------------------- core parse ----------------------------- */

/** Parse a Length input to pt. Bare number is pt; strings carry one of
 *  pt / cm / mm / in. Throws with a clear message on bad input. */
export function parseLengthPt(v: Length, fieldName = "length"): number {
  if (typeof v === "number") {
    if (!Number.isFinite(v)) {
      throw new Error(`${fieldName}: ${v} is not a finite number`)
    }
    return v
  }
  const m = v.trim().match(/^(-?\d+(?:\.\d+)?)\s*(pt|cm|mm|in)$/i)
  if (!m) {
    throw new Error(
      `${fieldName} "${v}": expected a number (pt) or a string like "12pt" / "2.54cm" / "5mm" / "1in". Supported units: ${SUPPORTED_LENGTH_UNITS}.`,
    )
  }
  const n = parseFloat(m[1]!)
  const unit = m[2]!.toLowerCase()
  switch (unit) {
    case "pt":
      return n
    case "cm":
      return n * PT_PER_CM
    case "mm":
      return n * PT_PER_MM
    case "in":
      return n * PT_PER_IN
  }
  // Unreachable — regex pins the unit set.
  throw new Error(`${fieldName} "${v}": unsupported unit "${unit}".`)
}

/** Parse an indent input into its OOXML-ready tagged unit.
 *
 *   number          → pt → twips (`w:firstLine` / `w:hanging`)
 *   "Npt|Ncm|..."   → pt → twips
 *   "Nchar"         → 1/100 char (`w:firstLineChars` / `w:hangingChars`)
 *   null            → null (caller drops the attribute)
 */
export function parseIndent(v: IndentInput, fieldName = "indent"): IndentParsed | null {
  if (v === null) return null
  if (typeof v === "number") {
    if (!Number.isFinite(v)) {
      throw new Error(`${fieldName}: ${v} is not a finite number`)
    }
    return { kind: "twip", value: Math.round(v * TWIPS_PER_PT) }
  }
  const trimmed = v.trim()
  const charMatch = trimmed.match(/^(-?\d+(?:\.\d+)?)\s*chars?$/i)
  if (charMatch) {
    return { kind: "char", value: Math.round(parseFloat(charMatch[1]!) * 100) }
  }
  // Fall through to length parsing for pt/cm/mm/in.
  const pt = parseLengthPt(trimmed, fieldName)
  return { kind: "twip", value: Math.round(pt * TWIPS_PER_PT) }
}

/** Parse a lineSpacing input into mode + OOXML-ready value.
 *
 *   number          → multiplier (auto), value in 240ths-of-line
 *   "Npt|Ncm|..."   → exact line height, value in twips
 *   { atLeast: L }  → at-least line height, value in twips
 */
export function parseLineSpacing(
  v: LineSpacingInput,
  fieldName = "lineSpacing",
): LineSpacingParsed {
  if (typeof v === "number") {
    if (!Number.isFinite(v) || v <= 0) {
      throw new Error(`${fieldName}: ${v} is not a positive finite multiplier`)
    }
    return { mode: "auto", value: Math.round(v * TWENTIETHS_PER_LINE) }
  }
  if (typeof v === "string") {
    const pt = parseLengthPt(v, fieldName)
    return { mode: "exact", value: Math.round(pt * TWIPS_PER_PT) }
  }
  if (v && typeof v === "object" && "atLeast" in v) {
    const pt = parseLengthPt(v.atLeast, `${fieldName}.atLeast`)
    return { mode: "atLeast", value: Math.round(pt * TWIPS_PER_PT) }
  }
  throw new Error(
    `${fieldName}: expected a number (multiplier, e.g. 1.5), a string (exact, e.g. "24pt"), or { atLeast: "<length>" }.`,
  )
}

/** Expand CSS-shorthand padding to four edges in pt. Throws on negative
 *  values — OOXML cell margins are `xsd:unsignedDecimalNumber`.
 *  Error labels reflect CSS shorthand semantics: 2-value reports
 *  `.vertical` / `.horizontal`, 3-value reports `.top` / `.horizontal` /
 *  `.bottom`, 4-value reports `.top` / `.right` / `.bottom` / `.left`. */
export function parsePadding(v: PaddingInput, fieldName = "padding"): PaddingEdges {
  const parts: Length[] = Array.isArray(v) ? (v as Length[]) : [v]
  const at = (i: number, label: string) =>
    parseLengthPt(parts[i]!, label ? `${fieldName}.${label}` : fieldName)
  let edges: PaddingEdges
  if (parts.length === 1) {
    const all = at(0, "")
    edges = { top: all, right: all, bottom: all, left: all }
  } else if (parts.length === 2) {
    const vert = at(0, "vertical")
    const horiz = at(1, "horizontal")
    edges = { top: vert, right: horiz, bottom: vert, left: horiz }
  } else if (parts.length === 3) {
    const horiz = at(1, "horizontal")
    edges = { top: at(0, "top"), right: horiz, bottom: at(2, "bottom"), left: horiz }
  } else if (parts.length === 4) {
    edges = {
      top: at(0, "top"),
      right: at(1, "right"),
      bottom: at(2, "bottom"),
      left: at(3, "left"),
    }
  } else {
    throw new Error(`${fieldName}: tuple must have 1-4 entries, got ${parts.length}`)
  }
  for (const side of ["top", "right", "bottom", "left"] as const) {
    if (edges[side] < 0) {
      throw new Error(
        `${fieldName}.${side}: ${edges[side]} is negative; OOXML cell margins must be non-negative`,
      )
    }
  }
  return edges
}

/* --------------------------- unit conversions --------------------------- */

export function toTwips(v: Length, fieldName?: string): number {
  return Math.round(parseLengthPt(v, fieldName) * TWIPS_PER_PT)
}

export function toHalfPt(v: Length, fieldName?: string): number {
  return Math.round(parseLengthPt(v, fieldName) * HALF_PT_PER_PT)
}

export function toEighthPt(v: Length, fieldName?: string): number {
  return Math.round(parseLengthPt(v, fieldName) * EIGHTH_PT_PER_PT)
}

export function toEmu(v: Length, fieldName?: string): number {
  return Math.round(parseLengthPt(v, fieldName) * EMU_PER_PT)
}

/* --------------------------- inverse helpers --------------------------- */

/** Format twips back into a `"Npt"` string for round-tripping engine state
 *  into config-shaped output (dry-run reports, fromParagraph extraction). */
export function twipsToPtString(twips: number): string {
  return `${twips / TWIPS_PER_PT}pt`
}

/** Format twips into a `"N.NNcm"` display string for dry-run / report
 *  contexts where users think in cm (margins, paper sizes). Two-decimal
 *  precision matches Word's UI display granularity. */
export function twipsToCmString(twips: number): string {
  return `${(twips / (TWIPS_PER_PT * PT_PER_CM)).toFixed(2)}cm`
}

/** Half-pt → pt (number form, since size fields take bare pt). */
export function halfPtToPt(halfPt: number): number {
  return halfPt / HALF_PT_PER_PT
}

/** Char-units (×100, as Word stores `firstLineChars`) → "Nchar" string. */
export function charUnitsToString(charUnits: number): string {
  return `${charUnits / 100}char`
}

/** Inverse of parseLineSpacing's value, given the mode it was parsed under.
 *  Returns the config-shaped value that would round-trip back. */
export function lineSpacingToConfig(parsed: LineSpacingParsed): LineSpacingInput {
  switch (parsed.mode) {
    case "auto":
      return parsed.value / TWENTIETHS_PER_LINE
    case "exact":
      return twipsToPtString(parsed.value)
    case "atLeast":
      return { atLeast: twipsToPtString(parsed.value) }
  }
}
