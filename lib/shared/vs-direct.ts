/**
 * "vs target direct" classifier for the dry-run Style Resolution report.
 *
 * The existing Δ-line compares an agent's declared style against the source
 * styles.xml's prior cascade — but the sparse-by-design invariant actually
 * lives at the **direct format** layer (the pPr / rPr literally on each
 * paragraph). A declaration of `bold: true` may be redundant because every
 * target paragraph already has `<w:b/>` direct, or it may override a direct
 * `<w:b w:val="false"/>` on those same paragraphs. The styles-cascade Δ-line
 * can't see either case.
 *
 * For each field the agent declares on a style, this classifier walks the
 * target paragraphs (post-rules-pass: any paragraph whose final pStyle is
 * the agent's styleId) and tallies:
 *   - override  — the paragraph carries a direct value that differs from
 *                 the declaration; the new style replaces it (after the
 *                 engine's selective-strip removes the direct conflict).
 *   - redundant — paragraph's direct value matches the declaration; the
 *                 declaration adds nothing new (the strip removes the
 *                 direct then the style cascade reinstates the same value).
 *   - new       — paragraph has no direct value for this field; the
 *                 declaration is a fresh addition via the style cascade.
 *
 * Mode A same-source silencing: when a style was built from a representative
 * paragraph (`fromParagraph: N`), the fields extracted necessarily match N's
 * direct values, and every paragraph sharing N's fingerprint will trivially
 * match too. Those redundant counts are noise — pass N's fingerprint here
 * and the classifier skips those paragraphs' redundant tallies (override
 * and new still count, since cross-fingerprint paragraphs are informative).
 */

import { computedStyleToEntry } from "@lib/apply/style-mutation.ts"
import type { StyleConfigEntry } from "@lib/config/config-types.ts"
import type { ParsedParagraph } from "@lib/parse/types.ts"
import { type LineSpacingInput, parseLineSpacing } from "@lib/shared/units.ts"

export type VsDirectClass = "override" | "redundant" | "new"

export interface VsDirectFieldReport {
  field: string
  override: number
  redundant: number
  fresh: number
  /** Representative direct value the agent's declaration replaces. */
  overrideFrom?: unknown
  /** Declared value (the "to" of override / value of redundant / value of new). */
  declared: unknown
  /** True when the engine's selective-strip touches this field — override
   * here just means "replaces the direct value being stripped." False when
   * the field stays on the paragraph (mixed-run rPr properties, etc.). */
  willStrip: boolean
}

export interface VsDirectReport {
  targetCount: number
  /** Per-field classification rolled up across all target paragraphs.
   * Only fields the agent declared appear; only fields with at least one
   * non-zero bucket appear. Order matches FORMAT_FIELDS for stable output. */
  fields: VsDirectFieldReport[]
}

/* ------------- which fields are subject to selective-strip -------------
 * Mirrors the predicate in `lib/apply/para-mutation.ts:stripConflictingDirectFormatting`.
 * pPr-direct strip fires for jc / spacing / ind / outlineLvl when the new
 * style declares them. Run-level rPr strip fires uniform-across-runs only
 * when the new style declares them. Mixed-run cases stay (intentional
 * inline emphasis); for simplicity this table assumes the strip will fire
 * — most chrome paragraphs are uniform. */

const STRIPPABLE_FIELDS: ReadonlySet<string> = new Set([
  // pPr
  "alignment",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "firstLineIndent",
  "hangingIndent",
  "outlineLevel",
  // rPr (uniform)
  "fontLatin",
  "fontCJK",
  "size",
  "bold",
  "italic",
  "color",
  "vertAlign",
])

/* ------------- schema-drift guard -------------
 * Every formatting field on `StyleConfigEntry` (excluding mechanical
 * fields id/name/basedOn/fromParagraph/overrides) must appear in
 * FORMAT_FIELDS. The two type constraints below make this enforceable
 * at compile time:
 *
 *   - `FormatFieldKey extends keyof StyleConfigEntry` — every entry in
 *     our union is a real schema key; adding a typo here fails to compile.
 *   - `satisfies readonly FormatFieldKey[]` — the array contains exactly
 *     declared FormatFieldKey values; the const map below is keyed by
 *     FormatFieldKey so missing entries fail to compile.
 *
 * Adding a new formatting field to the schema therefore breaks the build
 * here until the maintainer adds it to FORMAT_FIELDS too. */

type FormatFieldKey = Extract<
  keyof StyleConfigEntry,
  | "fontLatin"
  | "fontCJK"
  | "size"
  | "bold"
  | "italic"
  | "color"
  | "vertAlign"
  | "alignment"
  | "lineSpacing"
  | "spaceBefore"
  | "spaceAfter"
  | "firstLineIndent"
  | "hangingIndent"
  | "outlineLevel"
>

const FORMAT_FIELDS = [
  "fontLatin",
  "fontCJK",
  "size",
  "bold",
  "italic",
  "color",
  "vertAlign",
  "alignment",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "firstLineIndent",
  "hangingIndent",
  "outlineLevel",
] as const satisfies readonly FormatFieldKey[]

/* ------------- analyzer ------------- */

export function analyzeVsDirect(
  declared: StyleConfigEntry,
  targets: ParsedParagraph[],
  fromParagraphFingerprint?: string,
): VsDirectReport {
  const fields: VsDirectFieldReport[] = []

  for (const field of FORMAT_FIELDS) {
    const declaredRaw = declared[field]
    if (declaredRaw === undefined) continue
    const declaredCanonical = canonicalize(field, declaredRaw)

    let override = 0
    let redundant = 0
    let fresh = 0
    let overrideFrom: unknown

    for (const p of targets) {
      const directEntry = computedStyleToEntry(p.directRPr, p.directPPr)
      const directRaw = directEntry[field]
      if (directRaw === undefined) {
        fresh++
        continue
      }
      const directCanonical = canonicalize(field, directRaw)
      if (sameValue(directCanonical, declaredCanonical)) {
        // Mode A same-source silence: skip the redundant tally for the
        // paragraph that fed the extraction (and any sharing its fingerprint)
        // — necessarily matches by construction, not real signal.
        if (fromParagraphFingerprint && p.fingerprint === fromParagraphFingerprint) {
          continue
        }
        redundant++
      } else {
        override++
        if (overrideFrom === undefined) overrideFrom = directCanonical
      }
    }

    if (override === 0 && redundant === 0 && fresh === 0) continue
    fields.push({
      field,
      override,
      redundant,
      fresh,
      overrideFrom,
      declared: declaredCanonical,
      willStrip: STRIPPABLE_FIELDS.has(field),
    })
  }

  return { targetCount: targets.length, fields }
}

function canonicalize(field: string, val: unknown): unknown {
  // lineSpacing accepts number (multiplier) | "Npt" (exact) | { atLeast }
  // at the config surface. Normalize to the parsed mode+value tuple so the
  // three forms compare structurally.
  if (field === "lineSpacing") {
    const ls = parseLineSpacing(val as LineSpacingInput, "lineSpacing")
    return `${ls.mode}:${ls.value}`
  }
  // spaceBefore / spaceAfter / firstLineIndent / hangingIndent / size: any
  // Length / Indent form compares by stringified form for now; the engine
  // already normalizes to twips/half-pt at emit time.
  return val
}

/** Shallow value equality for restyle / report comparisons. Numbers
 * compare with a 1e-6 tolerance to absorb twip round-trip noise; other
 * types stringify and compare. Exported so `report.ts` and other
 * comparators stay aligned on what "same" means. */
export function sameValue(a: unknown, b: unknown): boolean {
  if (typeof a === "number" && typeof b === "number") {
    return Math.abs(a - b) < 1e-6
  }
  return String(a) === String(b)
}
