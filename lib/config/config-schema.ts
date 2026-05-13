/**
 * Zod-mini schema for the apply_styles family CLIs (apply_styles, restyle,
 * migrate_numbering, import_template). Replaces the hand-written `typeof` /
 * `Set.has()` validation that previously lived inline in apply-styles.ts.
 *
 * What lives here:
 *   - shape (required vs optional, types, enums)
 *   - rejection of unknown keys (strictObject everywhere)
 *   - non-empty constraints on critical strings (id / name / lvlText / ...)
 *   - pure-config invariants that don't need runtime docx context
 *     (e.g. stripPrefixPatterns must be ordered by descending placeholder
 *     count — fully derivable from the strings themselves)
 *
 * What does NOT live here:
 *   - cross-field checks that need the parsed docx (styleId existence in
 *     stylesDoc, paragraph-index validity, fingerprint resolution, name
 *     collision detection). Those stay in apply-styles.ts where the runtime
 *     context is available.
 *   - tool-narrowing rules ("restyle rejects template / numbering"). Those
 *     stay in each CLI's `validate` callback because the message format is
 *     tool-aware ("use import_template instead").
 *   - regex compilation for pattern_rules. The compiled RegExp is consumed
 *     downstream, so compile-once-and-store stays in apply-styles.ts.
 *
 * The error formatter (`formatConfigError`) maps known zod issues to
 * domain-specific hints so the agent gets the same "stripPrefixPatterns
 * belongs INSIDE each level" / "Allowed: [tab, space, nothing]. Omit to
 * auto-infer..." guidance as before. Generic issues fall back to zod's
 * default messages with paths.
 */

import * as z from "zod/mini"

/* ------------- atomic helpers ------------- */

import { IndentValue, NonEmptyString } from "@lib/config/zod-primitives.ts"

/** Fields shared between Mode B (`styles[i].*`) direct values and the
 * `styles[i].overrides` block. Kept as a plain object so it can be spread
 * into both schemas. */
const styleFormatFields = {
  basedOn: z.optional(z.string()),
  fontLatin: z.optional(z.string()),
  fontCJK: z.optional(z.string()),
  size: z.optional(z.number()),
  bold: z.optional(z.boolean()),
  italic: z.optional(z.boolean()),
  color: z.optional(z.string()),
  alignment: z.optional(z.enum(["left", "center", "right", "both"])),
  lineSpacing: z.optional(z.union([z.number(), z.string()])),
  lineRule: z.optional(z.enum(["auto", "exact", "atLeast"])),
  spaceBefore: z.optional(z.number()),
  spaceAfter: z.optional(z.number()),
  firstLineIndent: z.optional(IndentValue),
  hangingIndent: z.optional(IndentValue),
  outlineLevel: z.optional(z.number().check(z.gte(0), z.lte(9))),
  // Run-level baseline shift. "baseline" is the explicit reset — distinct
  // from omitting the field (= inherit). Use "baseline" only when the
  // basedOn cascade declares super/sub and this style needs to opt out.
  vertAlign: z.optional(z.enum(["superscript", "subscript", "baseline"])),
}

/* ------------- styles[] entry ------------- */

export const StyleOverridesSchema = z.strictObject(styleFormatFields)

export const StyleEntrySchema = z.strictObject({
  id: NonEmptyString,
  name: NonEmptyString,
  fromParagraph: z.optional(z.number()),
  ...styleFormatFields,
  overrides: z.optional(StyleOverridesSchema),
})

/* ------------- numbering ------------- */

export const NumRPrSchema = z.strictObject({
  fontLatin: z.optional(z.string()),
  fontCJK: z.optional(z.string()),
  size: z.optional(z.number()),
  bold: z.optional(z.boolean()),
  italic: z.optional(z.boolean()),
  color: z.optional(z.string()),
})

export const NumLevelSchema = z
  .strictObject({
    level: z.number().check(z.gte(0), z.lte(8)),
    numFmt: NonEmptyString,
    lvlText: NonEmptyString,
    styleId: NonEmptyString,
    start: z.optional(z.number()),
    stripPrefixPatterns: z.optional(z.array(z.string())),
    suff: z.optional(z.enum(["tab", "space", "nothing"])),
    numRPr: z.optional(NumRPrSchema),
    // Force every cross-level placeholder in this lvl's lvlText to render as
    // arabic numerals regardless of the referenced level's numFmt. Required
    // for "X.X" style headings where Heading1 uses chineseCounting (一、) but
    // Heading3's lvlText "%1.%3" should display "1.1" not "一.1".
    isLgl: z.optional(z.boolean()),
    // Counter scope for single-level schemes. Default "continuous" — one
    // counter shared by every paragraph bound to this level, matching native
    // OOXML <w:num> semantics. Use "perInstance" only for procedural list
    // shapes (ListNumber / ListBullet style 1./2./3. lists) where each
    // contiguous run of same-styleId paragraphs should restart at 1; the
    // engine forks a fresh numId per run and writes <w:startOverride>.
    // Ignored on multi-level schemes (which use lvlRestart for resets).
    restart: z.optional(z.enum(["continuous", "perInstance"])),
  })
  // Strip patterns are tried in array order, first match wins. If a shorter
  // pattern (fewer %N placeholders) appears before a longer one, the shorter
  // one swallows the prefix the longer one wanted (e.g. ["%1.", "%1.%2"]
  // strips just "1." from "1.1 ..." leaving ".1 ..."). Catch at config time.
  .check(
    z.refine(
      (lvl) => {
        const counts = (lvl.stripPrefixPatterns ?? []).map((p) => (p.match(/%\d/g) ?? []).length)
        for (let j = 1; j < counts.length; j++) {
          if (counts[j]! > counts[j - 1]!) return false
        }
        return true
      },
      {
        error: (issue) => {
          const lvl = issue.input as {
            stripPrefixPatterns?: string[]
          }
          const patterns = lvl.stripPrefixPatterns ?? []
          const counts = patterns.map((p) => (p.match(/%\d/g) ?? []).length)
          for (let j = 1; j < counts.length; j++) {
            if (counts[j]! > counts[j - 1]!) {
              return (
                `stripPrefixPatterns must be ordered by descending placeholder count — ` +
                `"${patterns[j - 1]}" (${counts[j - 1]} placeholders) before "${patterns[j]}" (${counts[j]} placeholders) ` +
                `means the shorter pattern matches first and strips a prefix the longer one wanted. ` +
                `Reorder so longer patterns come first, e.g. ["%1.%2", "%1."].`
              )
            }
          }
          return "stripPrefixPatterns ordering invalid"
        },
      },
    ),
  )

export const NumberingSchema = z.strictObject({
  levels: z.array(NumLevelSchema).check(z.minLength(1)),
})

/* ------------- captions ------------- */

/** Numeric format for SEQ counters. Maps to Word's `\*` switches —
 * see lib/edit/fields/seq-field.ts. */
export const CaptionFormatSchema = z.enum([
  "arabic",
  "alphabetic",
  "ALPHABETIC",
  "roman",
  "ROMAN",
  "chinese",
  "chinese-formal",
])

export const CaptionSubCounterSchema = z.strictObject({
  format: z.optional(CaptionFormatSchema),
  prefix: z.optional(z.string()),
  suffix: z.optional(z.string()),
})

/** Per-identifier caption configuration. Block-level `captionId`
 * references the key. `styleId` is required — every caption identifier
 * must declare what paragraph style to use; no implicit fallback. */
export const CaptionEntrySchema = z.strictObject({
  prefix: z.optional(z.string()),
  suffix: z.optional(z.string()),
  format: z.optional(CaptionFormatSchema),
  chapterPrefix: z.optional(z.array(NonEmptyString)),
  chapterSeparator: z.optional(z.string()),
  bodySeparator: z.optional(z.string()),
  styleId: NonEmptyString,
  subCounter: z.optional(CaptionSubCounterSchema),
})

/** Map of identifier → caption entry. Identifier is a free string;
 * conventional values "Equation" / "Figure" / "Table" / "Theorem" live
 * in ref docs, not the schema. */
export const CaptionsSchema = z.record(NonEmptyString, CaptionEntrySchema)

/* ------------- template ------------- */

export const TemplateSchema = z.strictObject({
  source: NonEmptyString,
  styles: z.array(NonEmptyString).check(z.minLength(1)),
  importNumbering: z.optional(z.boolean()),
})

/* ------------- theme ------------- */

export const ThemeFontsSchema = z.strictObject({
  majorLatin: z.optional(z.string()),
  majorEastAsia: z.optional(z.string()),
  minorLatin: z.optional(z.string()),
  minorEastAsia: z.optional(z.string()),
})

export const ThemeSchema = z.strictObject({
  fonts: z.optional(ThemeFontsSchema),
})

/* ------------- paragraph mapping ------------- */

export const AssignmentSchema = z.strictObject({
  para: z.number(),
  action: z.enum(["keep", "restyle", "flag"]),
  style: z.optional(z.string()),
  reason: z.optional(z.string()),
})

export const BulkRuleSchema = z.strictObject({
  fingerprint: NonEmptyString,
  style: NonEmptyString,
})

export const PatternRuleSchema = z.strictObject({
  regex: NonEmptyString,
  flags: z.optional(z.string()),
  style: NonEmptyString,
  stripMatch: z.optional(z.boolean()),
})

/* ------------- top-level apply config ------------- */

// edits live in lib/edit-config-schema.ts; reference by import to avoid circular
// engine dependencies. The schema is reused as-is, just embedded inside
// ApplyConfig as an optional block. When present, the engine runs edit ops
// after style + numbering install and before pattern_rules / bulk_rules /
// assignments — so the rules pass sees both pre-existing chrome paragraphs
// and any agent-inserted content uniformly.
import { EditOpSchema } from "@lib/config/edit-config-schema.ts"

export const ApplyConfigSchema = z.strictObject({
  source: NonEmptyString,
  output: NonEmptyString,
  dryRun: z.optional(z.boolean()),
  template: z.optional(TemplateSchema),
  theme: z.optional(ThemeSchema),
  styles: z.optional(z.array(StyleEntrySchema)),
  // Single scheme (most common: one multi-level scheme bound to Heading1–N)
  // or an array of schemes when the doc needs multiple parallel ones (e.g.
  // a multi-level heading scheme + a single-level list-bound scheme). The
  // engine processes them in array order, allocating fresh numIds for each.
  numbering: z.optional(z.union([NumberingSchema, z.array(NumberingSchema)])),
  captions: z.optional(CaptionsSchema),
  assignments: z.optional(z.array(AssignmentSchema)),
  bulk_rules: z.optional(z.array(BulkRuleSchema)),
  pattern_rules: z.optional(z.array(PatternRuleSchema)),
  requirements: z.optional(z.record(z.string(), z.string())),
  exclude: z.optional(z.array(z.number())),
  edits: z.optional(z.array(EditOpSchema)),
  trackChanges: z.optional(z.boolean()),
})

export type ApplyConfig = z.infer<typeof ApplyConfigSchema>

/* ------------- error formatting ------------- */

import { formatZodError, valueAtPath, type HintFn } from "@lib/config/zod-format.ts"

/** Domain-specific hints that override or augment zod's default message for
 * particular issue shapes. Returning null means "use the default zod
 * message". The hints replicate the agent-friendly guidance the hand-written
 * checks used to emit. */
const customHint: HintFn = (issue, pathStr, raw) => {
  // Top-of-numbering misplacement: stripPrefixPatterns naturally feels like
  // it should broadcast across all levels, but it lives per-level. The old
  // hand-written validator surfaced this; preserve the message verbatim.
  if (
    issue.code === "unrecognized_keys" &&
    pathStr === "numbering" &&
    "keys" in issue &&
    Array.isArray(issue.keys) &&
    issue.keys.includes("stripPrefixPatterns")
  ) {
    return (
      `unknown field "stripPrefixPatterns" on numbering — ` +
      `stripPrefixPatterns belongs INSIDE each level (numbering.levels[i].stripPrefixPatterns), not at the top of numbering.`
    )
  }
  // Suff enum mismatch: spell out the allowed set and the omit-to-infer
  // escape hatch. Without this the agent only sees "invalid value" and may
  // not realise the field is optional.
  if (issue.code === "invalid_value" && pathStr.endsWith(".suff") && "values" in issue) {
    const allowed = (issue as { values: readonly unknown[] }).values
    const got = valueAtPath(raw, issue.path)
    return (
      `invalid suff: ${JSON.stringify(got)}. ` +
      `Allowed: [${allowed.map((v) => String(v)).join(", ")}]. ` +
      `Omit to auto-infer from trailing whitespace in lvlText (0 spaces → nothing, 1 → space, 2+ → tab).`
    )
  }
  return null
}

export function formatConfigError(error: z.core.$ZodError, raw: unknown): string {
  return formatZodError(error, raw, customHint)
}

/** Parse + validate a raw JSON config. Throws an Error with a multi-line,
 * path-annotated message on shape failure. The thrown message preserves the
 * agent-friendly hints the previous hand-written validator emitted. */
export function parseConfig(raw: unknown): ApplyConfig {
  const result = z.safeParse(ApplyConfigSchema, raw)
  if (!result.success) {
    throw new Error(formatConfigError(result.error, raw))
  }
  return result.data
}
