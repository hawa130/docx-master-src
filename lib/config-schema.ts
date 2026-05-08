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

const NonEmptyString = z.string().check(z.minLength(1))

/** Indent value: "Nchar" / "Npt" / fixed pt number / null. */
const IndentValue = z.union([z.string(), z.number(), z.null()])

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
  lineSpacing: z.optional(z.number()),
  lineRule: z.optional(z.enum(["auto", "exact", "atLeast"])),
  spaceBefore: z.optional(z.number()),
  spaceAfter: z.optional(z.number()),
  firstLineIndent: z.optional(IndentValue),
  hangingIndent: z.optional(IndentValue),
  outlineLevel: z.optional(z.number().check(z.gte(0), z.lte(9))),
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
  assignments: z.optional(z.array(AssignmentSchema)),
  bulk_rules: z.optional(z.array(BulkRuleSchema)),
  pattern_rules: z.optional(z.array(PatternRuleSchema)),
  requirements: z.optional(z.record(z.string(), z.string())),
  exclude: z.optional(z.array(z.number())),
})

export type ApplyConfig = z.infer<typeof ApplyConfigSchema>

/* ------------- error formatting ------------- */

/** Convert a zod issue.path (PropertyKey[]) to dotted/bracketed JS form,
 * e.g. `["numbering","levels",0,"styleId"]` → `numbering.levels[0].styleId`. */
function formatPath(path: readonly PropertyKey[]): string {
  let out = ""
  for (const seg of path) {
    if (typeof seg === "number") out += `[${seg}]`
    else out += out ? `.${String(seg)}` : String(seg)
  }
  return out
}

/** Walk `path` into `raw` and return the value at that location, or undefined
 * if any segment is missing. Used so error messages can echo the actual
 * value the agent supplied — `issue.input` is unreliable across issue codes
 * in zod mini, the raw config is always the ground truth. */
function valueAtPath(raw: unknown, path: readonly PropertyKey[]): unknown {
  let cur: unknown = raw
  for (const seg of path) {
    if (cur == null || typeof cur !== "object") return undefined
    cur = (cur as Record<PropertyKey, unknown>)[seg]
  }
  return cur
}

/** Domain-specific hints that override or augment zod's default message for
 * particular issue shapes. Returning null means "use the default zod
 * message". The hints replicate the agent-friendly guidance the hand-written
 * checks used to emit. */
function customHint(issue: z.core.$ZodIssue, pathStr: string, raw: unknown): string | null {
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

/** Format a single zod issue into an agent-readable line. Falls back to the
 * issue's default message only when no specific code matches; for the common
 * codes we render concrete details (unknown keys, expected types, allowed
 * enum values) instead of the bare "Invalid input" zod emits at low cost. */
function formatIssue(issue: z.core.$ZodIssue, pathStr: string, raw: unknown): string {
  const got = valueAtPath(raw, issue.path)
  const gotStr = got === undefined ? "(missing)" : JSON.stringify(got)
  switch (issue.code) {
    case "unrecognized_keys": {
      const keys = (issue as { keys?: string[] }).keys ?? []
      const noun = keys.length === 1 ? "key" : "keys"
      return `unknown ${noun} ${keys.map((k) => `"${k}"`).join(", ")}`
    }
    case "invalid_type": {
      const expected = (issue as { expected?: string }).expected ?? "value"
      return got === undefined
        ? `missing required field (expected ${expected})`
        : `expected ${expected}, got ${gotStr}`
    }
    case "invalid_value": {
      const allowed = (issue as { values?: readonly unknown[] }).values ?? []
      return `invalid value ${gotStr}. Allowed: [${allowed.map((v) => JSON.stringify(v)).join(", ")}]`
    }
    case "too_small": {
      const min = (issue as { minimum?: unknown }).minimum
      const origin = (issue as { origin?: string }).origin
      if (origin === "string") return `must be a non-empty string (got ${gotStr})`
      if (origin === "array") return `must contain at least ${String(min)} item(s)`
      return `value too small (minimum ${String(min)})`
    }
    case "too_big": {
      const max = (issue as { maximum?: unknown }).maximum
      return `value too large (maximum ${String(max)})`
    }
    default:
      return issue.message
  }
}

/** Format a ZodError as a multi-line, agent-readable string. Each issue
 * gets a path prefix and either a custom hint, a code-specific render, or
 * zod's default message as a final fallback. */
export function formatConfigError(error: z.core.$ZodError, raw: unknown): string {
  const lines: string[] = []
  for (const issue of error.issues) {
    const pathStr = formatPath(issue.path)
    const hint = customHint(issue, pathStr, raw)
    const msg = hint ?? formatIssue(issue, pathStr, raw)
    lines.push(pathStr ? `${pathStr}: ${msg}` : msg)
  }
  return lines.join("\n")
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
