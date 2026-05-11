/**
 * Zod-mini schema for the `edit` command. Mirrors the config-schema.ts
 * conventions: strictObject everywhere, NonEmptyString helper, custom
 * formatter that maps known issue shapes to agent-friendly hints.
 *
 * Single source of truth for the EditConfig shape. The TS types in
 * edit-types.ts are inferred from these schemas, so adding/renaming a
 * field flows through to every consumer's compile-time check.
 *
 * What lives here:
 *   - shape (required vs optional, types, enums, regex constraints)
 *   - rejection of unknown keys (strictObject everywhere)
 *   - locally-derivable invariants (`range.from <= range.to`, color hex
 *     format, level bounds 0..8, image dimensions positive)
 *
 * What does NOT live here (left to edit-engine / locator / fragment-emit):
 *   - cross-doc validation: paragraph indices in range, table count,
 *     numId existence, styleId existence — needs the parsed docx
 *   - locator reachability under blockers — needs blocker scan results
 *   - track-changes-mode constraint enforcement (none currently; Phase 1
 *     supports tracking format ops via rPrChange / pPrChange too)
 */

import * as z from "zod/mini"

/* ------------- atomic helpers ------------- */

const NonEmptyString = z.string().check(z.minLength(1))

/** Indent value: "Nchar" / "Npt" / fixed pt number / null. Matches the
 * convention in lib/config-schema.ts so configs are uniform across
 * standardize and edit. */
const IndentValue = z.union([z.string(), z.number(), z.null()])

/** RGB hex without leading "#". Six-hex form, case-insensitive. Same shape
 * as styles in lib/config-schema.ts (which is laxer there for legacy
 * reasons; we tighten here since this is a new surface). */
const ColorHex = z.string().check(z.regex(/^[0-9a-fA-F]{6}$/))

/* ------------- run-level format ------------- */

export const RunFormatSchema = z.strictObject({
  bold: z.optional(z.boolean()),
  italic: z.optional(z.boolean()),
  underline: z.optional(z.boolean()),
  strike: z.optional(z.boolean()),
  color: z.optional(ColorHex),
  fontLatin: z.optional(z.string()),
  fontCJK: z.optional(z.string()),
  size: z.optional(z.number().check(z.gt(0))),
  // "baseline" is the explicit reset to normal baseline (use when a parent
  // character style declares super/sub and this run needs to opt out);
  // omitting the field inherits whatever the cascade resolves to.
  vertAlign: z.optional(z.enum(["superscript", "subscript", "baseline"])),
})

/* ------------- paragraph-level format ------------- */

export const ParagraphFormatSchema = z.strictObject({
  alignment: z.optional(z.enum(["left", "center", "right", "both"])),
  spaceBefore: z.optional(z.number()),
  spaceAfter: z.optional(z.number()),
  lineSpacing: z.optional(z.union([z.number(), z.string()])),
  lineRule: z.optional(z.enum(["auto", "exact", "atLeast"])),
  firstLineIndent: z.optional(IndentValue),
  hangingIndent: z.optional(IndentValue),
  indentLeft: z.optional(IndentValue),
  indentRight: z.optional(IndentValue),
  outlineLevel: z.optional(z.number().check(z.gte(0), z.lte(9))),
})

/* ------------- inline content ------------- */

export const InlineRunSchema = z.strictObject({
  text: z.string(),
  format: z.optional(RunFormatSchema),
})

/** Word bookmark name. ECMA-376 §17.13.7 mandates start-with-letter-or-
 * underscore + alphanumerics; Word's UI lenient set adds hyphens. Capped
 * at 40 chars to match Word's UI bookmark name limit (file format permits
 * more but going above 40 is asking for cross-tool friction). */
const AnchorNameSchema = z.string().check(
  z.refine((s) => /^[A-Za-z_][A-Za-z0-9_-]{0,39}$/.test(s), {
    error:
      'anchor name must start with a letter or underscore and contain only letters, digits, underscores, or hyphens (max 40 chars). Example: "fig-architecture", "ref_smith2024".',
  }),
)

const RefToParagraphSchema = z.strictObject({
  type: z.literal("paragraph"),
  index: z.number().check(z.gte(1)),
})

const RefToAnchorSchema = z.strictObject({
  type: z.literal("anchor"),
  name: AnchorNameSchema,
})

/** Inline cross-reference to an auto-numbered paragraph. Emits as an
 * OOXML REF field; Word resolves the visible text from the target's
 * bookmark at render time. Two locator forms:
 *
 *   `{ type: "paragraph", index: N }` — resolves against the pre-edit
 *     document state. Target paragraph must be in the source and bound
 *     to a numbering scheme (unless display === "full").
 *
 *   `{ type: "anchor", name: "..." }` — resolves against
 *     (a) anchors declared on ParagraphBlock.anchor in earlier ops of
 *     this same edits[] array, or (b) bookmarks already present in the
 *     source document. Lets refs target paragraphs created in the same
 *     apply run, which the paragraph-index form cannot do.
 *
 * The target must be bound to a numbering scheme when display is "label"
 * or "number"; "full" resolves to the target paragraph's body text and
 * works on any paragraph. See references/cross-references.md. */
export const InlineRefSchema = z.strictObject({
  refTo: z.union([RefToParagraphSchema, RefToAnchorSchema]),
  display: z.optional(z.enum(["full", "label", "number"])),
  format: z.optional(RunFormatSchema),
})

export const InlineNodeSchema = z.union([InlineRunSchema, InlineRefSchema])

/** Plain string is shorthand for a single run with no inline formatting.
 * The emitter expands strings on the fly — most paragraphs are plain text. */
export const RichTextSchema = z.union([z.string(), z.array(InlineNodeSchema)])

/* ------------- numbering ref (for list paragraphs) ------------- */

const NumberingRefSchema = z.strictObject({
  numId: NonEmptyString,
  level: z.number().check(z.gte(0), z.lte(8)),
})

/* ------------- blocks ------------- */

const ParagraphBlockSchema = z.strictObject({
  type: z.literal("paragraph"),
  text: RichTextSchema,
  styleId: z.optional(NonEmptyString),
  paraFormat: z.optional(ParagraphFormatSchema),
  runFormat: z.optional(RunFormatSchema),
  numbering: z.optional(NumberingRefSchema),
  /** Optional stable name that later InlineRefs in the same edits[] (or
   * future apply runs) can target via `refTo: { type: "anchor", name }`.
   * The engine wraps the emitted paragraph with `<w:bookmarkStart>` /
   * `<w:bookmarkEnd>` carrying this name. Collisions with existing source
   * bookmark names or with anchors declared earlier in this run fail at
   * apply time. */
  anchor: z.optional(AnchorNameSchema),
})

const ImageBlockSchema = z.strictObject({
  type: z.literal("image"),
  src: NonEmptyString,
  widthPt: z.number().check(z.gt(0)),
  heightPt: z.number().check(z.gt(0)),
  alt: z.optional(z.string()),
})

const PageBreakBlockSchema = z.strictObject({
  type: z.literal("page-break"),
})

const HorizontalRuleBlockSchema = z.strictObject({
  type: z.literal("horizontal-rule"),
})

export const BlockSchema = z.union([
  ParagraphBlockSchema,
  ImageBlockSchema,
  PageBreakBlockSchema,
  HorizontalRuleBlockSchema,
])

export const FragmentSchema = z.array(BlockSchema)

/* ------------- locators ------------- */

const ParagraphLocatorSchema = z.strictObject({
  type: z.literal("paragraph"),
  index: z.number().check(z.gte(1)),
})

const RangeLocatorSchema = z
  .strictObject({
    type: z.literal("range"),
    from: z.number().check(z.gte(1)),
    to: z.number().check(z.gte(1)),
  })
  .check(
    z.refine((loc) => loc.from <= loc.to, {
      error: "range locator: from must be <= to",
    }),
  )

const CellLocatorSchema = z.strictObject({
  type: z.literal("cell"),
  table: z.number().check(z.gte(1)),
  row: z.number().check(z.gte(1)),
  col: z.number().check(z.gte(1)),
})

const HeadingLocatorSchema = z.strictObject({
  type: z.literal("heading"),
  text: NonEmptyString,
  level: z.optional(z.number().check(z.gte(0), z.lte(9))),
})

const WholeBodyLocatorSchema = z.strictObject({
  type: z.literal("whole-body"),
})

/** Run-level locator. Targets one specific <w:r> within a paragraph — used
 * by `set-run` to replace blank/placeholder run text while preserving the
 * surrounding label runs. Pick the run by 1-based `runIndex`, or by `blank`
 * (Kth run whose text is whitespace-only and rPr carries `<w:u/>` — typical
 * form-fill placeholder). When both omitted, defaults to the first blank
 * run (`blank: 1`). */
export const RunLocatorSchema = z
  .strictObject({
    type: z.literal("run"),
    paragraph: z.number().check(z.gte(1)),
    blank: z.optional(z.number().check(z.gte(1))),
    runIndex: z.optional(z.number().check(z.gte(1))),
  })
  .check(
    z.refine((loc) => !(loc.blank !== undefined && loc.runIndex !== undefined), {
      error: "run locator: pass either `blank` or `runIndex`, not both",
    }),
  )

export const LocatorSchema = z.union([
  ParagraphLocatorSchema,
  RangeLocatorSchema,
  CellLocatorSchema,
  HeadingLocatorSchema,
  WholeBodyLocatorSchema,
])

/* ------------- edit ops ------------- */

const ReplaceOpSchema = z.strictObject({
  op: z.literal("replace"),
  at: LocatorSchema,
  with: FragmentSchema,
})

const InsertBeforeOpSchema = z.strictObject({
  op: z.literal("insert-before"),
  at: LocatorSchema,
  content: FragmentSchema,
})

const InsertAfterOpSchema = z.strictObject({
  op: z.literal("insert-after"),
  at: LocatorSchema,
  content: FragmentSchema,
})

const DeleteOpSchema = z.strictObject({
  op: z.literal("delete"),
  at: LocatorSchema,
})

const FormatOpSchema = z
  .strictObject({
    op: z.literal("format"),
    at: LocatorSchema,
    styleId: z.optional(NonEmptyString),
    runFormat: z.optional(RunFormatSchema),
    paraFormat: z.optional(ParagraphFormatSchema),
  })
  .check(
    z.refine((op) => !!(op.styleId || op.runFormat || op.paraFormat), {
      error: "format op needs at least one of: styleId, runFormat, paraFormat",
    }),
  )

const SetRunOpSchema = z.strictObject({
  op: z.literal("set-run"),
  at: RunLocatorSchema,
  with: z.string(),
  format: z.optional(RunFormatSchema),
})

export const EditOpSchema = z.union([
  ReplaceOpSchema,
  InsertBeforeOpSchema,
  InsertAfterOpSchema,
  DeleteOpSchema,
  FormatOpSchema,
  SetRunOpSchema,
])

/* ------------- top-level edit config ------------- */

export const EditConfigSchema = z.strictObject({
  source: NonEmptyString,
  output: NonEmptyString,
  edits: z.array(EditOpSchema).check(z.minLength(1)),
  trackChanges: z.optional(z.boolean()),
})

/* ------------- error formatting ------------- */

function formatPath(path: readonly PropertyKey[]): string {
  let out = ""
  for (const seg of path) {
    if (typeof seg === "number") out += `[${seg}]`
    else out += out ? `.${String(seg)}` : String(seg)
  }
  return out
}

function valueAtPath(raw: unknown, path: readonly PropertyKey[]): unknown {
  let cur: unknown = raw
  for (const seg of path) {
    if (cur == null || typeof cur !== "object") return undefined
    cur = (cur as Record<PropertyKey, unknown>)[seg]
  }
  return cur
}

/** Domain hints that augment zod's defaults for known issue shapes. Returning
 * null falls back to the generic formatter. */
function customHint(issue: z.core.$ZodIssue, pathStr: string, _raw: unknown): string | null {
  // Tag absent locator discriminator with a list of the supported kinds. The
  // raw zod message ("invalid input") doesn't tell the agent what the
  // alternatives are.
  if (
    issue.code === "invalid_union" &&
    (pathStr.endsWith(".at") || pathStr.endsWith(".target") || pathStr === "")
  ) {
    // No-op: the union message is already path-prefixed; let it through.
  }
  // Color hex regex — explain what's accepted.
  if (issue.code === "invalid_format" && pathStr.endsWith(".color")) {
    return `color must be 6 hex digits without leading "#" (e.g. "1F4E79"). Use "auto" via styleId for theme-aware colors.`
  }
  return null
}

function formatIssue(issue: z.core.$ZodIssue, _pathStr: string, raw: unknown): string {
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
    case "invalid_format":
      return `format check failed (${(issue as { format?: string }).format ?? "unknown"})`
    default:
      return issue.message
  }
}

/** When zod reports `invalid_union`, the real issues are nested per-variant.
 * Pick the variant that matched the discriminator (its issues are not just
 * `invalid_value` on a `type` field) and surface those. Recurse if the
 * matched variant itself fails another union. */
function flattenUnionIssues(issue: z.core.$ZodIssue): z.core.$ZodIssue[] {
  if (issue.code !== "invalid_union") return [issue]
  const variants = (issue as { errors?: z.core.$ZodIssue[][] }).errors ?? []
  // Score each variant by "real issue count" (excluding type-discriminator
  // mismatches). Prefer the variant with the lowest non-discriminator score
  // and at least one such issue — that's the variant the agent intended.
  let best: { issues: z.core.$ZodIssue[]; score: number } | null = null
  for (const variant of variants) {
    const realIssues = variant.filter(
      (sub) =>
        !(
          sub.code === "invalid_value" &&
          sub.path.length === 1 &&
          (sub.path[0] === "type" || sub.path[0] === "op")
        ),
    )
    if (realIssues.length === 0) continue
    const score = realIssues.length
    if (!best || score < best.score) best = { issues: realIssues, score }
  }
  if (!best) {
    // All variants failed only at the discriminator — surface those raw so
    // the agent sees which `type` / `op` values are accepted.
    const first = variants.flat()[0]
    if (first) return [first]
    return [issue]
  }
  return best.issues.flatMap((sub) => {
    const merged: z.core.$ZodIssue = {
      ...sub,
      // Merge the union's path prefix with the variant-internal path.
      path: [...issue.path, ...sub.path],
    } as z.core.$ZodIssue
    return flattenUnionIssues(merged)
  })
}

export function formatEditConfigError(error: z.core.$ZodError, raw: unknown): string {
  const lines: string[] = []
  for (const top of error.issues) {
    for (const issue of flattenUnionIssues(top)) {
      const pathStr = formatPath(issue.path)
      const hint = customHint(issue, pathStr, raw)
      const msg = hint ?? formatIssue(issue, pathStr, raw)
      lines.push(pathStr ? `${pathStr}: ${msg}` : msg)
    }
  }
  return lines.join("\n")
}

export function parseEditConfig(raw: unknown): z.infer<typeof EditConfigSchema> {
  const result = z.safeParse(EditConfigSchema, raw)
  if (!result.success) {
    throw new Error(formatEditConfigError(result.error, raw))
  }
  return result.data
}
