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

import {
  IndentValue,
  LengthValue,
  LineSpacingValue,
  NonEmptyString,
} from "@lib/config/zod-primitives.ts"

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
  size: z.optional(LengthValue),
  // "baseline" is the explicit reset to normal baseline (use when a parent
  // character style declares super/sub and this run needs to opt out);
  // omitting the field inherits whatever the cascade resolves to.
  vertAlign: z.optional(z.enum(["superscript", "subscript", "baseline"])),
})

/* ------------- paragraph-level format ------------- */

export const ParagraphFormatSchema = z.strictObject({
  alignment: z.optional(z.enum(["left", "center", "right", "both"])),
  spaceBefore: z.optional(LengthValue),
  spaceAfter: z.optional(LengthValue),
  lineSpacing: z.optional(LineSpacingValue),
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
 *     (a) any ParagraphBlock.anchor / EquationBlock.anchor declared
 *     anywhere in this same edits[] array (a pre-scan reserves the names
 *     before emit, so refs can address anchors that emit later — forward
 *     refs are fine), or (b) bookmarks already present in the source
 *     document. Lets refs target paragraphs created in the same apply
 *     run, which the paragraph-index form cannot do.
 *
 * The target must be bound to a numbering scheme when display is "label"
 * or "number"; "full" resolves to the target paragraph's body text and
 * works on any paragraph. See references/cross-references.md. */
export const InlineRefSchema = z.strictObject({
  refTo: z.union([RefToParagraphSchema, RefToAnchorSchema]),
  display: z.optional(z.enum(["full", "label", "number"])),
  format: z.optional(RunFormatSchema),
})

/** Inline math expression. Embedded as OMML alongside ordinary text runs in
 * the paragraph. The `math` field is LaTeX (Temml subset — see
 * references/equations.md for the supported command set and the v1 n-ary
 * operand bug). No `format` field: math runs carry their own OMML typography
 * (italic variables, upright numerals) which doesn't map cleanly to w:rPr. */
export const InlineEquationSchema = z.strictObject({
  math: NonEmptyString,
})

export const InlineNodeSchema = z.union([InlineRunSchema, InlineRefSchema, InlineEquationSchema])

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
  /** Optional stable name that InlineRefs in the same edits[] (or future
   * apply runs) can target via `refTo: { type: "anchor", name }`. The
   * engine wraps the emitted paragraph with `<w:bookmarkStart>` /
   * `<w:bookmarkEnd>` carrying this name. Names live in a flat namespace
   * with source bookmarks; collisions fail at pre-scan, before any
   * mutation. Refs can address an anchor regardless of declaration order
   * (forward refs are pre-scanned and resolved at backfill time). */
  anchor: z.optional(AnchorNameSchema),
})

const ImageBlockSchema = z.strictObject({
  type: z.literal("image"),
  src: NonEmptyString,
  width: LengthValue,
  height: LengthValue,
  alt: z.optional(z.string()),
  /** Optional paragraph-level style binding for the wrapping `<w:p>` —
   * lets a figure binding ("FigureImage" or similar) control centering /
   * keep-with-next / spaceBefore-After. Without it the image paragraph
   * has no pPr (default left alignment, no managed spacing). */
  styleId: z.optional(NonEmptyString),
  /** Per-call paragraph-format override. Same fields and override
   * precedence as on ParagraphBlock — `paraFormat` wins over `styleId`'s
   * same-named cascade output. Use for one-off tweaks (e.g. spaceAfter:0
   * to butt the figure against its caption) without minting a new style. */
  paraFormat: z.optional(ParagraphFormatSchema),
})

const PageBreakBlockSchema = z.strictObject({
  type: z.literal("page-break"),
})

const HorizontalRuleBlockSchema = z.strictObject({
  type: z.literal("horizontal-rule"),
})

/** Display equation. Becomes its own paragraph carrying `<m:oMathPara>`
 * (centered by Word's default OMML rendering, or override via styleId /
 * paraFormat). Caption + numbering follow the same caption-paragraph
 * pattern as figures and tables — see references/equations.md. */
const EquationBlockSchema = z
  .strictObject({
    type: z.literal("equation"),
    latex: z.optional(NonEmptyString),
    omml: z.optional(NonEmptyString),
    styleId: z.optional(NonEmptyString),
    paraFormat: z.optional(ParagraphFormatSchema),
    captionId: z.optional(NonEmptyString),
    subGroup: z.optional(z.enum(["start", "continue"])),
    anchor: z.optional(AnchorNameSchema),
  })
  .check(
    z.refine(
      (eq) => {
        const hasLatex = eq.latex !== undefined
        const hasOmml = eq.omml !== undefined
        if (hasLatex === hasOmml) return false // both or neither
        if (eq.subGroup !== undefined && eq.captionId === undefined) return false
        if (eq.anchor !== undefined && eq.captionId === undefined) return false
        return true
      },
      {
        error: (issue) => {
          const eq = issue.input as {
            latex?: string
            omml?: string
            subGroup?: string
            captionId?: string
            anchor?: string
          }
          const hasLatex = eq.latex !== undefined
          const hasOmml = eq.omml !== undefined
          if (hasLatex && hasOmml) {
            return "EquationBlock: latex and omml are mutually exclusive — set exactly one."
          }
          if (!hasLatex && !hasOmml) {
            return "EquationBlock: must set one of latex or omml. Use latex for standard LaTeX input; omml is the escape hatch when temml fails on the expression."
          }
          if (eq.subGroup !== undefined && eq.captionId === undefined) {
            return `EquationBlock: subGroup="${eq.subGroup}" requires captionId. Set captionId to opt into caption numbering, or remove subGroup for a standalone equation.`
          }
          if (eq.anchor !== undefined && eq.captionId === undefined) {
            return `EquationBlock: anchor="${eq.anchor}" requires captionId — without numbering the bookmark has no resolved target for REF cross-references. Set captionId, or remove anchor.`
          }
          return "EquationBlock: invalid field combination"
        },
      },
    ),
  )

/** A caption paragraph (figure title, table title, theorem statement,
 * etc.). Replaces the older pattern of `{ type: "paragraph", styleId:
 * "FigureCaption", ... }` paired with `numbering[]` binding — captions
 * config carries the numbering shape; this block carries the text. */
const CaptionBlockSchema = z.strictObject({
  type: z.literal("caption"),
  captionId: NonEmptyString,
  text: z.string(),
  anchor: z.optional(AnchorNameSchema),
})

/** Counter reset marker for a caption identifier. Emits a hidden SEQ
 * field at this position; counter sim resets accordingly. Use for
 * appendix sequences or multi-section docs where the default
 * outline-level restart isn't enough. */
const CaptionCounterResetSchema = z.strictObject({
  type: z.literal("caption-counter-reset"),
  captionId: NonEmptyString,
  newValue: z.optional(z.number().check(z.gte(0))),
})

/** Subset of blocks that may appear INSIDE a table cell. Excludes
 * TableBlock — nested tables via Block[] cell content are not supported in
 * v1 (schema would be cyclic, blocking type inference). Agents needing a
 * table inside an existing cell use a separate apply with a `cell` locator
 * + insert op. */
const CellBlockSchema = z.union([
  ParagraphBlockSchema,
  ImageBlockSchema,
  PageBreakBlockSchema,
  HorizontalRuleBlockSchema,
  EquationBlockSchema,
  CaptionBlockSchema,
  CaptionCounterResetSchema,
])

/* ------------- table block ------------- */

/** Per-edge or per-table-side border specification. */
const BorderEdgeStyle = z.enum(["single", "thick", "double", "dotted", "dashed"])

const BorderEdgeObjectSchema = z.strictObject({
  style: BorderEdgeStyle,
  /** Line size as a Length (`number` = pt, or `"Npt|Ncm|Nmm|Nin"`). Engine
   *  converts to OOXML's 1/8-pt units. Defaults: "single" → 0.5pt,
   *  "thick" → 1.5pt. */
  size: z.optional(LengthValue),
  /** Hex RGB without leading "#", or "auto" to inherit document defaults. */
  color: z.optional(z.union([ColorHex, z.literal("auto")])),
})

/** A single table border edge. String shorthand selects style with default
 * size + color "auto"; object form for full control. `"none"` suppresses
 * the edge — useful as a per-cell override. */
const BorderEdgeSchema = z.union([
  z.enum(["none", "single", "thick", "double", "dotted", "dashed"]),
  BorderEdgeObjectSchema,
])

/** Border preset for the entire table.
 *
 *   "all"        every edge thin black (Word's default-looking table)
 *   "none"       no borders anywhere
 *   "outer"      only the four outer edges
 *   "three-line" academic three-line table: thick top + thick bottom +
 *                thin line under header row. Requires `headerRows >= 1`
 *                to render the header-bottom line; with `headerRows: 0`
 *                degrades silently to "top + bottom only".
 */
const BordersPresetSchema = z.enum(["all", "none", "outer", "three-line"])

const BordersCustomSchema = z.strictObject({
  top: z.optional(BorderEdgeSchema),
  bottom: z.optional(BorderEdgeSchema),
  left: z.optional(BorderEdgeSchema),
  right: z.optional(BorderEdgeSchema),
  insideH: z.optional(BorderEdgeSchema),
  insideV: z.optional(BorderEdgeSchema),
})

export const BordersSchema = z.union([BordersPresetSchema, BordersCustomSchema])

/** Column width: `"auto"` (Word fits content) or a Length (pt by default,
 * or `"Ncm" / "Nmm" / "Nin"`). Percentage widths were considered but require
 * coordinated tblW + per-cell tcW emission with OOXML's fiftiethPercent
 * units; deferred out of v1 — use fixed widths with `layout: "fixed"` for
 * predictable sizing. */
const TableWidthSchema = z.union([z.literal("auto"), LengthValue])

const ColSpecSchema = z.strictObject({
  width: TableWidthSchema,
})

/** Per-cell properties available when content is wrapped in object form.
 * `borders` here override the table-level borders for THIS cell only.
 * Note that adjacent cells do not auto-coordinate — see tables.md. */
const TableCellObjectSchema = z.strictObject({
  content: z.union([RichTextSchema, z.array(CellBlockSchema)]),
  /** Number of columns this cell spans (gridSpan). Default 1. */
  colspan: z.optional(z.number().check(z.gte(1))),
  /** Number of rows this cell spans (vMerge restart). Engine inserts
   * continuation cells in subsequent rows at the same column position;
   * the agent must NOT declare cells at those claimed positions. */
  rowspan: z.optional(z.number().check(z.gte(1))),
  vAlign: z.optional(z.enum(["top", "center", "bottom"])),
  borders: z.optional(BordersCustomSchema),
  /** Cell background color, hex RGB. */
  shading: z.optional(ColorHex),
})

/** Cell content in four progressive forms:
 *
 *   string             plain text, single paragraph, no formatting
 *   InlineNode[]       mixed-format text or inline cross-refs, single para
 *   Block[]            multi-paragraph / images / (no nested tables — see below)
 *   { content, ... }   any of the above wrapped in an object that also
 *                      carries spans, vAlign, per-cell borders, shading
 *
 * The first three are unambiguously discriminated: string by type, the two
 * arrays by element shape (Block has `type` literal, InlineNode has only
 * `text` or `refTo`). The object form is recognized by its `content` key.
 *
 * Block[] excludes nested TableBlock in v1 to keep the schema acyclic
 * (recursion breaks TS inference). To put a table inside an existing
 * cell, use a separate apply with a `cell` locator + insert op.
 */
export const TableCellSchema = z.union([
  RichTextSchema,
  z.array(CellBlockSchema),
  TableCellObjectSchema,
])

export const TableBlockSchema = z
  .strictObject({
    type: z.literal("table"),
    rows: z.array(z.array(TableCellSchema).check(z.minLength(1))).check(z.minLength(1)),
    /** Number of top rows that repeat as the header on page breaks
     * (`<w:tblHeader/>`). Default 0 — no repeating header. Does NOT
     * auto-bold the header text; bind a styled paragraph via
     * `headerStyle`, or format each header cell explicitly. */
    headerRows: z.optional(z.number().check(z.gte(0))),
    /** styleId to apply to each cell paragraph in the header rows (top
     * `headerRows` rows). Cells that already carry an explicit styleId
     * via Block[] form win — this is a default, not a override. */
    headerStyle: z.optional(NonEmptyString),
    /** Per-column widths. Length must match the effective column count
     * (declared cells + cells claimed by ongoing rowspans, expanded by
     * colspans). When omitted, engine emits `<w:gridCol w:w="auto"/>` per
     * effective column. */
    cols: z.optional(z.array(ColSpecSchema)),
    /** Table-level borders. Default `"all"`. */
    borders: z.optional(BordersSchema),
    /** Horizontal alignment of the table on the page. Default Word
     * behavior (no `<w:jc>` emitted) is left. Academic / formal
     * documents typically center tables. */
    alignment: z.optional(z.enum(["left", "center", "right"])),
    /** Default vertical alignment for every cell whose object form
     * doesn't carry its own `vAlign`. Skill default is `"center"`
     * (academic / formal typography norm); set `"top"` for form-style
     * layouts where labels should hug the top of each cell. The
     * per-cell `vAlign` on the object form still wins per-cell. */
    vAlign: z.optional(z.enum(["top", "center", "bottom"])),
    /** Column-width interpretation. `"autofit"` (default) lets Word
     * adjust columns to content; `"fixed"` enforces declared widths even
     * if total exceeds page width (content may overflow). */
    layout: z.optional(z.enum(["fixed", "autofit"])),
  })
  .check(
    z.refine((block) => block.headerRows === undefined || block.headerRows <= block.rows.length, {
      error: "table: headerRows cannot exceed rows.length",
    }),
  )

export const BlockSchema = z.union([
  ParagraphBlockSchema,
  ImageBlockSchema,
  PageBreakBlockSchema,
  HorizontalRuleBlockSchema,
  TableBlockSchema,
  EquationBlockSchema,
  CaptionBlockSchema,
  CaptionCounterResetSchema,
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

// set-run takes a paragraph + run identification, not a paragraph-range
// locator. Schema is restricted to RunLocatorSchema (at.type === "run").
// Picking the wrong `at` form with op:"set-run" surfaces as a
// discriminator mismatch via the dispatch in EditOpSchema, so the agent
// sees "set-run requires at.type === 'run'" instead of a confusing
// "with: expected string" from a union fallback path.
const SetRunOpSchema = z.strictObject({
  op: z.literal("set-run"),
  at: RunLocatorSchema,
  with: z.string(),
  format: z.optional(RunFormatSchema),
})

/** Caption body-text edit. Targets an existing caption paragraph by
 * anchor name or by (captionId, 1-based index in body order). Replaces
 * the runs after the primary bookmarkEnd; the SEQ / STYLEREF fields and
 * bookmark itself stay intact so cross-references continue to resolve.
 *
 * Throws when target is EquationBlock (no body to edit) — agents
 * needing to change an equation use delete + re-emit instead. */
const EditCaptionTargetSchema = z.union([
  z.strictObject({ anchor: AnchorNameSchema }),
  z.strictObject({
    captionId: NonEmptyString,
    index: z.number().check(z.gte(1)),
  }),
])

const EditCaptionOpSchema = z.strictObject({
  op: z.literal("edit-caption"),
  target: EditCaptionTargetSchema,
  text: z.string(),
})

// Discriminate on `op` so a wrong-shape variant points at the right field.
// Without this, zod's plain union tries each option and reports the lowest-
// cost mismatch — which for `{ op: "set-run", at: { type: "paragraph", ... } }`
// surfaces as "with: expected string, got [...]" against ReplaceOpSchema,
// not "set-run requires at.type === 'run'". Discriminated dispatch picks the
// schema by op literal first, then validates the chosen schema's other
// fields, so all error messages reference the right op contract.
export const EditOpSchema = z.discriminatedUnion("op", [
  ReplaceOpSchema,
  InsertBeforeOpSchema,
  InsertAfterOpSchema,
  DeleteOpSchema,
  FormatOpSchema,
  SetRunOpSchema,
  EditCaptionOpSchema,
])

/* ------------- top-level edit config ------------- */

export const EditConfigSchema = z.strictObject({
  source: NonEmptyString,
  output: NonEmptyString,
  edits: z.array(EditOpSchema).check(z.minLength(1)),
  trackChanges: z.optional(z.boolean()),
})

/* ------------- error formatting ------------- */

import { formatZodError, type HintFn } from "@lib/config/zod-format.ts"

/** Domain hints that augment zod's defaults for known issue shapes. Returning
 * null falls back to the generic formatter. */
const customHint: HintFn = (issue, pathStr, raw) => {
  // Color hex regex — explain what's accepted.
  if (issue.code === "invalid_format" && pathStr.endsWith(".color")) {
    return `color must be 6 hex digits without leading "#" (e.g. "1F4E79"). Use "auto" via styleId for theme-aware colors.`
  }
  // set-run requires at.type === "run" — discriminated union surfaces the
  // type mismatch on the `at.type` path; explain the constraint inline so
  // the agent doesn't have to read the schema to fix the call site. Detect
  // the case by walking up the issue path to the parent op and reading its
  // `op` field.
  if (issue.code === "invalid_value" && pathStr.endsWith(".at.type")) {
    const opPath = issue.path.slice(0, -2)
    let cursor: unknown = raw
    for (const seg of opPath) {
      if (cursor && typeof cursor === "object") {
        cursor = (cursor as Record<string | number, unknown>)[seg as string | number]
      }
    }
    const opLiteral =
      cursor && typeof cursor === "object" && "op" in cursor
        ? (cursor as { op?: unknown }).op
        : undefined
    if (opLiteral === "set-run") {
      return `set-run requires at.type === "run" (use a RunLocator: { type: "run", paragraph, blank|runIndex }). For paragraph-range edits use op: "replace" / "format" / "insert-before" / "insert-after" instead.`
    }
  }
  return null
}

export function formatEditConfigError(error: z.core.$ZodError, raw: unknown): string {
  return formatZodError(error, raw, customHint)
}

export function parseEditConfig(raw: unknown): z.infer<typeof EditConfigSchema> {
  const result = z.safeParse(EditConfigSchema, raw)
  if (!result.success) {
    throw new Error(formatEditConfigError(result.error, raw))
  }
  return result.data
}
