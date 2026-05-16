/**
 * Zod-mini schema for the `apply` CLI's config. Replaces the hand-written
 * `typeof` / `Set.has()` validation that previously lived inline in
 * apply-styles.ts.
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
import { SECTION_SELECTOR_KEY_RE } from "@lib/apply/section-selector.ts"

/* ------------- atomic helpers ------------- */

import {
  IndentValue,
  LengthValue,
  LineSpacingValue,
  NonEmptyString,
} from "@lib/config/zod-primitives.ts"

/** Fields shared between Mode B (`styles[i].*`) direct values and the
 * `styles[i].overrides` block. Kept as a plain object so it can be spread
 * into both schemas. */
const styleFormatFields = {
  basedOn: z.optional(z.string()),
  fontLatin: z.optional(z.string()),
  fontCJK: z.optional(z.string()),
  size: z.optional(LengthValue),
  bold: z.optional(z.boolean()),
  italic: z.optional(z.boolean()),
  color: z.optional(z.string()),
  alignment: z.optional(z.enum(["left", "center", "right", "both"])),
  lineSpacing: z.optional(LineSpacingValue),
  spaceBefore: z.optional(LengthValue),
  spaceAfter: z.optional(LengthValue),
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
  size: z.optional(LengthValue),
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

/** A chapterPrefix entry. Bare string = use heading's native rendering
 * (e.g. H1 with chineseCounting numFmt renders "一"). Object form with
 * `format` forces the chapter number to render in that format — useful
 * for the common Chinese academic style where H1 displays "第一章" but
 * captions need Arabic "1.1". */
export const ChapterPrefixEntrySchema = z.union([
  NonEmptyString,
  z.strictObject({
    styleId: NonEmptyString,
    format: z.optional(CaptionFormatSchema),
  }),
])

/** Per-identifier caption configuration. Block-level `captionId`
 * references the key. `styleId` is required — every caption identifier
 * must declare what paragraph style to use; no implicit fallback. */
export const CaptionEntrySchema = z.strictObject({
  prefix: z.optional(z.string()),
  suffix: z.optional(z.string()),
  format: z.optional(CaptionFormatSchema),
  chapterPrefix: z.optional(z.array(ChapterPrefixEntrySchema)),
  chapterSeparator: z.optional(z.string()),
  bodySeparator: z.optional(z.string()),
  styleId: NonEmptyString,
  subCounter: z.optional(CaptionSubCounterSchema),
})

/** Caption identifier — free string EXCEPT the `_chap_` prefix, which
 * the engine reserves for the hidden auto-chapter SEQ counter paired
 * with chapterPrefix `format` overrides. */
const CaptionIdentifierSchema = z.string().check(
  z.refine((s) => s.length > 0 && !s.startsWith("_chap_"), {
    error:
      'caption identifier must be a non-empty string and must not start with "_chap_" (engine-reserved for hidden auto-chapter counters paired with chapterPrefix format overrides).',
  }),
)

/** Map of identifier → caption entry. Identifier is a free string;
 * conventional values "Equation" / "Figure" / "Table" / "Theorem" live
 * in ref docs, not the schema. */
export const CaptionsSchema = z.record(CaptionIdentifierSchema, CaptionEntrySchema)

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

/* ------------- page setup ------------- */

export const PAPER_SIZES = ["A4", "A3", "A5", "Letter", "Legal", "B5", "16K"] as const
const PaperSizeEnum = z.enum(PAPER_SIZES)
const PaperSizeCustom = z.strictObject({ width: LengthValue, height: LengthValue })
const PaperSizeSchema = z.union([PaperSizeEnum, PaperSizeCustom])

const MarginsSchema = z.strictObject({
  top: z.optional(LengthValue),
  bottom: z.optional(LengthValue),
  left: z.optional(LengthValue),
  right: z.optional(LengthValue),
  header: z.optional(LengthValue),
  footer: z.optional(LengthValue),
  gutter: z.optional(LengthValue),
})

/** Columns: `number` shorthand = equal-width count. Object form supports both
 *  equal-width (specify `count`) and unequal-width (specify `widths` array).
 *  `widths` and `count` are mutually exclusive — `widths.length` IS the count.
 *  When `widths` is given, OOXML `w:equalWidth="false"` is set automatically. */
const ColumnsSchema = z.union([
  z.number().check(z.gte(1)),
  z
    .strictObject({
      count: z.optional(z.number().check(z.gte(1))),
      space: z.optional(LengthValue),
      separator: z.optional(z.boolean()),
      widths: z.optional(z.array(LengthValue).check(z.minLength(1))),
      spaces: z.optional(z.array(LengthValue)),
    })
    .check(
      z.refine((c) => !(c.count !== undefined && c.widths !== undefined), {
        error:
          "columns: specify either `count` (equal-width) or `widths` (unequal-width), not both. `widths.length` is the column count.",
      }),
    )
    .check(
      z.refine((c) => c.count !== undefined || c.widths !== undefined, {
        error: "columns object: at least one of `count` or `widths` is required.",
      }),
    )
    .check(
      z.refine(
        (c) =>
          !c.spaces ||
          (c.widths !== undefined &&
            c.widths.length >= 2 &&
            c.spaces.length === c.widths.length - 1),
        {
          error:
            "columns.spaces: only valid alongside `widths` with at least 2 columns; length must equal `widths.length - 1` (one space between each pair).",
        },
      ),
    ),
])

/** Page numbering format + restart per section. `fmt` controls the numeral
 *  shape; `start` restarts the counter at that value (omit to continue). */
const PgNumTypeSchema = z.strictObject({
  fmt: z.optional(z.enum(["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"])),
  start: z.optional(z.number().check(z.gte(1))),
})

/** Fields shared between top-level (default for all sections) and per-section
 *  overrides under `sections`. */
const pageSetupFields = {
  paperSize: z.optional(PaperSizeSchema),
  orientation: z.optional(z.enum(["portrait", "landscape"])),
  margins: z.optional(MarginsSchema),
  columns: z.optional(ColumnsSchema),
  pgNumType: z.optional(PgNumTypeSchema),
}

const PageSetupSectionFieldsSchema = z.strictObject(pageSetupFields)

/** Per-section overrides keyed by section selector:
 *   "N"     — section N (1-based, matching `inspect_section`)
 *   "N-M"   — sections N through M inclusive
 *  No "all" selector — top-level fields already serve that role. */
const PageSetupSectionsSchema = z.record(z.string(), PageSetupSectionFieldsSchema).check(
  z.refine((rec) => Object.keys(rec).every((k) => SECTION_SELECTOR_KEY_RE.test(k)), {
    // zod prefixes with the field path automatically; don't repeat it here.
    error: (issue) => {
      const rec = issue.input as Record<string, unknown>
      const bad = Object.keys(rec).find((k) => !SECTION_SELECTOR_KEY_RE.test(k))
      return `key "${bad}" is invalid. Use "N" (1-based section index) or "N-M" (inclusive range).`
    },
  }),
)

export const PageSetupSchema = z.strictObject({
  ...pageSetupFields,
  sections: z.optional(PageSetupSectionsSchema),
})

/* ------------- header / footer ------------- */

import {
  BorderEdgeSchema,
  EditOpSchema,
  HorizontalRuleBlockSchema,
  ImageBlockSchema,
  ParagraphBlockSchema,
  TableBlockSchema,
} from "@lib/config/edit-config-schema.ts"

/** Block subset allowed inside a header / footer part.
 *  Excluded:
 *    - page-break  (meaningless inside HF — Word ignores it)
 *    - equation / caption / caption-counter-reset  (numbering-counter state
 *      lives in body; using these inside HF double-increments and breaks the
 *      counter sim)
 *  Table is allowed — the common header layout "left text | center text |
 *  right text" is typically a single-row 3-column borderless table. */
const HeaderFooterBlockSchema = z.union([
  ParagraphBlockSchema,
  ImageBlockSchema,
  HorizontalRuleBlockSchema,
  TableBlockSchema,
])

/** Paragraph styleIds that misbehave inside HF: heading IDs carry outline
 *  level + numbering bindings the body's chapter SEQ counter watches, so a
 *  Heading1 paragraph in a header would re-trigger chapter restarts. Title
 *  / Subtitle / BodyText carry doc-flow assumptions that look wrong in HF
 *  context. Agents wanting "this header text is bold 14pt" should use the
 *  built-in `Header` / `Footer` styleId plus a paraFormat / runFormat
 *  override, or mint a custom non-heading styleId. */
const HF_FORBIDDEN_STYLE_IDS = /^(Heading[1-9]|Title|Subtitle|BodyText)$/

/** Block `type` values disallowed at any nesting depth inside HF content,
 *  including inside table cells (where `CellBlockSchema` would otherwise
 *  permit them). caption / caption-counter-reset / equation pin counters
 *  to the body's numbering / SEQ state; emitting them inside HF would
 *  double-increment and break the counter sim. page-break is a no-op
 *  inside an HF part — Word ignores it. */
const HF_FORBIDDEN_NESTED_BLOCK_TYPES = new Set([
  "caption",
  "caption-counter-reset",
  "equation",
  "page-break",
])

/** DFS the block tree pushing into table cells via every cell-content
 *  shape that can carry Block[]. Visitor returns true to stop (violation
 *  found). Returns true when visitor signals stop, false otherwise. */
function hfWalk(blocks: unknown[], visit: (node: Record<string, unknown>) => boolean): boolean {
  const stack: unknown[] = [...blocks]
  while (stack.length > 0) {
    const node = stack.pop()
    if (!node || typeof node !== "object") continue
    const obj = node as Record<string, unknown>
    if (visit(obj)) return true
    if (obj.type === "table" && Array.isArray(obj.rows)) {
      for (const row of obj.rows as unknown[]) {
        if (!Array.isArray(row)) continue
        for (const cell of row) {
          if (Array.isArray(cell)) {
            stack.push(...cell)
          } else if (cell && typeof cell === "object" && "content" in cell) {
            const content = (cell as { content: unknown }).content
            if (Array.isArray(content)) stack.push(...content)
          }
        }
      }
    }
  }
  return false
}

const HeaderFooterContentSchema = z
  .array(HeaderFooterBlockSchema)
  .check(
    // Block-type pass: catches caption / equation / caption-counter-reset /
    // page-break smuggled in via table cells (CellBlockSchema permits them
    // even though the HF top-level union doesn't). The doc on
    // references/header-footer.md promises these are rejected at parse;
    // without this pass the rejection only surfaces at emit time.
    z.refine(
      (blocks) =>
        !hfWalk(blocks as unknown[], (obj) => {
          const t = obj.type
          return typeof t === "string" && HF_FORBIDDEN_NESTED_BLOCK_TYPES.has(t)
        }),
      {
        error:
          "header/footer cannot contain caption / equation / caption-counter-reset / " +
          "page-break blocks (allowed at any nesting depth, including inside table " +
          "cells). caption/equation/caption-counter-reset bind to body counters and " +
          "would double-increment; page-break is a no-op inside header/footer. Move " +
          "these blocks to body edits[].",
      },
    ),
  )
  .check(
    z.refine(
      (blocks) =>
        !hfWalk(blocks as unknown[], (obj) => {
          if (obj.type !== "paragraph") return false
          const id = obj.styleId
          return typeof id === "string" && HF_FORBIDDEN_STYLE_IDS.test(id)
        }),
      {
        error:
          "header/footer paragraphs cannot use heading/body-text styleIds " +
          "(Heading1..9 / Title / Subtitle / BodyText) at any nesting depth. " +
          "These styles carry outline-level and numbering bindings that " +
          "misbehave outside the body. Use the built-in `Header` / `Footer` " +
          "styleId with a paraFormat / runFormat override, or mint a custom " +
          "non-heading styleId.",
      },
    ),
  )

/** Per-surface variants. ECMA-376 sectPr supports three reference types:
 *
 *   default  — applies to every page that isn't covered by first / even
 *   first    — applies to the section's first page; needs <w:titlePg/> on
 *              the sectPr (engine auto-sets when this variant is declared)
 *   even     — applies to even-numbered pages; needs
 *              <w:evenAndOddHeaders/> in settings.xml (engine auto-sets
 *              when this variant is declared anywhere)
 *
 *  At least one variant must be declared per surface. Empty array `[]` is
 *  legal: it means "the variant exists for the trigger flags (titlePg /
 *  evenAndOdd) but renders no content" — useful for blanking the cover
 *  page's header.
 *
 *  `underline` (header) / `overline` (footer) — separator line between the
 *  HF surface and the body. Engine attaches `<w:pBdr>` to the variant's
 *  endpoint paragraph (header → last; footer → first). `true` is sugar for
 *  `"single"` 0.5pt black. Skipped silently when the variant has no
 *  paragraph at the endpoint (empty `[]`, or endpoint block is a table /
 *  image). */
const HeaderVariantsSchema = z
  .strictObject({
    default: z.optional(HeaderFooterContentSchema),
    first: z.optional(HeaderFooterContentSchema),
    even: z.optional(HeaderFooterContentSchema),
    underline: z.optional(z.union([BorderEdgeSchema, z.literal(true)])),
  })
  .check(
    z.refine((v) => v.default !== undefined || v.first !== undefined || v.even !== undefined, {
      error: "header: at least one of `default` / `first` / `even` must be declared",
    }),
  )

const FooterVariantsSchema = z
  .strictObject({
    default: z.optional(HeaderFooterContentSchema),
    first: z.optional(HeaderFooterContentSchema),
    even: z.optional(HeaderFooterContentSchema),
    overline: z.optional(z.union([BorderEdgeSchema, z.literal(true)])),
  })
  .check(
    z.refine((v) => v.default !== undefined || v.first !== undefined || v.even !== undefined, {
      error: "footer: at least one of `default` / `first` / `even` must be declared",
    }),
  )

/** Per-section override: replaces the top-level header/footer surface wholesale
 *  for the matched sections. A section that names only `header` keeps the
 *  top-level `footer` (and vice versa). At least one surface must be named —
 *  empty entries serve no purpose. */
const HeaderFooterSectionFieldsSchema = z
  .strictObject({
    header: z.optional(HeaderVariantsSchema),
    footer: z.optional(FooterVariantsSchema),
  })
  .check(
    z.refine((s) => s.header !== undefined || s.footer !== undefined, {
      error: "headerFooter.sections entry: at least one of `header` or `footer` must be declared",
    }),
  )

const HeaderFooterSectionsSchema = z.record(z.string(), HeaderFooterSectionFieldsSchema).check(
  z.refine((rec) => Object.keys(rec).every((k) => SECTION_SELECTOR_KEY_RE.test(k)), {
    error: (issue) => {
      const rec = issue.input as Record<string, unknown>
      const bad = Object.keys(rec).find((k) => !SECTION_SELECTOR_KEY_RE.test(k))
      return `key "${bad}" is invalid. Use "N" (1-based section index) or "N-M" (inclusive range).`
    },
  }),
)

export const HeaderFooterSchema = z
  .strictObject({
    header: z.optional(HeaderVariantsSchema),
    footer: z.optional(FooterVariantsSchema),
    sections: z.optional(HeaderFooterSectionsSchema),
  })
  .check(
    z.refine(
      (hf) =>
        hf.header !== undefined ||
        hf.footer !== undefined ||
        (hf.sections !== undefined && Object.keys(hf.sections).length > 0),
      {
        error:
          "headerFooter: at least one of `header`, `footer`, or non-empty `sections` must be declared",
      },
    ),
  )

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

// edits live in lib/edit-config-schema.ts; reference by import (alongside the
// header/footer Block schemas) to avoid circular engine dependencies. The
// schema is reused as-is, just embedded inside ApplyConfig as an optional
// block. When present, the engine runs edit ops after style + numbering
// install and before pattern_rules / bulk_rules / assignments — so the rules
// pass sees both pre-existing chrome paragraphs and any agent-inserted
// content uniformly.

export const ApplyConfigSchema = z
  .strictObject({
    // Omit `source` to scaffold from the bundled blank template (one empty
    // Normal paragraph on A4 portrait). Required when a `template` block is
    // declared — template-import without a host document has no resolution
    // context.
    source: z.optional(NonEmptyString),
    output: NonEmptyString,
    dryRun: z.optional(z.boolean()),
    template: z.optional(TemplateSchema),
    theme: z.optional(ThemeSchema),
    pageSetup: z.optional(PageSetupSchema),
    headerFooter: z.optional(HeaderFooterSchema),
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
  .check(
    z.refine((cfg) => !(cfg.template !== undefined && cfg.source === undefined), {
      error:
        "template: incompatible with blank-source mode (omitted `source`). Blank-source starts a clean document; template-import is a delta operation that transplants styles into an existing host with its own style cascade — combining them is conceptually inconsistent. Declare `source` when you want template-import, or drop `template` to use blank-source without the transplant.",
    }),
  )

export type ApplyConfig = z.infer<typeof ApplyConfigSchema>

/* ------------- error formatting ------------- */

import { formatZodError, valueAtPath, type HintFn } from "@lib/config/zod-format.ts"

/** Domain-specific hints that override or augment zod's default message for
 * particular issue shapes. Returning null means "use the default zod
 * message". The hints replicate the agent-friendly guidance the hand-written
 * checks used to emit. */
/** Walk up the issue path to the parent op entry and return its `op` literal,
 * or undefined when the path doesn't pass through an edit op. Shared by the
 * set-run hint with `edit-config-schema.ts` — same logic, same lookup. */
function lookupOpLiteral(raw: unknown, issuePath: readonly PropertyKey[]): unknown {
  // Path shape inside ApplyConfig: edits[N].at.type → ["edits", N, "at", "type"]
  // Truncate the last two segments (".at.type") and read .op on the parent.
  if (issuePath.length < 2) return undefined
  const opPath = issuePath.slice(0, -2)
  let cursor: unknown = raw
  for (const seg of opPath) {
    if (cursor && typeof cursor === "object") {
      cursor = (cursor as Record<PropertyKey, unknown>)[seg]
    } else {
      return undefined
    }
  }
  if (cursor && typeof cursor === "object" && "op" in cursor) {
    return (cursor as { op?: unknown }).op
  }
  return undefined
}

const customHint: HintFn = (issue, pathStr, raw) => {
  // set-run requires at.type === "run". The discriminated union dispatches
  // on `op` first, so the resulting error is an at.type mismatch — explain
  // the constraint inline instead of leaving the agent to read the schema.
  if (issue.code === "invalid_value" && pathStr.endsWith(".at.type")) {
    const opLiteral = lookupOpLiteral(raw, issue.path)
    if (opLiteral === "set-run") {
      return `set-run requires at.type === "run" (use a RunLocator: { type: "run", paragraph, blank|runIndex }). For paragraph-range edits use op: "replace" / "format" / "insert-before" / "insert-after" instead.`
    }
  }
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
