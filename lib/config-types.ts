import type * as z from "zod/mini"
import type { ParsedParagraph } from "@lib/types.ts"
import type {
  ApplyConfigSchema,
  AssignmentSchema,
  BulkRuleSchema,
  NumberingSchema,
  PatternRuleSchema,
  StyleEntrySchema,
  TemplateSchema,
  ThemeFontsSchema,
} from "./config-schema.ts"

/* ------------- public config types -------------
 *
 * Single source of truth: zod schemas in config-schema.ts. Static types here
 * are inferred from those schemas so runtime validation and TS types stay
 * locked together — adding/removing/renaming a field on the schema flows
 * straight through to every consumer's type checks. Field-level docs
 * (semantics, examples, when-to-use) live in references/apply-styles-config.md
 * because that's what the agent actually reads at runtime; types here just
 * declare the surface.
 */

export type StyleConfigEntry = z.infer<typeof StyleEntrySchema>
export type NumberingConfig = z.infer<typeof NumberingSchema>
export type AssignmentEntry = z.infer<typeof AssignmentSchema>
export type BulkRule = z.infer<typeof BulkRuleSchema>
export type PatternRule = z.infer<typeof PatternRuleSchema>
export type ThemeFontsSpec = z.infer<typeof ThemeFontsSchema>
export type TemplateImportConfig = z.infer<typeof TemplateSchema>
export type ApplyConfig = z.infer<typeof ApplyConfigSchema>

/* ------------- internal data shapes (cross-module passes) ------------- */

export interface FlagRecord {
  paraIndex: number
  reason: string
}

/**
 * Side-by-side display of "what the user said" vs "what the agent (i.e. the
 * caller of this tool) resolved to". The script does NOT parse the user's
 * natural language — that's the agent's job. This entry is purely an
 * annotation for the change report so a human reviewer or second-pass agent
 * can spot mistranslations by reading.
 */
export interface StyleResolutionEntry {
  styleId: string
  userSpec: string | null
  resolved: Record<string, unknown>
}

export interface CompiledPatternRule {
  regex: RegExp
  style: string
  stripMatch: boolean
  source: string
}

/**
 * Per-paragraph snapshot of what a restyle did. Tracked for the dry-run
 * sample preview and for debugging "did this paragraph get touched?"
 * questions.
 */
export interface RestyleSample {
  paraIndex: number
  oldStyle: string
  newStyle: string
  textPreview: string
  via: "assignment" | "pattern" | "bulk"
  patternSource?: string
  notes: string[]
}

/** Mutable bag passed from `applyStyles` (the orchestrator) to the
 * per-paragraph walker. Everything here is either filled in upfront from
 * config or populated as paragraphs are processed. The orchestrator reads
 * the populated fields after the walk to feed the change report. */
export interface ApplyContext {
  excludeSet: Set<number>
  assignmentMap: Map<number, AssignmentEntry>
  bulkMap: Map<string, string>
  patternRules: CompiledPatternRule[]
  patternStripStats: Map<string, number>
  patternMatchStats: Map<string, number>
  paragraphs: ParsedParagraph[]
  restyleStats: Map<string, number>
  flags: FlagRecord[]
  manualNumberingRemoved: Map<string, number>
  /** styleId → (pattern → count). Used to detect mixed manual numbering
   * schemes within one heading role across the document. */
  manualNumberingByStyle: Map<string, Map<string, number>>
  numLvlTextByStyle: Map<string, string[]>
  /** First N restyled paragraphs per style — surfaced in the change report. */
  samples: Map<string, RestyleSample[]>
  /** How many samples to keep per style. */
  samplesPerStyleCap: number
  /** Paragraphs that matched no rule (no exclude, no assignment, no
   * pattern_rule, no bulk_rule). Grouped by fingerprint and split by whether
   * the paragraph has visible text — empty paragraphs are likely intentional
   * spacers, non-empty are coverage signal. */
  implicitKeepByFingerprint: Map<string, { empty: number; nonEmpty: number; nonEmptySamples: string[] }>
  /** styleId → { count, samples[] } for paragraphs assigned to a numbered
   * style whose leading text wasn't matched by any of the level's
   * stripPrefixPatterns. Surfaces sample texts (not pre-classified shapes)
   * in the change report so the agent can read the actual prefix and
   * decide: is this a missed shape that needs adding to stripPrefixPatterns,
   * or a clean heading without a manual prefix (no action)? */
  unstrippedByStyle: Map<string, { count: number; samples: string[] }>
}
