import type { ParsedParagraph } from "@core/types.ts"

/* ------------- public config types ------------- */

export interface StyleConfigEntry {
  id: string
  name: string
  basedOn?: string
  font?: string
  fontEastAsia?: string
  size?: number // pt
  bold?: boolean
  italic?: boolean
  color?: string
  alignment?: "left" | "center" | "right" | "both"
  lineSpacing?: number // multiple or exact pt
  lineRule?: "auto" | "exact" | "atLeast"
  spaceBefore?: number // pt
  spaceAfter?: number // pt
  firstLineIndent?: string | number | null
  hangingIndent?: string | number | null
  outlineLevel?: number
  fromParagraph?: number // 1-based paragraph index — extract computed style from this paragraph
  overrides?: Partial<Omit<StyleConfigEntry, "id" | "name" | "fromParagraph" | "overrides">>
}

export interface NumberingConfig {
  levels: Array<{
    level: number
    /** OOXML numFmt value: "decimal" | "chineseCounting" | "bullet" | "lowerRoman" | ... */
    numFmt: string
    /** OOXML lvlText pattern: e.g. "%1." / "%1.%2" / "第%1章" */
    lvlText: string
    styleId: string
    start?: number
    /**
     * Additional manual-prefix patterns to strip from paragraphs at this level.
     * Same syntax as `text` (e.g. "%1.%2", "%1.", "（%1）"). Tried in order;
     * the first regex that matches the leading text of a run is removed.
     * If omitted, the level falls back to using only `text` for stripping.
     * Useful when authors mixed numbering styles across chapters
     * (e.g. some H2 written as "1.1 …", others as "1. …").
     */
    stripPrefixPatterns?: string[]
    /**
     * rPr applied to the auto-generated number marker only (not the title text).
     * Use to keep designs where headings have e.g. blue numbering + black title.
     */
    numRPr?: {
      font?: string
      fontEastAsia?: string
      size?: number
      bold?: boolean
      italic?: boolean
      color?: string
    }
  }>
}

export interface AssignmentEntry {
  para: number
  action: "keep" | "restyle" | "flag"
  style?: string
  reason?: string
}

export interface BulkRule {
  fingerprint: string
  style: string
}

export interface PatternRule {
  /**
   * JavaScript regex source matched against paragraph plain text. Useful for
   * roles that are best identified by content prefix rather than visual
   * fingerprint: figure captions ("^图\\s*\\d"), table captions ("^表\\s*\\d"),
   * references ("^\\[\\d+\\]"), keyword lines ("^(关键词|Keywords?)\\s*[:：]").
   */
  regex: string
  flags?: string
  style: string
  /**
   * If true, strip the matched leading text from the paragraph (like
   * stripPrefixPatterns does for numbering). Useful when the matched
   * content is a label that the new style provides via numbering or
   * bookmark fields.
   */
  stripMatch?: boolean
}

export interface TemplateImportConfig {
  /** Path to the template .docx whose styles will be copied into source. */
  source: string
  /**
   * styleIds to import from the template. Their basedOn ancestors are
   * pulled in transitively. If the source already declares the same ID,
   * the template's definition wins (template is treated as authoritative,
   * which is the "stylebook" use case).
   */
  styles: string[]
  /**
   * If any imported style references a numbering scheme (numPr), copy the
   * corresponding abstractNum from the template's numbering.xml and create
   * a fresh numId in the source. Default: true. Set false to keep imported
   * styles' numPr but rely on the source's existing numIds (rare).
   */
  importNumbering?: boolean
}

export interface ApplyConfig {
  /**
   * Preview mode: run the entire pipeline in memory and print the report,
   * but skip writing the output file and skip post-write validation. Use
   * this to iterate on configs quickly without disk churn or spurious
   * artifacts. Settable via the --dry-run CLI flag too.
   */
  dryRun?: boolean
  source: string
  output: string
  /**
   * Import named styles from a template document. Useful when a thesis /
   * report template defines the canonical Heading1, BodyText, Caption etc.
   * — pull them in wholesale instead of transcribing each field by hand.
   */
  template?: TemplateImportConfig
  /** Optional. When omitted (or undefined), the engine treats it as `[]` —
   * useful for pure template-import or numbering-only operations. CLIs decide
   * whether their entry point requires a non-empty styles list. */
  styles?: StyleConfigEntry[]
  numbering?: NumberingConfig
  assignments?: AssignmentEntry[]
  bulk_rules?: BulkRule[]
  /**
   * Regex-based assignment by paragraph text. Resolution order:
   *   exclude > assignments > pattern_rules > bulk_rules > implicit-keep.
   * First matching pattern wins (within pattern_rules they're tried in order).
   */
  pattern_rules?: PatternRule[]
  /**
   * Per-style record of the user's original natural-language spec, e.g.
   *   { Heading1: "标题用黑体三号加粗居中", BodyText: "正文宋体小四…" }
   *
   * IMPORTANT: This field is annotation-only. The script does NOT parse it.
   * The agent (LLM) is responsible for translating natural language into
   * the structured `styles[i]` fields — an LLM handles negation, synonyms,
   * hierarchical references, sentence structure, and unfamiliar fonts /
   * colors that no fixed regex parser ever could.
   *
   * What the script does with this string: records it in the change report
   * next to the agent-resolved structured fields so any reader (the user,
   * a second-pass agent, or a reviewer) can verify the translation by eye.
   * That side-by-side display is the verification mechanism — not a regex
   * match, which would silently mistranslate "不要加粗" as "加粗".
   */
  requirements?: Record<string, string>
  exclude?: number[]
}

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
}
