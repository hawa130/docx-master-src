import { readFileSync, unlinkSync, existsSync, copyFileSync } from "node:fs"
import { resolve } from "node:path"
import JSZip from "jszip"
import { DOMParser } from "@xmldom/xmldom"
import { DocxReader, serializeXml } from "../core/reader.ts"
import { importTemplateStyles } from "../core/template-import.ts"
import { StyleResolver } from "../core/style-resolver.ts"
import { DocumentParser } from "../core/document-parser.ts"
import { summarizeTable } from "../core/table-classifier.ts"
import {
  NS,
  type DocumentElement,
  type ParsedParagraph,
} from "../core/types.ts"
import {
  firstChildNS,
  getChildren,
  getChildrenNS,
  textContent,
  wAttr,
  wVal,
} from "../core/xml-utils.ts"

interface StyleConfigEntry {
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

interface NumberingConfig {
  levels: Array<{
    level: number
    format: string
    text: string
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

interface AssignmentEntry {
  para: number
  action: "keep" | "restyle" | "flag"
  style?: string
  reason?: string
}

interface BulkRule {
  fingerprint: string
  style: string
}

interface PatternRule {
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

interface TemplateImportConfig {
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

interface ApplyConfig {
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
  styles: StyleConfigEntry[]
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

interface FlagRecord {
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
interface StyleResolutionEntry {
  styleId: string
  userSpec: string
  resolved: Record<string, unknown>
}

/**
 * Pick the user-facing typographic fields from a resolved style for
 * side-by-side display. Excludes mechanical fields (id, name, basedOn,
 * fromParagraph, overrides) the user wouldn't recognize.
 */
function formatResolvedFields(fields: Record<string, unknown>): string {
  const parts: string[] = []
  for (const [k, v] of Object.entries(fields)) {
    parts.push(`${k}=${JSON.stringify(v)}`)
  }
  return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`
}

function extractDisplayFields(
  def: Record<string, unknown>,
): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  const interesting = [
    "font", "fontEastAsia", "size", "bold", "italic", "color",
    "alignment", "lineSpacing", "lineRule", "spaceBefore", "spaceAfter",
    "firstLineIndent", "hangingIndent", "outlineLevel",
  ]
  for (const k of interesting) {
    if (def[k] !== undefined) out[k] = def[k]
  }
  return out
}

interface RestyleStat {
  styleId: string
  count: number
}

async function main() {
  const args = process.argv.slice(2)
  const dryRun = args.includes("--dry-run")
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error("Usage: node scripts/apply_styles.js [--dry-run] <config.json>")
    process.exit(1)
  }
  let config: ApplyConfig
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"))
  } catch (err) {
    console.error(`Cannot read config: ${(err as Error).message}`)
    process.exit(1)
  }
  if (dryRun) (config as ApplyConfig & { dryRun?: boolean }).dryRun = true

  if (!config.source || !config.output) {
    console.error("config.source and config.output are required")
    process.exit(1)
  }
  const source = resolve(config.source)
  const output = resolve(config.output)
  if (!config.dryRun && source === output) {
    console.error("output must differ from source")
    process.exit(1)
  }
  if (!existsSync(source)) {
    console.error(`source not found: ${source}`)
    process.exit(1)
  }
  if (!Array.isArray(config.styles) || config.styles.length === 0) {
    console.error("config.styles must be a non-empty array")
    process.exit(1)
  }

  try {
    await applyStyles(source, output, config)
  } catch (err) {
    if (existsSync(output)) {
      try {
        unlinkSync(output)
      } catch {}
    }
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

async function applyStyles(source: string, output: string, config: ApplyConfig) {
  // 1. In dry-run: read source directly. Otherwise copy first, modify the copy.
  if (!config.dryRun) {
    copyFileSync(source, output)
  }

  // 2. Open the COPY as our working zip
  const reader = await DocxReader.open(config.dryRun ? source : output)
  const stylesDoc =
    (await reader.readXml("word/styles.xml")) ??
    blankStylesDoc()
  const numberingDoc =
    (await reader.readXml("word/numbering.xml")) ?? blankNumberingDoc()
  const documentDoc = await reader.readXml("word/document.xml")
  if (!documentDoc) throw new Error("word/document.xml not found")
  const themeDoc = await reader.readXml("word/theme/theme1.xml")

  // 3. Resolve original styles (used for paragraph indexing & assignments)
  const resolver = new StyleResolver(stylesDoc, themeDoc)
  const parser = new DocumentParser(documentDoc, resolver, numberingDoc)
  const parsed = parser.parse()
  // parsed.paragraphs already has fingerprints? No — fingerprinter is run in load.ts only.
  // Run fingerprinter here:
  const { Fingerprinter } = await import("../core/fingerprint.ts")
  new Fingerprinter().assign(parsed.paragraphs)

  // 3b. Import template styles (if configured). This injects styles directly
  // into stylesDoc — they participate in the final styles.xml without
  // going through the user-declared styles[] / fromParagraph path. Source
  // styles[] entries with the same ID can still override a template-imported
  // style (user-declared takes precedence so the template is the baseline,
  // not a hard ceiling).
  let templateImport: { imported: string[]; numIdRemap: Map<string, string>; pulledAncestors: string[] } | null = null
  if (config.template) {
    const tplCfg = config.template
    if (!Array.isArray(tplCfg.styles) || tplCfg.styles.length === 0) {
      throw new Error(
        "config.template.styles must be a non-empty array of styleIds to import",
      )
    }
    const tplPath = resolve(tplCfg.source)
    if (!existsSync(tplPath)) {
      throw new Error(`template not found: ${tplPath}`)
    }
    templateImport = await importTemplateStyles(
      tplPath,
      tplCfg.styles,
      stylesDoc,
      numberingDoc,
      { importNumbering: tplCfg.importNumbering },
    )
    // If we imported numbering, the source needs the numbering content type
    // and rels — same plumbing as our own injectNumbering branch handles
    // later. Mark this for later.
  }

  // 4. Resolve fromParagraph references, then merge in parsed requirements,
  // then inject styles into styles.xml.
  //
  // Note: `requirements` is annotation-only. The script does NOT parse the
  // user's natural-language spec — that's the agent's job (the agent is an
  // LLM, ideal for handling negation, synonyms, hierarchical references,
  // sentence structure, ambiguity). The script just records the user's raw
  // text and the agent-resolved structured fields side-by-side in the
  // change report so any reader can verify by eye that the resolution
  // matches intent. A regex parser would give false confidence (it can't
  // tell "不要加粗" from "加粗") and tempt the agent to be lazy.
  if (config.requirements) {
    const declared = new Set(config.styles.map((s) => s.id))
    for (const styleId of Object.keys(config.requirements)) {
      if (!declared.has(styleId)) {
        throw new Error(
          `requirements: style "${styleId}" is not declared in styles[].\n` +
            `  Declared: [${[...declared].join(", ")}]`,
        )
      }
    }
  }
  const styleResolutions: StyleResolutionEntry[] = []
  const resolvedStyles = config.styles.map((def) => {
    const final = resolveStyleDef(def, parsed.paragraphs)
    if (config.requirements?.[def.id]) {
      styleResolutions.push({
        styleId: def.id,
        userSpec: config.requirements[def.id]!,
        resolved: extractDisplayFields(final),
      })
    }
    return final
  })
  const injected: string[] = []
  const updated: string[] = []
  const derivedFrom = new Map<string, number>() // styleId → fromParagraph index (for report)
  for (let i = 0; i < resolvedStyles.length; i++) {
    const def = resolvedStyles[i]!
    const result = upsertStyle(stylesDoc, def)
    if (result === "created") injected.push(def.id)
    else updated.push(def.id)
    const src = config.styles[i]!
    if (src.fromParagraph !== undefined) derivedFrom.set(def.id, src.fromParagraph)
  }

  // 5. Inject numbering
  let numIdMap = new Map<string, string>() // styleId → numId
  if (config.numbering && config.numbering.levels.length > 0) {
    const declaredIds = new Set(config.styles.map((s) => s.id))
    for (const [i, lvl] of config.numbering.levels.entries()) {
      if (!declaredIds.has(lvl.styleId)) {
        throw new Error(
          `numbering.levels[${i}]: styleId "${lvl.styleId}" is not declared in styles[].\n` +
            `  Declared: [${[...declaredIds].join(", ")}]`,
        )
      }
      // sanity-check stripPrefixPatterns vs lvlText placeholder count: a
      // pattern with more %N placeholders than the level can produce will
      // never match (e.g. "%1.%2.%3" on a level-0 lvlText "%1.")
      const numPlaceholders = (lvl.text.match(/%\d/g) ?? []).length
      for (const p of lvl.stripPrefixPatterns ?? []) {
        const pn = (p.match(/%\d/g) ?? []).length
        if (pn > numPlaceholders) {
          console.error(
            `Warning: numbering.levels[${i}].stripPrefixPatterns "${p}" has ${pn} placeholders but level lvlText "${lvl.text}" has only ${numPlaceholders}. Pattern may match more than intended.`,
          )
        }
      }
    }
    const newNumId = injectNumbering(numberingDoc, config.numbering)
    // Update the relevant styles' numPr
    for (const lvl of config.numbering.levels) {
      numIdMap.set(lvl.styleId, newNumId)
      attachNumberingToStyle(stylesDoc, lvl.styleId, newNumId, lvl.level)
    }
  }

  // 6. Build action map for paragraphs. Cross-check that every styleId
  // referenced from rules actually exists in the styles array — catching
  // typos here is much friendlier than producing a broken docx that Word
  // silently treats as Normal.
  const declaredStyleIds = new Set(config.styles.map((s) => s.id))
  const checkStyleId = (id: string | undefined, where: string): void => {
    if (!id) return
    if (!declaredStyleIds.has(id)) {
      const all = [...declaredStyleIds].join(", ")
      throw new Error(
        `${where}: style "${id}" is not declared in styles[].\n` +
          `  Declared: [${all || "(none)"}]`,
      )
    }
  }
  const validIndices = new Set(parsed.paragraphs.map((p) => p.index))
  const checkParaIndex = (idx: number, where: string): void => {
    if (validIndices.has(idx)) return
    const max = parsed.paragraphs.length
    const range = max > 0
      ? `#${parsed.paragraphs[0]!.index}–#${parsed.paragraphs[max - 1]!.index}`
      : "(none)"
    throw new Error(
      `${where}: paragraph #${idx} not found. Document has ${max} indexed paragraphs (${range}). Paragraphs inside data/form tables are not indexed.`,
    )
  }
  const excludeSet = new Set(config.exclude ?? [])
  for (const idx of excludeSet) checkParaIndex(idx, `exclude`)
  const assignmentMap = new Map<number, AssignmentEntry>()
  for (const [i, a] of (config.assignments ?? []).entries()) {
    if (a.action === "restyle") checkStyleId(a.style, `assignments[${i}]`)
    checkParaIndex(a.para, `assignments[${i}].para`)
    assignmentMap.set(a.para, a)
  }
  const bulkMap = new Map<string, string>()
  const declaredFingerprints = new Set(parsed.paragraphs.map((p) => p.fingerprint))
  for (const [i, b] of (config.bulk_rules ?? []).entries()) {
    checkStyleId(b.style, `bulk_rules[${i}]`)
    if (!declaredFingerprints.has(b.fingerprint)) {
      const all = [...declaredFingerprints].sort().join(", ")
      throw new Error(
        `bulk_rules[${i}]: fingerprint "${b.fingerprint}" doesn't exist in this document.\n` +
          `  Available fingerprints: [${all}]`,
      )
    }
    bulkMap.set(b.fingerprint, b.style)
  }

  const restyleStats = new Map<string, number>()
  const flags: FlagRecord[] = []
  let manualNumberingRemoved: Map<string, number> = new Map()
  const patternMatchStats = new Map<string, number>()
  const patternStripStats = new Map<string, number>()

  // Compile pattern_rules eagerly so a bad regex fails before we touch
  // the docx, and we can reuse the compiled regex per paragraph.
  const definedStyleIds = new Set(config.styles.map((s) => s.id))
  const patternRules: CompiledPatternRule[] = []
  for (const [i, p] of (config.pattern_rules ?? []).entries()) {
    if (!definedStyleIds.has(p.style)) {
      throw new Error(
        `pattern_rules[${i}].style "${p.style}" is not declared in styles[]. Defined: [${[...definedStyleIds].join(", ")}]`,
      )
    }
    let re: RegExp
    try {
      re = new RegExp(p.regex, p.flags ?? "")
    } catch (err) {
      throw new Error(
        `pattern_rules[${i}].regex "${p.regex}" is invalid: ${(err as Error).message}`,
      )
    }
    patternRules.push({
      regex: re,
      style: p.style,
      stripMatch: p.stripMatch ?? false,
      source: p.regex,
    })
  }

  // Build a map of styleId → list of lvlText patterns (to support manual prefix stripping).
  // Each style can have multiple alternative patterns tried in order — useful when the
  // source document mixed numbering styles across sections.
  const numLvlTextByStyle = new Map<string, string[]>()
  if (config.numbering) {
    for (const lvl of config.numbering.levels) {
      const patterns = lvl.stripPrefixPatterns && lvl.stripPrefixPatterns.length > 0
        ? lvl.stripPrefixPatterns
        : [lvl.text]
      numLvlTextByStyle.set(lvl.styleId, patterns)
    }
  }

  // 7. Walk document.xml in order and apply actions to each indexed paragraph
  const samples = new Map<string, RestyleSample[]>()
  const ctx: ApplyContext = {
    excludeSet,
    assignmentMap,
    bulkMap,
    patternRules,
    patternMatchStats,
    patternStripStats,
    paragraphs: parsed.paragraphs,
    restyleStats,
    flags,
    manualNumberingRemoved,
    numLvlTextByStyle,
    samples,
    samplesPerStyleCap: 5,
    config,
  }
  applyToBody(documentDoc, ctx)

  // 8. Serialize and write
  const replacements = new Map<string, string>()
  replacements.set("word/styles.xml", serializeXml(stylesDoc))
  if (numberingDoc) replacements.set("word/numbering.xml", serializeXml(numberingDoc))
  replacements.set("word/document.xml", serializeXml(documentDoc))
  // Make sure numbering.xml is referenced from [Content_Types].xml when
  // newly created — covers our own injectNumbering and template numbering
  // migration paths.
  const numberingTouched =
    !!config.numbering ||
    (templateImport && templateImport.numIdRemap.size > 0)
  if (numberingTouched) {
    await ensureNumberingContentType(reader, replacements)
    await ensureNumberingRelationship(reader, replacements)
  }
  // 8b. Dry-run short-circuits here: skip writing the file and the
  // post-write validation. The agent gets the report and can iterate.
  let validation: { ok: boolean; error?: string } = { ok: true }
  if (!config.dryRun) {
    await reader.copyAndModify(output, replacements)
    // 9. Validate by re-parsing modified entries
    validation = await validateOutput(output, Array.from(replacements.keys()))
    if (!validation.ok) {
      if (existsSync(output)) {
        try {
          unlinkSync(output)
        } catch {}
      }
      console.error(`Validation FAILED: ${validation.error}`)
      process.exit(1)
    }
  }

  // 10. Print report
  printReport({
    injected,
    updated,
    restyleStats,
    flags,
    manualNumberingRemoved,
    patternMatchStats,
    patternStripStats,
    styleResolutions,
    derivedFrom,
    output,
    dryRun: !!config.dryRun,
    samples: ctx.samples,
    templateImport,
  })
}

interface CompiledPatternRule {
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
interface RestyleSample {
  paraIndex: number
  oldStyle: string
  newStyle: string
  textPreview: string
  via: "assignment" | "pattern" | "bulk"
  patternSource?: string
  prefixStripped?: string
  notes: string[]
}

interface ApplyContext {
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
  numLvlTextByStyle: Map<string, string[]>
  /** First N restyled paragraphs per style — surfaced in the change report. */
  samples: Map<string, RestyleSample[]>
  /** How many samples to keep per style. */
  samplesPerStyleCap: number
  config: ApplyConfig
}

/* ------------- paragraph processing ------------- */

function applyToBody(documentDoc: Document, ctx: ApplyContext) {
  const root = documentDoc.documentElement
  const body = firstChildNS(root, NS.w, "body")
  if (!body) return
  // Build a map of paragraph index → ParsedParagraph for action lookup
  const byIdx = new Map<number, ParsedParagraph>()
  for (const p of ctx.paragraphs) byIdx.set(p.index, p)

  // Walk in same order as DocumentParser to assign indices; modify in place
  let nextIdx = 1
  const traverseChildren = (parentEl: Element | Document, insideLayout: boolean) => {
    const children = getChildren(parentEl)
    for (const child of children) {
      if (child.namespaceURI !== NS.w) continue
      if (child.localName === "p") {
        const idx = nextIdx++
        const para = byIdx.get(idx)
        if (para) processOneParagraph(child, para, ctx)
      } else if (child.localName === "tbl") {
        const summary = summarizeTable(child)
        if (summary.classification === "layout") {
          for (const tr of getChildrenNS(child, NS.w, "tr")) {
            for (const tc of getChildrenNS(tr, NS.w, "tc")) {
              traverseChildren(tc, true)
            }
          }
        }
        // data/form tables: leave content untouched
      }
    }
  }
  traverseChildren(body, false)
}

function processOneParagraph(
  pEl: Element,
  para: ParsedParagraph,
  ctx: ApplyContext,
) {
  if (ctx.excludeSet.has(para.index)) return

  const a = ctx.assignmentMap.get(para.index)
  let action: "keep" | "restyle" | "flag" = "keep"
  let targetStyle: string | undefined
  let reason: string | undefined

  let matchedPattern: { rule: CompiledPatternRule; matchLen: number } | null = null
  if (a) {
    action = a.action
    targetStyle = a.style
    reason = a.reason
  } else {
    // Try pattern_rules first — text-based classification beats fingerprint
    // when the role is content-defined (figure/table captions, references).
    for (const rule of ctx.patternRules) {
      const m = para.text.match(rule.regex)
      if (m && m.index === 0) {
        action = "restyle"
        targetStyle = rule.style
        matchedPattern = { rule, matchLen: m[0].length }
        break
      }
    }
    if (!targetStyle) {
      const bulkStyle = ctx.bulkMap.get(para.fingerprint)
      if (bulkStyle) {
        action = "restyle"
        targetStyle = bulkStyle
      }
    }
  }

  if (action === "flag" && reason) {
    ctx.flags.push({ paraIndex: para.index, reason })
    return
  }
  if (action !== "restyle" || !targetStyle) return

  // apply restyle
  const oldPStyle = para.styleId
  setParagraphStyle(pEl, targetStyle)
  stripConflictingDirectFormatting(pEl, targetStyle, ctx)
  ctx.restyleStats.set(targetStyle, (ctx.restyleStats.get(targetStyle) ?? 0) + 1)

  // Record sample for the change report (cap per style to keep output bounded)
  const existingSamples = ctx.samples.get(targetStyle) ?? []
  if (existingSamples.length < ctx.samplesPerStyleCap) {
    const via: RestyleSample["via"] = a
      ? "assignment"
      : matchedPattern
        ? "pattern"
        : "bulk"
    const sample: RestyleSample = {
      paraIndex: para.index,
      oldStyle: oldPStyle,
      newStyle: targetStyle,
      textPreview: para.text.slice(0, 60) + (para.text.length > 60 ? "…" : ""),
      via,
      patternSource: matchedPattern?.rule.source,
      notes: [],
    }
    existingSamples.push(sample)
    ctx.samples.set(targetStyle, existingSamples)
  }

  const latestSample = existingSamples[existingSamples.length - 1]
  if (matchedPattern) {
    const key = matchedPattern.rule.source
    ctx.patternMatchStats.set(key, (ctx.patternMatchStats.get(key) ?? 0) + 1)
    if (matchedPattern.rule.stripMatch) {
      const removed = removeRegexPrefix(pEl, matchedPattern.rule.regex)
      if (removed) {
        ctx.patternStripStats.set(key, (ctx.patternStripStats.get(key) ?? 0) + 1)
        if (latestSample) latestSample.notes.push(`stripped pattern /${key}/`)
      }
    }
  }

  // Manual numbering text stripping. Try each configured pattern in order;
  // first match wins. Report which pattern was used so the change report can
  // call out mixed authoring (e.g. some headings stripped via "%1.%2" and
  // others via "%1." for the same style).
  const lvlPatterns = ctx.numLvlTextByStyle.get(targetStyle)
  if (lvlPatterns && lvlPatterns.length > 0) {
    for (const pat of lvlPatterns) {
      if (removeManualNumberingPrefix(pEl, pat)) {
        ctx.manualNumberingRemoved.set(
          pat,
          (ctx.manualNumberingRemoved.get(pat) ?? 0) + 1,
        )
        if (latestSample) {
          latestSample.prefixStripped = pat
          latestSample.notes.push(`stripped manual prefix "${pat}"`)
        }
        break
      }
    }
  }
}

function setParagraphStyle(pEl: Element, styleId: string) {
  const w = NS.w
  let pPr = firstChildNS(pEl, w, "pPr")
  if (!pPr) {
    pPr = pEl.ownerDocument!.createElementNS(w, "w:pPr")
    pEl.insertBefore(pPr, pEl.firstChild)
  }
  let pStyle = firstChildNS(pPr, w, "pStyle")
  if (!pStyle) {
    pStyle = pEl.ownerDocument!.createElementNS(w, "w:pStyle")
    pPr.insertBefore(pStyle, pPr.firstChild)
  }
  pStyle.setAttributeNS(w, "w:val", styleId)
}

const RPR_CONFLICT_NAMES = [
  "rFonts",
  "sz",
  "szCs",
  "b",
  "bCs",
  "i",
  "iCs",
  "color",
] as const

function stripConflictingDirectFormatting(
  pEl: Element,
  _styleId: string,
  _ctx: ApplyContext,
) {
  const w = NS.w
  const pPr = firstChildNS(pEl, w, "pPr")
  if (pPr) {
    // remove direct paragraph formatting that the style now controls
    const conflictNames = new Set(["jc", "spacing", "ind", "outlineLvl"])
    for (const c of Array.from(getChildren(pPr))) {
      if (c.namespaceURI === w && conflictNames.has(c.localName!)) {
        pPr.removeChild(c)
      }
    }
    // paragraph-mark rPr applies only to the paragraph mark itself; safe to strip wholesale
    const paraRPr = firstChildNS(pPr, w, "rPr")
    if (paraRPr) {
      stripAllConflicts(paraRPr)
      if (getChildren(paraRPr).length === 0) pPr.removeChild(paraRPr)
    }
  }

  // Run-level rPr: only strip a property when ALL runs in the paragraph carry
  // the same value for it (it's redundant direct formatting that the style can
  // safely take over). When runs disagree on a property, that disagreement is
  // intentional mixed formatting (e.g. a bold lead phrase + non-bold body, or
  // a colored numbering prefix + a bold title) and must be preserved.
  const runs = getChildrenNS(pEl, w, "r")
  if (runs.length === 0) return

  const valuesByProp = new Map<string, Set<string>>()
  for (const name of RPR_CONFLICT_NAMES) valuesByProp.set(name, new Set())
  for (const r of runs) {
    const rPr = firstChildNS(r, w, "rPr")
    for (const name of RPR_CONFLICT_NAMES) {
      const child = rPr ? firstChildNS(rPr, w, name) : null
      valuesByProp.get(name)!.add(child ? rPrChildSignature(child, name) : "<absent>")
    }
  }
  const uniformToStrip = new Set<string>()
  for (const [name, vals] of valuesByProp) {
    if (vals.size <= 1) uniformToStrip.add(name)
  }

  for (const r of runs) {
    const rPr = firstChildNS(r, w, "rPr")
    if (!rPr) continue
    for (const c of Array.from(getChildren(rPr))) {
      if (c.namespaceURI === w && uniformToStrip.has(c.localName!)) {
        rPr.removeChild(c)
      }
    }
    if (getChildren(rPr).length === 0) r.removeChild(rPr)
  }
}

function stripAllConflicts(rPr: Element) {
  const w = NS.w
  const conflictNames = new Set<string>(RPR_CONFLICT_NAMES)
  for (const c of Array.from(getChildren(rPr))) {
    if (c.namespaceURI === w && conflictNames.has(c.localName!)) {
      rPr.removeChild(c)
    }
  }
}

function rPrChildSignature(el: Element, name: string): string {
  if (name === "rFonts") {
    return ["ascii", "hAnsi", "eastAsia", "cs"]
      .map((a) => `${a}=${wAttr(el, a) ?? ""}`)
      .join("|")
  }
  // toggles (b, bCs, i, iCs): absence==off; presence==on unless val="0"/"false"
  if (name === "b" || name === "bCs" || name === "i" || name === "iCs") {
    const v = wAttr(el, "val")
    return v === "0" || v === "false" ? "off" : "on"
  }
  return wAttr(el, "val") ?? ""
}

/**
 * Strip a regex match (anchored at start) from the leading text run of a
 * paragraph. The match may span multiple runs in a single w:t — we only
 * touch the first non-empty w:t. If the matched prefix is longer than that
 * w:t's text, we strip what's there and leave the rest alone (rare in
 * practice; manual prefixes are usually a single run).
 */
function removeRegexPrefix(pEl: Element, regex: RegExp): boolean {
  const w = NS.w
  const re = regex.source.startsWith("^")
    ? regex
    : new RegExp("^" + regex.source, regex.flags)
  for (const run of getChildrenNS(pEl, w, "r")) {
    const tEl = firstChildNS(run, w, "t")
    if (!tEl) continue
    const txt = textContent(tEl)
    if (re.test(txt)) {
      const replaced = txt.replace(re, "")
      while (tEl.firstChild) tEl.removeChild(tEl.firstChild)
      tEl.appendChild(tEl.ownerDocument!.createTextNode(replaced))
      tEl.setAttribute("xml:space", "preserve")
      return true
    }
    if (txt.trim().length > 0) break
  }
  return false
}

function removeManualNumberingPrefix(pEl: Element, lvlText: string): boolean {
  const w = NS.w
  // build regex from lvlText: replace %N with a generic numeric/Chinese capture
  const pattern = lvlText
    .replace(/[.*+?^${}()|[\]\\]/g, (m) => "\\" + m)
    .replace(/%\d/g, "(?:\\d+|[一二三四五六七八九十百千]+)")
  const re = new RegExp("^\\s*" + pattern + "\\s*")

  // find first w:t descendant and modify its text if it matches
  const runs = getChildrenNS(pEl, w, "r")
  for (const run of runs) {
    const tEl = firstChildNS(run, w, "t")
    if (!tEl) continue
    const txt = textContent(tEl)
    if (re.test(txt)) {
      const replaced = txt.replace(re, "")
      // overwrite text
      while (tEl.firstChild) tEl.removeChild(tEl.firstChild)
      tEl.appendChild(tEl.ownerDocument!.createTextNode(replaced))
      // preserve spaces
      tEl.setAttribute("xml:space", "preserve")
      return true
    }
    // only check first non-empty run
    if (txt.trim().length > 0) break
  }
  return false
}

/* ------------- fromParagraph resolution ------------- */

function resolveStyleDef(
  def: StyleConfigEntry,
  paragraphs: ParsedParagraph[],
): StyleConfigEntry {
  if (def.fromParagraph === undefined) return def
  const para = paragraphs.find((p) => p.index === def.fromParagraph)
  if (!para) {
    const indices = paragraphs.map((p) => p.index)
    const minIdx = indices[0] ?? 0
    const maxIdx = indices[indices.length - 1] ?? 0
    const closest = paragraphs.reduce(
      (best, p) =>
        Math.abs(p.index - def.fromParagraph!) < Math.abs(best.index - def.fromParagraph!) ? p : best,
      paragraphs[0]!,
    )
    throw new Error(
      `style "${def.id}": fromParagraph #${def.fromParagraph} not found.\n` +
        `  Document has ${paragraphs.length} indexed paragraphs (range: #${minIdx}–#${maxIdx}).\n` +
        `  Closest valid: #${closest.index} ("${closest.text.slice(0, 40)}${closest.text.length > 40 ? "…" : ""}")\n` +
        `  Note: paragraphs inside data tables and form tables are not indexed and cannot be referenced.`,
    )
  }
  const extracted = paragraphToStyleEntry(para)
  return {
    basedOn: "Normal",
    ...extracted,
    ...(def.overrides ?? {}),
    id: def.id,
    name: def.name,
    ...(def.basedOn !== undefined ? { basedOn: def.basedOn } : {}),
  }
}

function paragraphToStyleEntry(p: ParsedParagraph): Partial<StyleConfigEntry> {
  const r = p.rPr
  const pp = p.pPr
  const out: Partial<StyleConfigEntry> = {}
  const font = r.fontAscii ?? r.fontHAnsi
  if (font) out.font = font
  if (r.fontEastAsia && r.fontEastAsia !== font) out.fontEastAsia = r.fontEastAsia
  if (r.size !== undefined) out.size = r.size / 2
  if (r.bold) out.bold = true
  if (r.italic) out.italic = true
  if (r.color && r.color !== "auto") out.color = r.color
  if (pp.alignment) out.alignment = pp.alignment as StyleConfigEntry["alignment"]
  if (pp.spaceBefore !== undefined) out.spaceBefore = pp.spaceBefore / 20
  if (pp.spaceAfter !== undefined) out.spaceAfter = pp.spaceAfter / 20
  if (pp.lineSpacing !== undefined) {
    const rule = (pp.lineRule || "auto") as "auto" | "exact" | "atLeast"
    if (rule === "auto") {
      out.lineSpacing = pp.lineSpacing / 240
    } else {
      out.lineSpacing = pp.lineSpacing / 20
      out.lineRule = rule
    }
  }
  // Preserve character-based indent semantics when the source paragraph used
  // `firstLineChars` / `hangingChars` (Word writes these for "首行缩进 N 字符").
  // Round-tripping through pt would drop the font-size auto-scale.
  if (pp.firstLineIndentChars !== undefined) {
    out.firstLineIndent = `${pp.firstLineIndentChars / 100}char`
  } else if (pp.firstLineIndent !== undefined) {
    out.firstLineIndent = `${pp.firstLineIndent / 20}pt`
  }
  if (pp.hangingIndentChars !== undefined) {
    out.hangingIndent = `${pp.hangingIndentChars / 100}char`
  } else if (pp.hangingIndent !== undefined) {
    out.hangingIndent = `${pp.hangingIndent / 20}pt`
  }
  if (pp.outlineLevel !== undefined) out.outlineLevel = pp.outlineLevel
  // intentionally omitted: pStyle (would self-reference), numId/numLevel (bound via numbering config)
  return out
}

/* ------------- styles.xml manipulation ------------- */

function upsertStyle(stylesDoc: Document, def: StyleConfigEntry): "created" | "updated" {
  const w = NS.w
  const root = stylesDoc.documentElement!
  const existing = getChildrenNS(root, w, "style").find((s) => wAttr(s, "styleId") === def.id)
  let target: Element
  let created = false
  if (existing) {
    target = existing
    // wipe basedOn, pPr, rPr — we'll rebuild them
    for (const c of Array.from(getChildren(target))) {
      if (c.namespaceURI === w && (c.localName === "pPr" || c.localName === "rPr" || c.localName === "basedOn" || c.localName === "name")) {
        target.removeChild(c)
      }
    }
  } else {
    target = stylesDoc.createElementNS(w, "w:style")
    target.setAttributeNS(w, "w:type", "paragraph")
    target.setAttributeNS(w, "w:styleId", def.id)
    root.appendChild(target)
    created = true
  }

  const nameEl = stylesDoc.createElementNS(w, "w:name")
  nameEl.setAttributeNS(w, "w:val", def.name)
  target.appendChild(nameEl)

  if (def.basedOn) {
    const bo = stylesDoc.createElementNS(w, "w:basedOn")
    bo.setAttributeNS(w, "w:val", def.basedOn)
    target.appendChild(bo)
  }

  // pPr
  const pPr = stylesDoc.createElementNS(w, "w:pPr")
  if (def.outlineLevel !== undefined) {
    const ol = stylesDoc.createElementNS(w, "w:outlineLvl")
    ol.setAttributeNS(w, "w:val", String(def.outlineLevel))
    pPr.appendChild(ol)
  }
  if (def.alignment) {
    const jc = stylesDoc.createElementNS(w, "w:jc")
    jc.setAttributeNS(w, "w:val", def.alignment)
    pPr.appendChild(jc)
  }
  if (
    def.spaceBefore !== undefined ||
    def.spaceAfter !== undefined ||
    def.lineSpacing !== undefined
  ) {
    const spacing = stylesDoc.createElementNS(w, "w:spacing")
    if (def.spaceBefore !== undefined)
      spacing.setAttributeNS(w, "w:before", String(Math.round(def.spaceBefore * 20)))
    if (def.spaceAfter !== undefined)
      spacing.setAttributeNS(w, "w:after", String(Math.round(def.spaceAfter * 20)))
    if (def.lineSpacing !== undefined) {
      const rule = def.lineRule ?? (def.lineSpacing < 10 ? "auto" : "exact")
      if (rule === "auto") {
        spacing.setAttributeNS(w, "w:line", String(Math.round(def.lineSpacing * 240)))
        spacing.setAttributeNS(w, "w:lineRule", "auto")
      } else {
        spacing.setAttributeNS(w, "w:line", String(Math.round(def.lineSpacing * 20)))
        spacing.setAttributeNS(w, "w:lineRule", rule)
      }
    }
    pPr.appendChild(spacing)
  }
  if (def.firstLineIndent != null || def.hangingIndent != null) {
    const ind = stylesDoc.createElementNS(w, "w:ind")
    if (def.firstLineIndent != null && def.firstLineIndent !== 0) {
      const r = parseIndent(def.firstLineIndent)
      if (r.kind === "char") {
        ind.setAttributeNS(w, "w:firstLineChars", String(r.value))
      } else {
        ind.setAttributeNS(w, "w:firstLine", String(r.value))
      }
    }
    if (def.hangingIndent != null && def.hangingIndent !== 0) {
      const r = parseIndent(def.hangingIndent)
      if (r.kind === "char") {
        ind.setAttributeNS(w, "w:hangingChars", String(r.value))
      } else {
        ind.setAttributeNS(w, "w:hanging", String(r.value))
      }
    }
    pPr.appendChild(ind)
  }
  if (pPr.childNodes.length > 0) target.appendChild(pPr)

  // rPr
  const rPr = stylesDoc.createElementNS(w, "w:rPr")
  if (def.font || def.fontEastAsia) {
    const rFonts = stylesDoc.createElementNS(w, "w:rFonts")
    const ascii = def.font ?? def.fontEastAsia ?? ""
    const ea = def.fontEastAsia ?? def.font ?? ""
    if (ascii) {
      rFonts.setAttributeNS(w, "w:ascii", ascii)
      rFonts.setAttributeNS(w, "w:hAnsi", ascii)
    }
    if (ea) rFonts.setAttributeNS(w, "w:eastAsia", ea)
    rPr.appendChild(rFonts)
  }
  if (def.size !== undefined) {
    const sz = stylesDoc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(Math.round(def.size * 2)))
    rPr.appendChild(sz)
    const szCs = stylesDoc.createElementNS(w, "w:szCs")
    szCs.setAttributeNS(w, "w:val", String(Math.round(def.size * 2)))
    rPr.appendChild(szCs)
  }
  if (def.bold) {
    rPr.appendChild(stylesDoc.createElementNS(w, "w:b"))
    rPr.appendChild(stylesDoc.createElementNS(w, "w:bCs"))
  }
  if (def.italic) {
    rPr.appendChild(stylesDoc.createElementNS(w, "w:i"))
    rPr.appendChild(stylesDoc.createElementNS(w, "w:iCs"))
  }
  if (def.color) {
    const color = stylesDoc.createElementNS(w, "w:color")
    color.setAttributeNS(w, "w:val", def.color)
    rPr.appendChild(color)
  }
  if (rPr.childNodes.length > 0) target.appendChild(rPr)

  return created ? "created" : "updated"
}

/**
 * Parse an indent config value into a tagged twip-or-char unit.
 *
 *   number      → pt, written as fixed-twip indent
 *   "Npt"       → pt, fixed-twip
 *   "Nchar"     → 1/100 character, written as `firstLineChars`/`hangingChars`
 *                 so Word auto-scales the indent with the run font size
 *
 * The previous implementation collapsed both to twips by hard-coding
 * 240 twips/char (12pt assumption), which silently broke "首行缩进 2 字符"
 * for any non-12pt body and disabled font-size tracking on round-trip.
 */
function parseIndent(v: string | number): { kind: "twip" | "char"; value: number } {
  if (typeof v === "number") return { kind: "twip", value: Math.round(v * 20) }
  const m = v.trim().match(/^(-?\d+(?:\.\d+)?)\s*(char|chars|pt)?$/i)
  if (!m) return { kind: "twip", value: 0 }
  const n = parseFloat(m[1]!)
  const unit = (m[2] || "").toLowerCase()
  if (unit.startsWith("char")) return { kind: "char", value: Math.round(n * 100) }
  return { kind: "twip", value: Math.round(n * 20) }
}

/* ------------- numbering.xml manipulation ------------- */

function injectNumbering(numberingDoc: Document, config: NumberingConfig): string {
  const w = NS.w
  const root = numberingDoc.documentElement!
  // pick fresh abstractNumId and numId
  const existingAbsIds = getChildrenNS(root, w, "abstractNum").map((e) =>
    parseInt(wAttr(e, "abstractNumId") || "0", 10),
  )
  const existingNumIds = getChildrenNS(root, w, "num").map((e) =>
    parseInt(wAttr(e, "numId") || "0", 10),
  )
  const nextAbs = (existingAbsIds.length ? Math.max(...existingAbsIds) : -1) + 1
  const nextNum = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1

  const abs = numberingDoc.createElementNS(w, "w:abstractNum")
  abs.setAttributeNS(w, "w:abstractNumId", String(nextAbs))
  for (const lvl of config.levels) {
    const lvlEl = numberingDoc.createElementNS(w, "w:lvl")
    lvlEl.setAttributeNS(w, "w:ilvl", String(lvl.level))
    const start = numberingDoc.createElementNS(w, "w:start")
    start.setAttributeNS(w, "w:val", String(lvl.start ?? 1))
    lvlEl.appendChild(start)
    const numFmt = numberingDoc.createElementNS(w, "w:numFmt")
    numFmt.setAttributeNS(w, "w:val", lvl.format)
    lvlEl.appendChild(numFmt)
    const lvlText = numberingDoc.createElementNS(w, "w:lvlText")
    lvlText.setAttributeNS(w, "w:val", lvl.text)
    lvlEl.appendChild(lvlText)
    const lvlJc = numberingDoc.createElementNS(w, "w:lvlJc")
    lvlJc.setAttributeNS(w, "w:val", "left")
    lvlEl.appendChild(lvlJc)
    const pStyle = numberingDoc.createElementNS(w, "w:pStyle")
    pStyle.setAttributeNS(w, "w:val", lvl.styleId)
    lvlEl.appendChild(pStyle)
    if (lvl.numRPr) {
      const rPr = buildLvlRPr(numberingDoc, lvl.numRPr)
      if (rPr) lvlEl.appendChild(rPr)
    }
    abs.appendChild(lvlEl)
  }
  // abstractNum must come before num children — insert before any existing num
  const firstNum = getChildrenNS(root, w, "num")[0]
  if (firstNum) root.insertBefore(abs, firstNum)
  else root.appendChild(abs)

  const num = numberingDoc.createElementNS(w, "w:num")
  num.setAttributeNS(w, "w:numId", String(nextNum))
  const absRef = numberingDoc.createElementNS(w, "w:abstractNumId")
  absRef.setAttributeNS(w, "w:val", String(nextAbs))
  num.appendChild(absRef)
  root.appendChild(num)

  return String(nextNum)
}

function buildLvlRPr(
  doc: Document,
  spec: NonNullable<NumberingConfig["levels"][number]["numRPr"]>,
): Element | null {
  const w = NS.w
  const rPr = doc.createElementNS(w, "w:rPr")
  if (spec.font || spec.fontEastAsia) {
    const rFonts = doc.createElementNS(w, "w:rFonts")
    if (spec.font) {
      rFonts.setAttributeNS(w, "w:ascii", spec.font)
      rFonts.setAttributeNS(w, "w:hAnsi", spec.font)
    }
    if (spec.fontEastAsia) rFonts.setAttributeNS(w, "w:eastAsia", spec.fontEastAsia)
    rPr.appendChild(rFonts)
  }
  if (spec.size !== undefined) {
    const sz = doc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(Math.round(spec.size * 2)))
    rPr.appendChild(sz)
  }
  if (spec.bold !== undefined) {
    const b = doc.createElementNS(w, "w:b")
    if (!spec.bold) b.setAttributeNS(w, "w:val", "0")
    rPr.appendChild(b)
  }
  if (spec.italic !== undefined) {
    const i = doc.createElementNS(w, "w:i")
    if (!spec.italic) i.setAttributeNS(w, "w:val", "0")
    rPr.appendChild(i)
  }
  if (spec.color) {
    const color = doc.createElementNS(w, "w:color")
    color.setAttributeNS(w, "w:val", spec.color)
    rPr.appendChild(color)
  }
  return getChildren(rPr).length > 0 ? rPr : null
}

function attachNumberingToStyle(
  stylesDoc: Document,
  styleId: string,
  numId: string,
  level: number,
) {
  const w = NS.w
  const styleEl = getChildrenNS(stylesDoc.documentElement!, w, "style").find(
    (s) => wAttr(s, "styleId") === styleId,
  )
  if (!styleEl) return
  let pPr = firstChildNS(styleEl, w, "pPr")
  if (!pPr) {
    pPr = stylesDoc.createElementNS(w, "w:pPr")
    styleEl.appendChild(pPr)
  }
  // remove existing numPr
  const existing = firstChildNS(pPr, w, "numPr")
  if (existing) pPr.removeChild(existing)

  const numPr = stylesDoc.createElementNS(w, "w:numPr")
  const ilvl = stylesDoc.createElementNS(w, "w:ilvl")
  ilvl.setAttributeNS(w, "w:val", String(level))
  numPr.appendChild(ilvl)
  const numIdEl = stylesDoc.createElementNS(w, "w:numId")
  numIdEl.setAttributeNS(w, "w:val", numId)
  numPr.appendChild(numIdEl)
  pPr.insertBefore(numPr, pPr.firstChild)
}

/* ------------- bootstrap blank docs ------------- */

function blankStylesDoc(): Document {
  const text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}"></w:styles>`
  return new DOMParser().parseFromString(text, "text/xml") as unknown as Document
}

function blankNumberingDoc(): Document {
  const text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS.w}"></w:numbering>`
  return new DOMParser().parseFromString(text, "text/xml") as unknown as Document
}

async function ensureNumberingContentType(
  reader: DocxReader,
  replacements: Map<string, string>,
): Promise<void> {
  const ctText = await reader.readText("[Content_Types].xml")
  if (!ctText) return
  if (ctText.includes("/word/numbering.xml")) return
  const insert = `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
  const updated = ctText.replace("</Types>", `${insert}</Types>`)
  replacements.set("[Content_Types].xml", updated)
}

async function ensureNumberingRelationship(
  reader: DocxReader,
  replacements: Map<string, string>,
): Promise<void> {
  const path = "word/_rels/document.xml.rels"
  const text = await reader.readText(path)
  if (!text) return
  if (text.includes('Target="numbering.xml"')) return
  // pick a fresh rId
  const ids = Array.from(text.matchAll(/Id="rId(\d+)"/g)).map((m) => parseInt(m[1]!, 10))
  const next = (ids.length ? Math.max(...ids) : 0) + 1
  const insert = `<Relationship Id="rId${next}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`
  const updated = text.replace("</Relationships>", `${insert}</Relationships>`)
  replacements.set(path, updated)
}

/* ------------- validation ------------- */

async function validateOutput(
  outputPath: string,
  modifiedEntries: string[],
): Promise<{ ok: true } | { ok: false; error: string }> {
  try {
    const buf = readFileSync(outputPath)
    const zip = await JSZip.loadAsync(buf)
    for (const entry of modifiedEntries) {
      const file = zip.file(entry)
      if (!file) continue
      const text = await file.async("string")
      let parseError: string | null = null
      const parser = new DOMParser({
        onError: (level: any, msg: any) => {
          if (level === "error" || level === "fatalError") parseError = String(msg)
        },
      } as any)
      const doc = parser.parseFromString(text, "text/xml")
      if (parseError) return { ok: false, error: `${entry}: ${parseError}` }
      if (!doc || !(doc as any).documentElement) return { ok: false, error: `${entry}: empty doc` }
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: (err as Error).message }
  }
}

/* ------------- report ------------- */

function printReport(args: {
  injected: string[]
  updated: string[]
  restyleStats: Map<string, number>
  flags: FlagRecord[]
  manualNumberingRemoved: Map<string, number>
  patternMatchStats: Map<string, number>
  patternStripStats: Map<string, number>
  styleResolutions: StyleResolutionEntry[]
  derivedFrom: Map<string, number>
  output: string
  dryRun: boolean
  samples: Map<string, RestyleSample[]>
  templateImport: { imported: string[]; numIdRemap: Map<string, string>; pulledAncestors: string[] } | null
}) {
  const lines: string[] = []
  lines.push(args.dryRun ? "=== Change Report (DRY RUN — no file written) ===" : "=== Change Report ===")
  if (args.templateImport) {
    const ti = args.templateImport
    const directly = ti.imported.filter((id) => !ti.pulledAncestors.includes(id))
    lines.push(
      `Imported from template: ${directly.length} requested + ${ti.pulledAncestors.length} basedOn ancestors`,
    )
    lines.push(`  styles: [${ti.imported.join(", ")}]`)
    if (ti.pulledAncestors.length > 0)
      lines.push(`  pulled ancestors: [${ti.pulledAncestors.join(", ")}]`)
    if (ti.numIdRemap.size > 0) {
      const remaps = [...ti.numIdRemap.entries()].map(([o, n]) => `${o}→${n}`).join(", ")
      lines.push(`  numIds migrated: ${remaps}`)
    }
    lines.push("")
  }
  const annotate = (id: string) => {
    const src = args.derivedFrom.get(id)
    return src !== undefined ? `${id} (from #${src})` : id
  }
  lines.push(
    `Styles injected: ${args.injected.length} (${args.injected.map(annotate).join(", ")})`,
  )
  lines.push(
    `Styles updated:  ${args.updated.length} (${args.updated.map(annotate).join(", ")})`,
  )
  lines.push("")
  let totalRestyled = 0
  for (const [, c] of args.restyleStats) totalRestyled += c
  lines.push(`Paragraphs restyled: ${totalRestyled}`)
  for (const [styleId, count] of args.restyleStats) {
    lines.push(`  ${styleId}: ${count} paragraphs`)
  }
  lines.push("")
  if (args.manualNumberingRemoved.size > 0) {
    lines.push("Manual numbering converted:")
    for (const [pat, count] of args.manualNumberingRemoved) {
      lines.push(`  Prefix removed: "${pat}" (${count})`)
    }
    lines.push("")
  }
  if (args.patternMatchStats.size > 0) {
    lines.push("Pattern rules matched:")
    for (const [src, count] of args.patternMatchStats) {
      const stripped = args.patternStripStats.get(src) ?? 0
      const stripNote = stripped > 0 ? ` (stripped match in ${stripped})` : ""
      lines.push(`  /${src}/: ${count} paragraphs${stripNote}`)
    }
    lines.push("")
  }
  if (args.flags.length > 0) {
    lines.push(`Flagged (not modified): ${args.flags.length}`)
    for (const f of args.flags) {
      lines.push(`  #${f.paraIndex}: ${f.reason}`)
    }
    lines.push("")
  }
  if (args.styleResolutions.length > 0) {
    lines.push("=== Style Resolution (verify by reading) ===")
    lines.push("  The script does not parse natural language. Compare the user's")
    lines.push("  spec against the agent-resolved fields below — any mismatch")
    lines.push("  means the agent's translation needs adjustment.")
    lines.push("")
    for (const r of args.styleResolutions) {
      lines.push(`  ${r.styleId}`)
      lines.push(`    User specified: "${r.userSpec}"`)
      lines.push(`    Agent resolved: ${formatResolvedFields(r.resolved)}`)
    }
    lines.push("")
  }
  if (args.samples.size > 0) {
    lines.push("=== Sample Affected Paragraphs (first per style) ===")
    for (const [styleId, samples] of args.samples) {
      lines.push(`  ${styleId}:`)
      for (const s of samples) {
        const notes = s.notes.length > 0 ? `  [${s.notes.join("; ")}]` : ""
        lines.push(
          `    #${s.paraIndex} via=${s.via}${s.patternSource ? ` /${s.patternSource}/` : ""}: "${s.textPreview}"${notes}`,
        )
      }
    }
    lines.push("")
  }
  if (args.dryRun) {
    lines.push("Dry run — no file written, no validation performed.")
    lines.push("Re-run without --dry-run to commit changes.")
  } else {
    lines.push("Validation: PASS")
    lines.push(`Output: ${args.output}`)
  }
  console.log(lines.join("\n"))
}

main()
