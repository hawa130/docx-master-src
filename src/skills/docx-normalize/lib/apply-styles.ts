import { unlinkSync, existsSync, copyFileSync, mkdirSync } from "node:fs"
import { dirname, resolve } from "node:path"
import { DocxReader, serializeXml } from "@core/reader.ts"
import { importTemplateStyles, type ImportResult } from "@core/template-import.ts"
import { StyleResolver } from "@core/style-resolver.ts"
import { DocumentParser } from "@core/document-parser.ts"
import { Fingerprinter } from "@core/fingerprint.ts"
import { NS } from "@core/types.ts"
import { getChildrenNS, wAttr } from "@core/xml-utils.ts"
import {
  blankNumberingDoc,
  blankStylesDoc,
  ensureNumberingContentType,
  ensureNumberingRelationship,
  validateOutput,
} from "./docx-plumbing.ts"
import { applyToBody } from "./para-mutation.ts"
import {
  attachNumberingToStyle,
  injectNumbering,
} from "./numbering-mutation.ts"
import { extractDisplayFields, printReport } from "./report.ts"
import {
  reorderAgentTouchedStylesFirst,
  resolveStyleDef,
  upsertStyle,
} from "./style-mutation.ts"
import type {
  ApplyConfig,
  ApplyContext,
  AssignmentEntry,
  CompiledPatternRule,
  FlagRecord,
  RestyleSample,
  StyleResolutionEntry,
} from "./types.ts"

// Re-export the public config type so direct callers (and the narrow CLIs)
// can keep importing it from this module.
export type { ApplyConfig } from "./types.ts"

export async function applyStyles(source: string, output: string, config: ApplyConfig) {
  // 0. Default styles[] to [] when omitted. Pure template-import and
  // numbering-only configs don't need to declare any styles; CLIs that
  // require non-empty styles enforce that themselves before calling.
  config.styles ??= []

  // 1. Dry-run reads the source directly; otherwise copy first and modify
  // the copy so the original stays untouched on validation failure.
  if (!config.dryRun) {
    mkdirSync(dirname(output), { recursive: true })
    copyFileSync(source, output)
  }

  // 2. Open whichever path we'll be reading from (source or output copy).
  const reader = await DocxReader.open(config.dryRun ? source : output)
  const stylesDoc =
    (await reader.readXml("word/styles.xml")) ??
    blankStylesDoc()
  const numberingDoc =
    (await reader.readXml("word/numbering.xml")) ?? blankNumberingDoc()
  const documentDoc = await reader.readXml("word/document.xml")
  if (!documentDoc) throw new Error("word/document.xml not found")
  const themeDoc = await reader.readXml("word/theme/theme1.xml")

  // 3. Resolve original styles (used for paragraph indexing & assignments).
  // Pre-expand themed font attrs in stylesDoc so the cascade reports honest
  // values and the output XML doesn't carry a docDefaults theme reference
  // that would silently override agent-injected literal fonts at render
  // time. See StyleResolver.expandThemedFontsInStyles for the OOXML rule.
  const resolver = new StyleResolver(stylesDoc, themeDoc)
  resolver.expandThemedFontsInStyles(stylesDoc)
  const parser = new DocumentParser(documentDoc, resolver, numberingDoc)
  const parsed = parser.parse()
  // DocumentParser doesn't fingerprint — that's a fingerprinter pass run by
  // load.ts (which we don't use here, since we want the in-memory mutated
  // documentDoc, not a fresh re-open via loadDocx). Run it ourselves.
  const fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver)
  // Build hash → letter map so bulk_rules can reference fingerprints by
  // either the in-session letter (A, B, ...) or the content-derived hash
  // (stable across runs / edits — survives doc changes that shuffle
  // frequency-rank).
  const hashToLetter = new Map<string, string>()
  for (const s of fpResult.summary) {
    hashToLetter.set(s.hash, s.label)
  }

  // 3b. Import template styles (if configured). This injects styles directly
  // into stylesDoc — they participate in the final styles.xml without
  // going through the user-declared styles[] / fromParagraph path. Source
  // styles[] entries with the same ID can still override a template-imported
  // style (user-declared takes precedence so the template is the baseline,
  // not a hard ceiling).
  let templateImport: ImportResult | null = null
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
    // If template numbering was imported, ensureNumberingContentType /
    // ensureNumberingRelationship below picks it up via numIdRemap.size > 0.
  }

  // 4. Resolve fromParagraph references and inject styles into styles.xml.
  // `requirements` is annotation-only; see ApplyConfig.requirements docs.
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
  // Record every injected style for the resolution report. Styles with a
  // `requirements` entry get the user spec rendered next to the resolved
  // fields; styles without one still get listed so the agent (and any
  // reviewer) sees what actually got injected — easy to forget a self-
  // declared style otherwise.
  const styleResolutions: StyleResolutionEntry[] = []
  const resolvedStyles = config.styles.map((def) => {
    const final = resolveStyleDef(def, parsed.paragraphs)
    styleResolutions.push({
      styleId: def.id,
      userSpec: config.requirements?.[def.id] ?? null,
      resolved: extractDisplayFields(final),
    })
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
  if (config.numbering && config.numbering.levels.length > 0) {
    const declaredIds = new Set(config.styles.map((s) => s.id))
    // Numbering levels can also bind to styles already present in the
    // document's styles.xml without redeclaring them in config.styles[].
    // This unlocks the "migrate numbering only" workflow — agent specifies
    // numbering levels that target the doc's existing Heading1/2/3 without
    // having to re-extract their definitions. Template-imported styles also
    // landed in stylesDoc by this point, so they're valid numbering targets
    // too.
    const existingStyleIds = new Set<string>()
    for (const s of getChildrenNS(stylesDoc.documentElement!, NS.w, "style")) {
      const sid = wAttr(s, "styleId")
      if (sid) existingStyleIds.add(sid)
    }
    const validNumberingTargets = new Set([...declaredIds, ...existingStyleIds])
    for (const [i, lvl] of config.numbering.levels.entries()) {
      // Required-field validation. Without this an undefined `lvlText` /
      // `numFmt` crashes downstream with "Cannot read properties of
      // undefined" instead of pointing at the offending field.
      if (typeof lvl.level !== "number") {
        throw new Error(`numbering.levels[${i}]: missing required field "level" (number, 0-8)`)
      }
      if (typeof lvl.numFmt !== "string" || !lvl.numFmt) {
        throw new Error(
          `numbering.levels[${i}]: missing required field "numFmt" (e.g. "decimal", "chineseCounting", "bullet")`,
        )
      }
      if (typeof lvl.lvlText !== "string") {
        throw new Error(
          `numbering.levels[${i}]: missing required field "lvlText" (level text pattern, e.g. "%1.", "%1.%2", "第%1章")`,
        )
      }
      if (typeof lvl.styleId !== "string" || !lvl.styleId) {
        throw new Error(`numbering.levels[${i}]: missing required field "styleId" (paragraph style id this level binds to)`)
      }
      if (!validNumberingTargets.has(lvl.styleId)) {
        throw new Error(
          `numbering.levels[${i}]: styleId "${lvl.styleId}" doesn't exist.\n` +
            `  Declared in styles[]: [${[...declaredIds].join(", ")}]\n` +
            `  Existing in styles.xml: [${[...existingStyleIds].sort().join(", ")}]`,
        )
      }
      // sanity-check stripPrefixPatterns vs lvlText placeholder count: a
      // pattern with more %N placeholders than the level can produce will
      // never match (e.g. "%1.%2.%3" on a level-0 lvlText "%1.")
      const numPlaceholders = (lvl.lvlText.match(/%\d/g) ?? []).length
      const stripPatterns = lvl.stripPrefixPatterns ?? []
      for (const p of stripPatterns) {
        const pn = (p.match(/%\d/g) ?? []).length
        if (pn > numPlaceholders) {
          console.error(
            `Warning: numbering.levels[${i}].stripPrefixPatterns "${p}" has ${pn} placeholders but lvlText "${lvl.lvlText}" has only ${numPlaceholders}. Pattern may match more than intended.`,
          )
        }
      }
      // Strip patterns are tried in array order, first match wins. If a
      // shorter pattern (fewer %N placeholders) appears before a longer one,
      // it will swallow the prefix the longer one wanted — e.g.
      // ["%1.", "%1.%2"] strips just "1." from "1.1 ..." leaving ".1 ...".
      // Catch this at config time instead of letting it write bad data.
      const placeholderCounts = stripPatterns.map(
        (p) => (p.match(/%\d/g) ?? []).length,
      )
      for (let j = 1; j < placeholderCounts.length; j++) {
        if (placeholderCounts[j]! > placeholderCounts[j - 1]!) {
          throw new Error(
            `numbering.levels[${i}].stripPrefixPatterns must be ordered by descending placeholder count — ` +
              `"${stripPatterns[j - 1]}" (${placeholderCounts[j - 1]} placeholders) before "${stripPatterns[j]}" (${placeholderCounts[j]} placeholders) ` +
              `means the shorter pattern matches first and strips a prefix the longer one wanted. ` +
              `Reorder so longer patterns come first, e.g. ["%1.%2", "%1."].`,
          )
        }
      }
    }
    const newNumId = injectNumbering(numberingDoc, config.numbering)
    for (const lvl of config.numbering.levels) {
      attachNumberingToStyle(stylesDoc, lvl.styleId, newNumId, lvl.level)
    }
  }

  // 5b. Reorder agent-touched <w:style> entries to the top of the style list.
  // The default OOXML behavior (created styles append at the end) buries the
  // agent's intentional styles below the doc's pre-existing leftovers. By
  // surfacing them first, the styles.xml DOM order matches the agent's
  // mental order in config.styles[] (and the Style Resolution block in the
  // report). docDefaults / latentStyles stay above — those aren't <w:style>
  // entries and aren't touched by this reorder.
  reorderAgentTouchedStylesFirst(stylesDoc, [
    ...(templateImport?.imported ?? []),
    ...config.styles.map((s) => s.id),
  ])

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
    // Accept either the letter label or the content hash. Resolve to the
    // letter (which is what para.fingerprint stores) for the bulkMap key.
    let letter: string | undefined
    if (declaredFingerprints.has(b.fingerprint)) {
      letter = b.fingerprint
    } else if (hashToLetter.has(b.fingerprint)) {
      letter = hashToLetter.get(b.fingerprint)!
    }
    if (!letter) {
      const letters = [...declaredFingerprints].sort().join(", ")
      const hashes = [...hashToLetter.keys()].sort().join(", ")
      throw new Error(
        `bulk_rules[${i}]: fingerprint "${b.fingerprint}" doesn't exist in this document.\n` +
          `  Available letters: [${letters}]\n` +
          `  Available hashes:  [${hashes}]`,
      )
    }
    bulkMap.set(letter, b.style)
  }

  const restyleStats = new Map<string, number>()
  const flags: FlagRecord[] = []
  const manualNumberingRemoved = new Map<string, number>()
  // Parallel tracking keyed by styleId → pattern → count so the report can
  // detect "this heading style had >1 distinct strip pattern hit" — i.e. the
  // source mixed manual numbering schemes within one role (e.g. some H2 as
  // "1.1" and some as "1." across chapters). manualNumberingRemoved alone
  // can't surface this because it merges across styles.
  const manualNumberingByStyle = new Map<string, Map<string, number>>()
  const patternMatchStats = new Map<string, number>()
  const patternStripStats = new Map<string, number>()

  // Compile pattern_rules eagerly so a bad regex fails before we touch
  // the docx, and we can reuse the compiled regex per paragraph.
  const patternRules: CompiledPatternRule[] = []
  for (const [i, p] of (config.pattern_rules ?? []).entries()) {
    if (!declaredStyleIds.has(p.style)) {
      throw new Error(
        `pattern_rules[${i}].style "${p.style}" is not declared in styles[]. Defined: [${[...declaredStyleIds].join(", ")}]`,
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
        : [lvl.lvlText]
      numLvlTextByStyle.set(lvl.styleId, patterns)
    }
  }

  // 7. Walk document.xml in order and apply actions to each indexed paragraph
  const samples = new Map<string, RestyleSample[]>()
  // Track implicit-keep paragraphs split by emptiness. Empty (whitespace-only)
  // paragraphs are usually intentional spacers — silently keeping them is the
  // right behavior. Non-empty untouched paragraphs are the agent's coverage
  // signal on the Full Standardization path: if the count is unexpected, a
  // role got missed. Splitting the report makes that signal cheap to read.
  const implicitKeepByFingerprint = new Map<
    string,
    { empty: number; nonEmpty: number; nonEmptySamples: string[] }
  >()
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
    manualNumberingByStyle,
    numLvlTextByStyle,
    samples,
    samplesPerStyleCap: 5,
    implicitKeepByFingerprint,
  }
  applyToBody(documentDoc, ctx)

  // 8. Serialize and write
  const replacements = new Map<string, string>()
  replacements.set("word/styles.xml", serializeXml(stylesDoc))
  if (numberingDoc) replacements.set("word/numbering.xml", serializeXml(numberingDoc))
  replacements.set("word/document.xml", serializeXml(documentDoc))
  // Make sure numbering.xml is referenced from [Content_Types].xml when
  // newly created — covers both injectNumbering and template numbering
  // migration paths.
  const numberingTouched =
    !!config.numbering || !!templateImport?.numIdRemap.size
  if (numberingTouched) {
    await ensureNumberingContentType(reader, replacements)
    await ensureNumberingRelationship(reader, replacements)
  }
  // 8b. Dry-run skips both write and validation; the agent iterates on the
  // report alone. Otherwise: write, then re-parse the modified entries to
  // catch malformed XML before claiming success.
  if (!config.dryRun) {
    await reader.copyAndModify(output, replacements)
    const validation = await validateOutput(output, Array.from(replacements.keys()))
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

  // 9. Print report
  printReport({
    source,
    injected,
    updated,
    restyleStats,
    flags,
    manualNumberingRemoved,
    manualNumberingByStyle,
    patternMatchStats,
    patternStripStats,
    styleResolutions,
    derivedFrom,
    output,
    dryRun: !!config.dryRun,
    samples: ctx.samples,
    implicitKeepByFingerprint: ctx.implicitKeepByFingerprint,
    templateImport,
  })
}
