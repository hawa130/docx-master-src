import { unlinkSync, existsSync, copyFileSync, mkdirSync } from "node:fs"
import { dirname, resolve } from "node:path"
import { DocxReader, serializeXml } from "@lib/reader.ts"
import { importTemplateStyles, type ImportResult } from "@lib/template-import.ts"
import { StyleResolver, applyThemeFontOverrides } from "@lib/style-resolver.ts"
import { DocumentParser } from "@lib/document-parser.ts"
import { Fingerprinter } from "@lib/fingerprint.ts"
import { NS } from "@lib/types.ts"
import { firstChildNS, getChildrenNS, wAttr } from "@lib/xml-utils.ts"
import {
  blankNumberingDoc,
  blankStylesDoc,
  ensureNumberingContentType,
  ensureNumberingRelationship,
  validateOutput,
} from "./docx-plumbing.ts"
import { applyToBody } from "./para-mutation.ts"
import { attachNumberingToStyle, injectNumbering, resolveSuff } from "./numbering-mutation.ts"
import { extractDisplayFields, printReport } from "./report.ts"
import { reorderAgentTouchedStylesFirst, resolveStyleDef, upsertStyle } from "./style-mutation.ts"
import type {
  ApplyConfig,
  ApplyContext,
  AssignmentEntry,
  CompiledPatternRule,
  FlagRecord,
  RestyleSample,
  StyleResolutionEntry,
} from "./config-types.ts"

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
  const stylesDoc = (await reader.readXml("word/styles.xml")) ?? blankStylesDoc()
  const numberingDoc = (await reader.readXml("word/numbering.xml")) ?? blankNumberingDoc()
  const documentDoc = await reader.readXml("word/document.xml")
  if (!documentDoc) throw new Error("word/document.xml not found")
  const themeDoc = await reader.readXml("word/theme/theme1.xml")

  // 3. Resolve original styles (used for paragraph indexing & assignments).
  // If the agent asked for theme-level font overrides ("design intent"
  // changes — see ThemeFontsSpec), apply them to theme1.xml first so the
  // subsequent stylesDoc expansion uses the new values. Then expand themed
  // font attrs in stylesDoc so the cascade reports honest values and the
  // output XML doesn't carry a docDefaults theme reference that would
  // silently override agent-injected literal fonts at render time. See
  // StyleResolver.expandThemedFontsInStyles for the underlying OOXML rule.
  let themeMutated = false
  if (config.theme?.fonts && themeDoc) {
    applyThemeFontOverrides(themeDoc, config.theme.fonts)
    themeMutated = true
  }
  const resolver = new StyleResolver(stylesDoc, themeDoc)
  if (config.theme?.fonts) {
    // resolver was constructed against the (possibly-updated) themeDoc, but
    // explicit setThemeFontOverrides is a belt-and-suspenders sync in case
    // theme parsing finds no fontScheme to update (rare but possible).
    resolver.setThemeFontOverrides(config.theme.fonts)
  }
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
    const tplPath = resolve(tplCfg.source)
    if (!existsSync(tplPath)) {
      throw new Error(`template not found: ${tplPath}`)
    }
    templateImport = await importTemplateStyles(tplPath, tplCfg.styles, stylesDoc, numberingDoc, {
      importNumbering: tplCfg.importNumbering,
    })
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
  // Preflight: detect <w:name> collisions with the source's existing styles.
  // Word treats `<w:name>` as the built-in style identity marker, with
  // locale-alias equivalence ("Normal" ≡ "正文" ≡ "標準"). When two
  // different styleIds claim the same identity (directly or via alias), Word
  // silently drops the new style's rPr at render — full chain (agent →
  // script → dry-run report) looks correct, only Word disagrees with no
  // error anywhere. Catch via canonical-name comparison: each name is mapped
  // to a canonical key, then equal canonical keys with different styleIds
  // are flagged. The alias table covers the common en-US ↔ zh-CN pairs
  // (which is most real Chinese-academic doc traffic for this skill);
  // unknown locales fall back to string equality.
  const canonicalNameKey = makeCanonicalNameKey()
  const sourceCanonicalToStyleId = new Map<string, string>()
  const sourceCanonicalToOriginalName = new Map<string, string>()
  for (const styleEl of getChildrenNS(stylesDoc.documentElement!, NS.w, "style")) {
    const sid = wAttr(styleEl, "styleId")
    const nameEl = firstChildNS(styleEl, NS.w, "name")
    const nm = nameEl ? wAttr(nameEl, "val") : null
    if (sid && nm) {
      const key = canonicalNameKey(nm)
      sourceCanonicalToStyleId.set(key, sid)
      sourceCanonicalToOriginalName.set(key, nm)
    }
  }
  for (const def of resolvedStyles) {
    const key = canonicalNameKey(def.name)
    const collidingId = sourceCanonicalToStyleId.get(key)
    if (collidingId && collidingId !== def.id) {
      const existingName = sourceCanonicalToOriginalName.get(key)!
      const aliasNote =
        existingName !== def.name ? ` (locale alias of existing "${existingName}")` : ""
      throw new Error(
        `name "${def.name}"${aliasNote} is already used by styleId="${collidingId}" in the source.\n` +
          `  Word treats matching names (and their locale aliases) as the same built-in identity\n` +
          `  and silently drops the new style's rPr at render time — the style would have no effect\n` +
          `  even though dry-run looks correct. Either:\n` +
          `    (a) override the existing style — set styles[].id="${collidingId}" instead of "${def.id}".\n` +
          `        upsertStyle mutates in place, preserves the doc's wiring, and avoids the collision.\n` +
          `    (b) use a non-aliasing name — typically the styleId's canonical English built-in name\n` +
          `        (e.g., "Body Text" for id="BodyText", "heading 1" for id="Heading1", "Caption" for id="Caption").`,
      )
    }
  }

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

  // 5. Inject numbering. Shape (required fields, enums, unknown-key
  // rejection, descending-placeholder ordering) is enforced upstream by the
  // zod schema; what's left here is the cross-field check that needs runtime
  // context — the level's styleId must resolve in the source's styles.xml or
  // in config.styles[]. Template-imported styles have already landed in
  // stylesDoc by this point, so they count as valid targets too. Also emit
  // the placeholder-count sanity warning, which is a heuristic (warn, not
  // throw) that doesn't fit a schema rule.
  //
  // `numbering` accepts either a single scheme or an array of schemes
  // (multi-level heading scheme + a separate single-level list-bound scheme
  // is the canonical multi-scheme case). Each scheme allocates a fresh
  // numId and binds its levels independently.
  if (config.numbering) {
    const numberingSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
    const declaredIds = new Set(config.styles.map((s) => s.id))
    const existingStyleIds = new Set<string>()
    for (const s of getChildrenNS(stylesDoc.documentElement!, NS.w, "style")) {
      const sid = wAttr(s, "styleId")
      if (sid) existingStyleIds.add(sid)
    }
    const validNumberingTargets = new Set([...declaredIds, ...existingStyleIds])
    for (const [schemeIdx, scheme] of numberingSchemes.entries()) {
      if (scheme.levels.length === 0) continue
      const path = numberingSchemes.length === 1 ? "numbering" : `numbering[${schemeIdx}]`
      for (const [i, lvl] of scheme.levels.entries()) {
        if (!validNumberingTargets.has(lvl.styleId)) {
          throw new Error(
            `${path}.levels[${i}]: styleId "${lvl.styleId}" doesn't exist.\n` +
              `  Declared in styles[]: [${[...declaredIds].join(", ")}]\n` +
              `  Existing in styles.xml: [${[...existingStyleIds].sort().join(", ")}]`,
          )
        }
        const numPlaceholders = (lvl.lvlText.match(/%\d/g) ?? []).length
        for (const p of lvl.stripPrefixPatterns ?? []) {
          const pn = (p.match(/%\d/g) ?? []).length
          if (pn > numPlaceholders) {
            console.error(
              `Warning: ${path}.levels[${i}].stripPrefixPatterns "${p}" has ${pn} placeholders but lvlText "${lvl.lvlText}" has only ${numPlaceholders}. Pattern may match more than intended.`,
            )
          }
        }
        // %N is positional (1-indexed): %1 → level 0, %2 → level 1, ...
        // A level whose lvlText omits its own counter (%(level+1)) renders
        // every item with a higher level's number. Statically detectable.
        const ownPlaceholder = `%${lvl.level + 1}`
        if (!lvl.lvlText.includes(ownPlaceholder)) {
          const referenced = lvl.lvlText.match(/%\d/g) ?? []
          const refDesc = referenced.length
            ? `only references ${referenced.map((p) => `${p} (level ${Number(p[1]) - 1})`).join(", ")}`
            : "references no counter"
          console.error(
            `Warning: ${path}.levels[${i}].lvlText "${lvl.lvlText}" does not reference its own counter ${ownPlaceholder} — ${refDesc}.\n` +
              `  All level-${lvl.level} items will display the same number, restarting only when a higher level changes.\n` +
              `  %N is positional (1-indexed): %1 → level 0, %2 → level 1, ... — so level ${lvl.level} needs ${ownPlaceholder} to render its own counter.\n` +
              `  Did you mean lvlText: "${lvl.lvlText.replace(/%\d/, ownPlaceholder)}"?`,
          )
        }
      }
      const newNumId = injectNumbering(numberingDoc, scheme)
      for (const lvl of scheme.levels) {
        attachNumberingToStyle(stylesDoc, lvl.styleId, newNumId, lvl.level)
      }
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
    const range =
      max > 0 ? `#${parsed.paragraphs[0]!.index}–#${parsed.paragraphs[max - 1]!.index}` : "(none)"
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
        { cause: err },
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
    const numberingSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
    for (const scheme of numberingSchemes) {
      for (const lvl of scheme.levels) {
        const patterns =
          lvl.stripPrefixPatterns && lvl.stripPrefixPatterns.length > 0
            ? lvl.stripPrefixPatterns
            : [lvl.lvlText]
        numLvlTextByStyle.set(lvl.styleId, patterns)
      }
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
  const unstrippedByStyle = new Map<string, { count: number; samples: string[] }>()
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
    unstrippedByStyle,
  }
  applyToBody(documentDoc, ctx)

  // 8. Serialize and write
  const replacements = new Map<string, string>()
  replacements.set("word/styles.xml", serializeXml(stylesDoc))
  if (numberingDoc) replacements.set("word/numbering.xml", serializeXml(numberingDoc))
  replacements.set("word/document.xml", serializeXml(documentDoc))
  // Only serialize theme back when we actually mutated it; otherwise keep the
  // original theme1.xml bytes untouched to avoid spurious whitespace changes.
  if (themeMutated && themeDoc) {
    replacements.set("word/theme/theme1.xml", serializeXml(themeDoc))
  }
  // Make sure numbering.xml is referenced from [Content_Types].xml when
  // newly created — covers both injectNumbering and template numbering
  // migration paths.
  const numberingTouched = !!config.numbering || !!templateImport?.numIdRemap.size
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
  // Auto-numbering bindings the agent should see surfaced — for each level
  // we pass styleId + level + lvlText so the report can show "Heading2 →
  // '%1.%2' (level 1)" without the agent needing to mentally cross-reference
  // numbering.levels[] against restyleStats.
  const numberingBindings = (() => {
    if (!config.numbering) return []
    const schemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
    return schemes.flatMap((scheme) =>
      scheme.levels.map((lvl) => {
        const { suff, effectiveLvlText } = resolveSuff(lvl.lvlText, lvl.suff)
        return {
          styleId: lvl.styleId,
          level: lvl.level,
          lvlText: effectiveLvlText,
          suff,
        }
      }),
    )
  })()

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
    unstrippedByStyle: ctx.unstrippedByStyle,
    numberingBindings,
    templateImport,
  })
}

/**
 * Build a function that maps a style display name to a locale-independent
 * canonical key. Word's built-in styles have one English name and one
 * localized name per UI language; both forms render to the same internal
 * identity, so two styleIds claiming "Normal" and "正文" respectively
 * collide just as if they both claimed "Normal".
 *
 * The pair list below covers the en-US ↔ zh-CN aliases for the built-in
 * styles agents typically inject (Normal, headings, Title, Body Text, list
 * styles, captions, headers/footers, TOC entries). Unknown names fall
 * through to themselves — string equality still catches direct collisions
 * outside the listed pairs. Add other locale pairs (zh-TW, ja, ko) when a
 * real-doc collision surfaces; this is OOXML-spec data, not natural-
 * language enumeration.
 */
function makeCanonicalNameKey(): (name: string) => string {
  // English form is the canonical key for each pair. Pairs are stored
  // both directions in a single Map for O(1) lookup either way.
  const aliasPairs: Array<readonly [string, string]> = [
    ["Normal", "正文"],
    ["heading 1", "标题 1"],
    ["heading 2", "标题 2"],
    ["heading 3", "标题 3"],
    ["heading 4", "标题 4"],
    ["heading 5", "标题 5"],
    ["heading 6", "标题 6"],
    ["heading 7", "标题 7"],
    ["heading 8", "标题 8"],
    ["heading 9", "标题 9"],
    ["Title", "标题"],
    ["Subtitle", "副标题"],
    ["Body Text", "正文文本"],
    ["Caption", "题注"],
    ["Quote", "引用"],
    ["List Bullet", "列表项目符号"],
    ["List Number", "列表编号"],
    ["Header", "页眉"],
    ["Footer", "页脚"],
    ["TOC 1", "目录 1"],
    ["TOC 2", "目录 2"],
    ["TOC 3", "目录 3"],
    ["Default Paragraph Font", "默认段落字体"],
    ["Normal Table", "普通表格"],
    ["No List", "无列表"],
  ]
  const toCanonical = new Map<string, string>()
  for (const [eng, zh] of aliasPairs) {
    toCanonical.set(eng, eng)
    toCanonical.set(zh, eng)
  }
  return (name: string) => toCanonical.get(name) ?? name
}
