import { unlinkSync, existsSync, copyFileSync, mkdirSync } from "node:fs"
import { dirname, resolve } from "node:path"
import { DocxReader, serializeXml } from "@lib/xml/reader.ts"
import { WritableArchive } from "@lib/xml/writable-archive.ts"
import { importTemplateStyles, type ImportResult } from "@lib/apply/template-import.ts"
import { StyleResolver, applyThemeFontOverrides } from "@lib/parse/style-resolver.ts"
import { DocumentParser } from "@lib/parse/document-parser.ts"
import { Fingerprinter } from "@lib/parse/fingerprint.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import { canonicalStyleName } from "@lib/parse/builtin-styles.ts"
import { blankNumberingDoc, blankStylesDoc } from "@lib/xml/docx-plumbing.ts"
import { applyToBody } from "@lib/apply/para-mutation.ts"
import { walkIndexedParagraphs } from "@lib/edit/locator.ts"
import { analyzeVsDirect } from "@lib/shared/vs-direct.ts"
import type { ParsedParagraph } from "@lib/parse/types.ts"
import {
  attachNumberingToStyle,
  injectNumbering,
  resolveSuff,
} from "@lib/apply/numbering-mutation.ts"
import { applyListRestartPass, buildHeadingStyleIdSet } from "@lib/apply/list-restart.ts"
import { detectManualNumbering } from "@lib/parse/manual-numbering-detect.ts"
import { validateDocxFile } from "@lib/shared/docx-validate.ts"
import { extractDisplayFields, printReport } from "@lib/shared/report.ts"
import {
  extractPriorDisplayFields,
  reorderAgentTouchedStylesFirst,
  resolveStyleDef,
  upsertStyle,
} from "@lib/apply/style-mutation.ts"
import { previewEditOps, runEditOps, type EditsPreviewEntry } from "@lib/edit/edit-engine.ts"
import {
  lintPanguInEdits,
  lintPanguInHeaderFooter,
  type PanguWarning,
} from "@lib/edit/pangu-lint.ts"
import { DocxAssetRegistry } from "@lib/edit/asset-registry.ts"
import { ensureHyperlinkCharStyle } from "@lib/edit/hyperlink.ts"
import { simulateNumberingCounters, extractParagraphText } from "@lib/apply/numbering-counter.ts"
import { resolveCaptions } from "@lib/parse/caption-resolver.ts"
import {
  simulateCaptions,
  type PendingCaptionFill,
  type PendingCaptionReset,
} from "@lib/edit/caption-counter.ts"
import { BookmarkAllocator } from "@lib/edit/bookmark.ts"
import type { PendingRefBackfill } from "@lib/edit/fields/ref-field.ts"
import { standardizeCaptions } from "@lib/apply/standardize-captions.ts"
import {
  ensureHiddenChapterCounterStyle,
  injectChapterCounters,
} from "@lib/apply/inject-chapter-counters.ts"
import { ensureUpdateFieldsFlag, setEvenAndOddHeadersFlag } from "@lib/apply/settings-mutation.ts"
import { applyPageSetup, type PageSetupReport } from "@lib/apply/page-setup-mutation.ts"
import {
  applyHeaderFooter,
  applyHeaderFooterBinding,
  ensureHeaderFooterStyles,
  type HeaderFooterReport,
  type HeaderFooterBindingReport,
} from "@lib/apply/header-footer-mutation.ts"
import type {
  ApplyConfig,
  ApplyContext,
  AssignmentEntry,
  CompiledPatternRule,
  FlagRecord,
  RestyleSample,
  StyleConfigEntry,
  StyleResolutionEntry,
} from "@lib/config/config-types.ts"

/**
 * Normalize a ValidationError into a stable key for baseline comparison.
 * Strips file paths, line/column coordinates, and paragraph indices — all
 * of which differ between source and output but don't indicate a new
 * structural problem introduced by apply. Two errors that vary only in
 * position (same schema violation, different location in same part) map to
 * the same key so a pre-existing positional shift doesn't create a false
 * "new" error.
 */
function normalizeValidationError(e: { part: string; message: string }): string {
  const msg = e.message
    .replace(/\/[^\s"']+\.docx/g, "<docx>")
    .replace(/\bline \d+\b/g, "line N")
    .replace(/\bcol \d+\b/g, "col M")
    .replace(/\bparagraph #\d+\b/g, "paragraph #N")
    .replace(/\binput_\d+\.xml\b/g, "input_N.xml")
    .trim()
  return `${e.part}||${msg}`
}

/** Heuristic checks on a resolved style entry. The engine doesn't reject
 * these — they're informational signals surfaced in dry-run + the change
 * report so the agent can fix before commit. Narrow by design: only fire
 * on Mode A extraction artifacts, not on Mode B explicit declarations
 * (agent may legitimately want the CJK font's Latin glyphs for English /
 * digit text, which is what `fontLatin` controls). */
function detectStyleResolutionWarnings(def: StyleConfigEntry, final: StyleConfigEntry): string[] {
  const out: string[] = []
  // fromParagraph extraction + Latin slot holds a CJK-character value +
  // no fontCJK extracted = source paragraph almost certainly had the CJK
  // font on its ascii/hAnsi slot (common Word/POI authoring artifact);
  // extraction faithfully reproduced the slot mismatch. CJK characters in
  // the new style will fall through to docDefaults eastAsia at render.
  if (
    def.fromParagraph !== undefined &&
    typeof final.fontLatin === "string" &&
    /[一-鿿]/.test(final.fontLatin) &&
    final.fontCJK === undefined
  ) {
    out.push(
      `fromParagraph extracted fontLatin="${final.fontLatin}" with no fontCJK — ` +
        `source paragraph likely had the CJK font on its ascii/hAnsi slot only ` +
        `(common Word authoring artifact). CJK characters fall through to ` +
        `docDefaults eastAsia at render; add fontCJK to this style if you want ` +
        `consistent CJK appearance, or ignore if the Latin-only slot was intentional.`,
    )
  }
  return out
}

/** Build per-styleId target paragraph sets from the post-rules-pass doc
 *  state, then attach a vs-direct classification to each StyleResolutionEntry.
 *  Walks the live documentDoc to read each paragraph's CURRENT pStyle (after
 *  edits + rules), matches back to parsed.paragraphs by index for the
 *  direct-format snapshot captured at parse time. edits[]-inserted paragraphs
 *  whose direct format came from agent's own paraFormat/runFormat would
 *  trivially appear redundant — they're included for now (acceptable noise
 *  given the small footprint of inserts vs chrome). */
/** True iff `el` is reachable from `doc.documentElement` (not removed). Walks
 * parent links rather than relying on any cached structure; constant-time per
 * call for shallow trees and bounded for paragraphs. */
function isAttachedToDoc(el: Element, doc: Document): boolean {
  let cur: Node | null = el
  while (cur) {
    if (cur === doc.documentElement) return true
    cur = cur.parentNode
  }
  return false
}

function populateVsDirectReports(
  documentDoc: Document,
  paragraphs: readonly ParsedParagraph[],
  declaredStyles: readonly StyleConfigEntry[],
  styleResolutions: StyleResolutionEntry[],
  resolver: StyleResolver,
): void {
  const w = NS.w
  const paraByIndex = new Map<number, ParsedParagraph>()
  for (const p of paragraphs) paraByIndex.set(p.index, p)
  const defaultStyleId = resolver.getDefaultParagraphStyleId() ?? "Normal"

  // Walk post-mutation doc; read each paragraph's current pStyle.
  const targetsByStyleId = new Map<string, ParsedParagraph[]>()
  for (const indexed of walkIndexedParagraphs(documentDoc)) {
    const p = paraByIndex.get(indexed.index)
    if (!p) continue
    const pPr = firstChildNS(indexed.element, w, "pPr")
    const pStyleEl = pPr ? firstChildNS(pPr, w, "pStyle") : null
    const currentStyleId = (pStyleEl && wAttr(pStyleEl, "val")) || defaultStyleId
    let bucket = targetsByStyleId.get(currentStyleId)
    if (!bucket) {
      bucket = []
      targetsByStyleId.set(currentStyleId, bucket)
    }
    bucket.push(p)
  }

  // Attach the analysis result onto the matching StyleResolutionEntry.
  const resolutionsById = new Map<string, StyleResolutionEntry>()
  for (const r of styleResolutions) resolutionsById.set(r.styleId, r)
  for (const def of declaredStyles) {
    const entry = resolutionsById.get(def.id)
    if (!entry) continue
    const targets = targetsByStyleId.get(def.id) ?? []
    // Mode A same-source silencing: get the source paragraph's fingerprint
    // so the analyzer can skip the redundant tally for paragraphs that
    // necessarily match by construction.
    let fromFp: string | undefined
    if (def.fromParagraph !== undefined) {
      fromFp = paraByIndex.get(def.fromParagraph)?.fingerprint
    }
    entry.vsDirect = analyzeVsDirect(def, targets, fromFp)
  }
}

export async function applyStyles(source: string, output: string, config: ApplyConfig) {
  // 0. Default styles[] to [] when omitted. Pure template-import and
  // numbering-only configs don't need to declare any styles; CLIs that
  // require non-empty styles enforce that themselves before calling.
  config.styles ??= []

  // Cross-config invariant: a styleId can't be bound to both numbering[]
  // (numPr-based outline / list counter) and captions[] (SEQ-based caption
  // counter). The two mechanisms would fight at render time.
  if (config.captions && config.numbering) {
    const captionStyleIds = new Set<string>()
    for (const entry of Object.values(config.captions)) {
      captionStyleIds.add(entry.styleId)
    }
    const numSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
    for (const scheme of numSchemes) {
      for (const lvl of scheme.levels) {
        if (captionStyleIds.has(lvl.styleId)) {
          throw new Error(
            `Config invariant: styleId "${lvl.styleId}" is bound to both numbering[] (numPr) and captions[] (SEQ). ` +
              `Caption-class styleIds must not appear in numbering[]; they get their counter from the captions table instead. ` +
              `Drop the numbering[] level referencing "${lvl.styleId}".`,
          )
        }
      }
    }
  }

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
  let parsed = parser.parse()
  // DocumentParser doesn't fingerprint — that's a fingerprinter pass run by
  // load.ts (which we don't use here, since we want the in-memory mutated
  // documentDoc, not a fresh re-open via loadDocx). Run it ourselves.
  let fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver)
  // Build hash → letter map so bulk_rules can reference fingerprints by
  // either the in-session letter (A, B, ...) or the content-derived hash
  // (stable across runs / edits — survives doc changes that shuffle
  // frequency-rank).
  const hashToLetter = new Map<string, string>()
  for (const s of fpResult.summary) {
    hashToLetter.set(s.hash, s.label)
  }

  // Snapshot numIds from the **original** numbering.xml before template
  // import runs. The downstream collision check at step 5 needs to
  // distinguish "id pre-existed in user's doc" from "id was just created by
  // this apply run's template import" — the two produce different remedies
  // and the latter is not the agent's fault. Snapshotting after the import
  // mutated numberingDoc would conflate them.
  const preTemplateSourceNumIds = new Set<number>()
  for (const numEl of getChildrenNS(numberingDoc.documentElement!, NS.w, "num")) {
    const idStr = wAttr(numEl, "numId")
    if (idStr) {
      const id = parseInt(idStr, 10)
      if (Number.isFinite(id)) preTemplateSourceNumIds.add(id)
    }
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
    // If template numbering was imported, the numbering content-type +
    // rels registration below picks it up via numIdRemap.size > 0.
  }

  // 4. Resolve fromParagraph references and inject styles into styles.xml.
  // `requirements` is annotation-only; see ApplyConfig.requirements docs.
  if (config.requirements) {
    const declared = new Set(config.styles.map((s) => s.id))
    for (const styleId of Object.keys(config.requirements)) {
      if (!declared.has(styleId)) {
        throw new Error(
          `requirements: style "${styleId}" is not declared in styles[].\n` +
            `  Declared: [${[...declared].join(", ")}]\n` +
            `  Note: requirements is keyed by styleId (not a free-form label).\n` +
            `  If you renamed a style (e.g. to bind to an existing id like "a3"),\n` +
            `  update the requirements key to match — the side-by-side report\n` +
            `  pairs entries by exact id.`,
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
    // Snapshot prior cascade-resolved state BEFORE injection mutates
    // styles.xml. resolver was built at load; its cache reflects the
    // source's pre-apply definition for each styleId.
    const priorState = extractPriorDisplayFields(resolver, def.id)
    const priorUsage =
      priorState !== null ? (resolver.getStyleDefinition(def.id)?.usageCount ?? 0) : undefined
    const final = resolveStyleDef(def, parsed.paragraphs)
    const warnings = detectStyleResolutionWarnings(def, final)
    styleResolutions.push({
      styleId: def.id,
      userSpec: config.requirements?.[def.id] ?? null,
      resolved: extractDisplayFields(final),
      priorState,
      priorUsage,
      warnings: warnings.length > 0 ? warnings : undefined,
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
  const canonicalNameKey = canonicalStyleName
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
  // Collect all collisions before failing — agents iterating with --dry-run
  // benefit from seeing every conflict at once with the suggested override
  // styleId, instead of fixing one and rerunning to discover the next.
  const styleNameConflicts: Array<{
    def: StyleConfigEntry
    collidingId: string
    existingName: string
  }> = []
  const sourceStyleIds = new Set(
    getChildrenNS(stylesDoc.documentElement!, NS.w, "style")
      .map((s) => wAttr(s, "styleId"))
      .filter((id): id is string => id !== null),
  )
  for (const def of resolvedStyles) {
    // When name is omitted on an override (style already exists by id),
    // the existing <w:name> is preserved — no new name enters the doc,
    // so no collision is possible. Skip the check entirely.
    if (def.name === undefined && sourceStyleIds.has(def.id)) continue
    // When name is omitted on a create (new style), the engine defaults
    // name to id — check that value for collisions.
    const effectiveName = def.name ?? def.id
    const key = canonicalNameKey(effectiveName)
    const collidingId = sourceCanonicalToStyleId.get(key)
    if (collidingId && collidingId !== def.id) {
      const existingName = sourceCanonicalToOriginalName.get(key)!
      styleNameConflicts.push({ def, collidingId, existingName })
    }
  }
  if (styleNameConflicts.length > 0) {
    const lines: string[] = []
    lines.push(
      `${styleNameConflicts.length} style name collision(s) — Word treats matching names (and locale aliases) as the same built-in identity and would silently drop the new style's rPr at render:`,
    )
    for (const c of styleNameConflicts) {
      const effectiveName = c.def.name ?? c.def.id
      const aliasNote =
        c.existingName !== effectiveName ? ` (locale alias of existing "${c.existingName}")` : ""
      const nameDisplay = c.def.name !== undefined ? `name="${c.def.name}"` : `name omitted (defaults to id "${c.def.id}")`
      lines.push(
        `  styles[].id="${c.def.id}" ${nameDisplay}${aliasNote} → already used by source styleId="${c.collidingId}"`,
      )
    }
    lines.push("")
    lines.push("  Resolve each by either:")
    lines.push(`    (a) override the existing style: set styles[].id to the source styleId`)
    lines.push(`        (above) — minimum-change default. Declare only fields the user`)
    lines.push(`        spec explicitly requires; Mode A fromParagraph or piled-on`)
    lines.push(`        locale defaults (CJK 2-char indent, etc.) rewrite more than`)
    lines.push(`        the user asked to change.`)
    lines.push(`    (b) use the canonical English built-in name when the conflict is a`)
    lines.push(`        locale alias (e.g. name="Body Text" for id="BodyText").`)
    lines.push(`    (c) use a fresh styleId + name pair (e.g. id="BodyMain", name="正文主体")`)
    lines.push(`        when the source has no styleId for this role, or compresses many`)
    lines.push(`        roles onto one styleId so override can't separate them. \`name\` is`)
    lines.push(`        what end users see in Word's style panel — pick a human-readable`)
    lines.push(`        label, not the styleId.`)
    if (config.dryRun) {
      lines.unshift("=== Style Name Conflicts (dry-run; would FAIL on real apply) ===")
      console.error(lines.join("\n") + "\n")
    } else {
      throw new Error(lines.join("\n"))
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
  const installedSchemes: Array<{
    levels: ReadonlyArray<{
      level: number
      styleId: string
      restart: "continuous" | "perInstance" | "byHeading" | { atStyleChange: string }
    }>
    numId: string
    abstractNumId: string
    /** true when scheme.numId was set explicitly in config; false = engine allocated. */
    numIdExplicit: boolean
  }> = []
  if (config.numbering) {
    const numberingSchemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
    const declaredIds = new Set(config.styles.map((s) => s.id))
    const existingStyleIds = new Set<string>()
    for (const s of getChildrenNS(stylesDoc.documentElement!, NS.w, "style")) {
      const sid = wAttr(s, "styleId")
      if (sid) existingStyleIds.add(sid)
    }
    const validNumberingTargets = new Set([...declaredIds, ...existingStyleIds])

    // Pre-scan numbering.xml for existing <w:num> ids. Used in two places below:
    //   1. Explicit numId collision check — throws if config pins an id already
    //      present (duplicate <w:num w:numId="N"> is invalid OOXML).
    //   2. Auto-allocator exclusion — prevents the engine from picking an id that
    //      already exists even when no config scheme claims it.
    // Split into two sets so the collision error can attribute correctly:
    // ids the user's source carried in vs. ids this apply run's template
    // import just allocated. Same OOXML outcome (duplicate <w:num>), but
    // the agent's fix is different — for source collisions, pick another
    // numId or drop the pin; for template-import collisions, dropping the
    // pin is the natural fix because the agent didn't know the template
    // would grab that id.
    const allNumIds = new Set<number>()
    for (const numEl of getChildrenNS(numberingDoc.documentElement!, NS.w, "num")) {
      const idStr = wAttr(numEl, "numId")
      if (idStr) {
        const id = parseInt(idStr, 10)
        if (Number.isFinite(id)) allNumIds.add(id)
      }
    }
    const templateImportedNumIds = new Set<number>(
      [...allNumIds].filter((id) => !preTemplateSourceNumIds.has(id)),
    )

    // First pass: collect explicit numIds and detect collisions.
    const claimedNumIds = new Set<number>(allNumIds)
    for (const [schemeIdx, scheme] of numberingSchemes.entries()) {
      if (scheme.numId === undefined) continue
      const path = numberingSchemes.length === 1 ? "numbering" : `numbering[${schemeIdx}]`
      if (preTemplateSourceNumIds.has(scheme.numId)) {
        throw new Error(
          `${path}.numId = ${scheme.numId} collides with an existing <w:num w:numId="${scheme.numId}"/> in the source. ` +
            `Either pin to a different numId, or remove the explicit numId to let the engine allocate ` +
            `(the engine will pick an id not in use).`,
        )
      }
      if (templateImportedNumIds.has(scheme.numId)) {
        throw new Error(
          `${path}.numId = ${scheme.numId} was just claimed by template import in this apply run. ` +
            `Either pin to a different numId, or remove the explicit numId to let the engine allocate ` +
            `(the auto-allocator already skips template-imported ids).`,
        )
      }
      // Check config-internal duplicates.
      if (claimedNumIds.has(scheme.numId)) {
        const prev = numberingSchemes.findIndex(
          (s, i) => i < schemeIdx && s.numId === scheme.numId,
        )
        const prevPath = numberingSchemes.length === 1 ? "numbering" : `numbering[${prev}]`
        throw new Error(
          `${path}.numId ${scheme.numId} conflicts with ${prevPath}.numId ${scheme.numId} — ` +
            `two schemes cannot share the same numId. ` +
            `Set a different numId on one of them, or remove numId to let the engine allocate.`,
        )
      }
      claimedNumIds.add(scheme.numId)
    }

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
        // Bullet / none numFmts have no counter to display — skip the check.
        if (lvl.numFmt === "bullet" || lvl.numFmt === "none") continue
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
      const { numId, abstractNumId } = injectNumbering(numberingDoc, scheme, {
        claimedNumIds,
        requestedNumId: scheme.numId,
      })
      for (const lvl of scheme.levels) {
        attachNumberingToStyle(stylesDoc, lvl.styleId, numId, lvl.level)
      }
      installedSchemes.push({
        levels: scheme.levels.map((l) => ({
          level: l.level,
          styleId: l.styleId,
          restart: l.restart ?? "continuous",
        })),
        numId,
        abstractNumId,
        numIdExplicit: scheme.numId !== undefined,
      })
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

  // 5c. Apply edits (if config.edits present). Runs *after* style + numbering
  // install (so edit ops can reference styleIds + numIds we just created) and
  // *before* the rules pass (so pattern_rules / bulk_rules / assignments see
  // both pre-existing chrome paragraphs AND the agent-inserted content
  // uniformly — one regex match cleans both kinds of typed prefixes).
  //
  // Captions are resolved once here so they're available to both runEditOps
  // (emit-time identifier lookup) and the cross-reference post-pass below
  // (chapter SEQ injection, standardize re-emit, counter simulation). A pure
  // standardize run with no edits still goes through the cross-ref pipeline
  // after applyToBody — see step 7c.
  const resolvedCaptions = resolveCaptions(config.captions, stylesDoc)
  if (!config.dryRun) {
    for (const wmsg of resolvedCaptions.warnings) console.warn(`Warning: ${wmsg}`)
  }
  // Body asset registry — owns word/_rels/document.xml.rels and the shared
  // [Content_Types].xml accumulator. Constructed up-front (independent of
  // whether edits / HF are declared) so both the edits pass and the HF
  // pass route rels appends through the same instance; rId allocation
  // stays linear across image / hyperlink / header / footer entries.
  // flushTo at the end is a no-op when no subsystem registered anything.
  const bodyAssetRegistry = await DocxAssetRegistry.open(reader)
  let editsApplied = 0
  let editsTrackChanges = false
  let editsPreview: EditsPreviewEntry[] = []
  let editTouchedIndices: Set<number> | undefined
  let crossRefsTouched = false
  let listRestartApplied = false
  let bookmarkAllocator: BookmarkAllocator | null = null
  let pendingBackfills: PendingRefBackfill[] = []
  let pendingCaptionFills: PendingCaptionFill[] = []
  let pendingCaptionResets: PendingCaptionReset[] = []
  const panguWarnings: PanguWarning[] = []
  if (config.headerFooter) {
    // Pangu-spacing lint runs on every author-supplied prose surface —
    // `edits[]` and `headerFooter` are the two prose-bearing config
    // paths today. Run HF first so source labels appear in declaration
    // order in the report (HF is positionally above `edits[]`).
    panguWarnings.push(...lintPanguInHeaderFooter(config.headerFooter))
  }
  if (config.edits && config.edits.length > 0) {
    // Pangu-spacing lint: flag literal ASCII spaces between CJK and Latin/digit
    // glyphs in author-supplied text. Word's autoSpace inserts the gap at
    // render time; stacking a typed space on top renders too wide. Non-fatal:
    // the agent gets the warnings in the report and decides whether to scrub.
    panguWarnings.push(...lintPanguInEdits(config.edits))
    // Preview pass: resolve locators against the pre-edit document so the
    // report's "Edits Preview" + implicit-keep accounting use ORIGINAL paragraph
    // indices (the locators were authored against those). Cheap, non-mutating —
    // safe to run unconditionally even though `runEditOps` below also resolves
    // locators internally; the preview entries carry pre-edit-index metadata
    // that runEditOps doesn't expose.
    const preview = previewEditOps({
      documentDoc,
      parsedParagraphs: parsed.paragraphs,
      edits: config.edits,
    })
    editsPreview = preview.entries
    editTouchedIndices = preview.replacedOrDeletedIndices

    // Apply edits. Dry-run also walks this path so the cross-ref pipeline below
    // sees pendingCaptionFills / pendingBackfills / pendingCaptionResets for
    // freshly inserted captions, refs, and equations — otherwise the dry-run
    // caption-preview report only reflects standardize-reemit (existing source
    // captions) and predicts no text for the edits[] inserts. Disk-side effects
    // are gated separately by `if (!config.dryRun)` further down: documentDoc is
    // mutated in memory, image-registry binaries / rels are only staged into the
    // `replacements` map. Both are discarded on dry-run exit.
    const result = await runEditOps({
      documentDoc,
      parsedParagraphs: parsed.paragraphs,
      reader,
      edits: config.edits,
      trackChanges: config.trackChanges ?? false,
      author: config.author,
      stylesDoc,
      sections: parsed.sections,
      captions: resolvedCaptions.byIdentifier,
      imageRegistry: bodyAssetRegistry,
      numberingDoc,
    })
    editsApplied = result.report.applied
    editsTrackChanges = result.report.trackChanges
    ;({ bookmarkAllocator, pendingBackfills, pendingCaptionFills, pendingCaptionResets } =
      result.crossRefs)

    // Re-parse documentDoc — paragraph indices have shifted since edits inserted
    // new paragraphs. The rules pass below uses parsed.paragraphs directly, so it
    // needs the post-edit state. This must happen *before* applyToBody so the
    // action map (assignments / pattern_rules / bulk_rules) sees correct indices,
    // and *before* the cross-ref pipeline (step 7c) so chapter-SEQ injection
    // operates on the post-edit + post-restyle document.
    const reParser = new DocumentParser(documentDoc, resolver, numberingDoc)
    parsed = reParser.parse()
    fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver)
    hashToLetter.clear()
    for (const s of fpResult.summary) {
      hashToLetter.set(s.hash, s.label)
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
    const range =
      max > 0 ? `#${parsed.paragraphs[0]!.index}–#${parsed.paragraphs[max - 1]!.index}` : "(none)"
    throw new Error(
      `${where}: paragraph #${idx} not found. Document has ${max} indexed paragraphs (${range}). Paragraphs inside data tables are not indexed.`,
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

  // 7. Walk document.xml in order and apply actions to each indexed paragraph.
  // Pre-compute per-style pPr/rPr cascade (style + basedOn ancestors) so the
  // selective-strip pass only removes a paragraph's direct property when the
  // new style's cascade actually declares one for it. Without this gate the
  // strip falls through to docDefaults and silently loses chrome-baked values.
  const stylePPrCascade = buildStyleChildCascade(stylesDoc, "pPr")
  const styleRPrCascade = buildStyleChildCascade(stylesDoc, "rPr")
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
    editTouchedIndices,
    stylePPrCascade,
    styleRPrCascade,
  }
  applyToBody(documentDoc, ctx)

  // 7b. List-restart pass. perInstance single-level schemes fork a fresh
  // numId per contiguous run of paragraphs with the bound styleId, so each
  // list instance restarts at 1. Continuous schemes (the default) and
  // multi-level schemes (which use lvlRestart) are skipped inside the pass.
  // Must precede the caption counter simulator (which honors
  // <w:startOverride> when initializing counters); without this ordering
  // dry-run placeholder text would disagree with the live render. Runs
  // unconditionally so standardize-only configs (no edits, no cross-refs
  // beyond captions) still trigger it.
  if (installedSchemes.length > 0 && !listRestartApplied) {
    const headingStyleIds = buildHeadingStyleIdSet(stylesDoc)
    applyListRestartPass(documentDoc, numberingDoc, installedSchemes, headingStyleIds)
    listRestartApplied = true
  }

  // 7c. Cross-reference post-pass. Runs after applyToBody so paragraphs
  // whose pStyle was set by assignments / pattern_rules / bulk_rules are
  // visible to injectChapterCounters and the caption simulator (the bug
  // this restructure fixes: previously the pipeline ran inside the edits
  // branch, before applyToBody, so only emit-time-styled paragraphs got
  // chapter SEQ counters). Fires whenever the agent declared a captions
  // config (chapter SEQs to inject, pre-existing caption paragraphs to
  // re-emit) OR the edits emitted REFs / captions / resets.
  //
  // Dry-run runs the same pipeline so the change report can preview what
  // chapter SEQs / caption text WOULD be produced; documentDoc mutations
  // are discarded with the in-memory document because dry-run never writes
  // to disk. captionsPreview captures the summary surfaced in the report.
  //
  // INVARIANT (dry-run ↔ real-apply caption text consistency): the
  // `Predicted text` lines surfaced in the dry-run report show the full
  // visible caption string: field-portion (prefix + chapter parts + counter
  // + sub + suffix) from `fullCaptionText`, plus bodySeparator + body from
  // the fill record. The field-portion contract is enforced structurally —
  // both modes call the same `injectChapterCounters` → `standardizeCaptions`
  // → `simulateCaptions` sequence on the same documentDoc, and
  // `captionsPreview.samples` is populated from the very `fullCaptionText`
  // map that real-apply backfill also reads. Future edits to this pipeline
  // (counter sim, fill ordering, chapter STYLEREF resolution, format
  // override) MUST preserve this unconditional shared-pipeline shape. If a
  // step becomes mode-specific, verify post-change that dry-run `Predicted
  // text` still matches the full visible content of the corresponding caption
  // paragraph in the written docx; a divergence makes the dry-run preview a
  // liar.
  //
  // Steps: (1) inject hidden auto-chapter SEQ fields into outline
  // paragraphs whose style is referenced under any captions.chapterPrefix
  // with a `format` override. (2) Standardize re-emit: rebuild the
  // pre-body run sequence of any caption paragraph already in the doc
  // (source-doc captions or output from a prior apply). (3) Simulate
  // counters and backfill REF placeholder text + commit bookmark
  // wrapping. (4) Flip settings.xml's updateFields flag (deferred to
  // serialize, gated on `crossRefsTouched`).
  let captionsPreview: {
    chapterSeqsInjected: number
    standardizeReemitted: number
    freshlyEmitted: number
    samples: Array<{ identifier: string; text: string }>
  } | null = null
  {
    const hasCaptionWork =
      resolvedCaptions.byIdentifier.size > 0 ||
      pendingBackfills.length > 0 ||
      pendingCaptionFills.length > 0 ||
      pendingCaptionResets.length > 0
    if (hasCaptionWork) {
      if (!bookmarkAllocator) bookmarkAllocator = new BookmarkAllocator(documentDoc)

      // Inject hidden auto-chapter SEQ fields into paragraphs whose style is
      // referenced under any captions.chapterPrefix with a `format` override.
      // Idempotent; runs after applyToBody so paragraphs restyled to an
      // outline style by assignments / pattern_rules / bulk_rules pick up
      // the counter too.
      const chapterSeqsInjected = injectChapterCounters(documentDoc, resolvedCaptions.byIdentifier)
      if (chapterSeqsInjected > 0) {
        ensureHiddenChapterCounterStyle(stylesDoc)
      }

      // Standardize re-emit: any caption paragraph already in the doc
      // (source-doc captions, or output from a prior apply) gets its
      // pre-body run sequence rebuilt against the current captions config.
      // Skip paragraphs freshly emitted in this pass — their emit already
      // used the current config.
      const freshlyEmitted = new Set<Element>(pendingCaptionFills.map((f) => f.paragraph))
      const standardizeResult = standardizeCaptions(
        documentDoc,
        resolvedCaptions.byIdentifier,
        bookmarkAllocator,
        freshlyEmitted,
      )
      for (const w of standardizeResult.warnings) console.warn(w)
      const allCaptionFills = [...pendingCaptionFills, ...standardizeResult.fills]

      // Caption counter simulator: walks the body, advances per-identifier
      // counters, resolves STYLEREF chapter prefixes, returns fieldValues
      // (per result text element) + fullCaptionText (per caption
      // paragraph). Backfill below uses fullCaptionText for caption-class
      // REF targets; outline-numbered targets fall through to the
      // numbering counter sim's lvlText.
      const captionSimOutput =
        allCaptionFills.length > 0 || pendingCaptionResets.length > 0
          ? simulateCaptions(documentDoc, {
              fills: allCaptionFills,
              resets: pendingCaptionResets,
              configs: resolvedCaptions.byIdentifier,
              outlineParagraphs: buildOutlineParagraphsMap(
                documentDoc,
                numberingDoc,
                stylesDoc,
                resolvedCaptions.styleIdToName,
              ),
            })
          : { fieldValues: new Map<Element, string>(), fullCaptionText: new Map<Element, string>() }
      for (const [el, text] of captionSimOutput.fieldValues) {
        el.textContent = text
      }
      if (captionSimOutput.fieldValues.size > 0 || captionSimOutput.fullCaptionText.size > 0) {
        crossRefsTouched = true
      }

      // Capture dry-run preview stats. First 5 caption samples reported in
      // pipeline order — agent sees predicted text without running real
      // apply. Counter text comes from the same `fullCaptionText` map that
      // real-apply backfill reads; body text is appended from the fill record
      // so the preview shows the full visible caption string (counter + body).
      if (config.dryRun) {
        const previewSamples: Array<{ identifier: string; text: string }> = []
        // Iterate allCaptionFills in pipeline order; map .paragraph → fullText.
        for (const fill of allCaptionFills) {
          if (previewSamples.length >= 5) break
          const counterText = captionSimOutput.fullCaptionText.get(fill.paragraph)
          if (counterText !== undefined) {
            const captionConfig = resolvedCaptions.byIdentifier.get(fill.identifier)
            const body = fill.bodyText
            const text =
              body !== undefined && captionConfig !== undefined
                ? counterText + captionConfig.bodySeparator + body
                : counterText
            previewSamples.push({ identifier: fill.identifier, text })
          }
        }
        captionsPreview = {
          chapterSeqsInjected,
          standardizeReemitted: standardizeResult.fills.length,
          freshlyEmitted: pendingCaptionFills.length,
          samples: previewSamples,
        }
      }

      if (bookmarkAllocator.hasAllocations() || captionSimOutput.fullCaptionText.size > 0) {
        // Resolve each pending backfill's target via the allocator. Every
        // emit path (paragraph-index ref, anchor ref already-adopted,
        // anchor ref forward) registered the name with the allocator by
        // now, so a single resolve handles all three. A miss here is an
        // engine bug — pre-scan + emit-time presence checks should have
        // failed earlier.
        const resolvedBackfills: Array<{
          placeholderTextEl: Element
          targetParagraph: Element
          targetName: string
          display: "full" | "label" | "number"
        }> = pendingBackfills.map((pending) => {
          const rec = bookmarkAllocator!.resolveByName(pending.targetName)
          if (!rec) {
            throw new Error(
              `InlineRef: target bookmark "${pending.targetName}" could not be resolved post-emit. ` +
                `This is an engine invariant violation; pre-scan should have caught a missing anchor.`,
            )
          }
          return { ...pending, targetParagraph: rec.element }
        })
        // Detached-target guard: a later edit op may have removed the
        // paragraph an earlier InlineRef pointed at. The stale-element
        // check below catches this BEFORE the counter sim (which iterates
        // only attached paragraphs and would silently leave the placeholder
        // empty). Throw with the original locator so the agent can see
        // which ref needs updating.
        for (const pending of resolvedBackfills) {
          if (!isAttachedToDoc(pending.targetParagraph, documentDoc)) {
            throw new Error(
              `InlineRef target paragraph was removed by a later edit op. ` +
                `Reorder edits so the ref op runs before any op that replaces / deletes the target, ` +
                `or remove the conflicting edit.`,
            )
          }
        }
        const counters = simulateNumberingCounters(documentDoc, numberingDoc, stylesDoc)
        for (const pending of resolvedBackfills) {
          // Caption-class targets: REF \h returns the SEQ-rendered text
          // (prefix + chapter + counter + suffix). label / number / full
          // all use the same primary bookmark (display:"full" is rejected
          // at pre-scan — see InlineRef emit-time check in edit-engine).
          const captionText = captionSimOutput.fullCaptionText.get(pending.targetParagraph)
          if (captionText !== undefined) {
            pending.placeholderTextEl.textContent = captionText
            continue
          }
          // Outline-numbered targets: fall through to lvlText backfill.
          const rendered = counters.get(pending.targetParagraph)
          if (!rendered && pending.display !== "full") continue
          const text =
            pending.display === "number"
              ? (rendered?.number ?? "")
              : pending.display === "full"
                ? extractParagraphText(pending.targetParagraph)
                : (rendered?.label ?? "")
          pending.placeholderTextEl.textContent = text
        }
        bookmarkAllocator.commit(documentDoc)
        crossRefsTouched = true
      }
    }
  }

  // 7d. Target-set + vs-direct analysis (dry-run only). For each declared
  // style, find the paragraphs whose final pStyle is this style (after
  // edits + rules pass), then classify each declared field's per-paragraph
  // direct values as override / redundant / new. Surfaces the gap left by
  // the styles-cascade Δ-line — sparse-by-design's real invariant is the
  // direct-format layer this pass touches. Skip on real apply: same compute
  // cost, and agent already saw it during iteration.
  if (config.dryRun) {
    populateVsDirectReports(
      documentDoc,
      parsed.paragraphs,
      config.styles,
      styleResolutions,
      resolver,
    )
  }

  // 7e. Page setup — mutates every relevant <w:sectPr> in documentDoc before
  // serialization. Sparse-by-design: only declared fields change; per-section
  // overrides layer on top of top-level defaults.
  let pageSetupReport: PageSetupReport | undefined
  if (config.pageSetup) {
    pageSetupReport = applyPageSetup(documentDoc, config.pageSetup)
  }

  // 8. Replacement map — every following subsystem stages its part-level
  //    mutations here (new HF parts + their rels, settings.xml flips, the
  //    final body / styles / numbering serializations). Single writer
  //    pass at the end consumes it.
  const replacements = new WritableArchive()

  // 7f. Header / footer — generates header*.xml / footer*.xml parts, plugs
  //     references into every sectPr, ensures the Header/Footer paragraph
  //     styles exist, and (when an `even` variant is declared) flips the
  //     evenAndOddHeaders flag in settings.xml.
  let headerFooterReport: HeaderFooterReport | undefined
  let headerFooterBindingReport: HeaderFooterBindingReport | undefined
  if (config.headerFooter) {
    ensureHeaderFooterStyles(stylesDoc)
    headerFooterReport = await applyHeaderFooter(
      reader,
      documentDoc,
      config.headerFooter,
      bodyAssetRegistry.getPartRels(),
      bodyAssetRegistry.getContentTypes(),
      replacements,
      stylesDoc,
    )
    headerFooterBindingReport = applyHeaderFooterBinding(documentDoc, headerFooterReport)
    if (headerFooterReport.parts.some((p) => p.hasHyperlinks)) {
      // HF emitted at least one external hyperlink — the runs carry the
      // built-in `Hyperlink` character style; inject it into styles.xml on
      // first encounter (no-op when already present).
      ensureHyperlinkCharStyle(stylesDoc)
    }
    // HF config is the source of truth for evenAndOddHeaders: set when an
    // `even` variant is declared anywhere, clear otherwise. Without the
    // clear branch, a re-run that drops `even` leaves the flag set and
    // Word renders even pages as blank.
    await setEvenAndOddHeadersFlag(
      reader,
      replacements,
      headerFooterReport.hasEven,
      bodyAssetRegistry,
    )
  }

  // Stage stylesDoc / numberingDoc / documentDoc AFTER all in-place
  // mutation is finished — HF binding writes to documentDoc, styles
  // ensures write to stylesDoc.
  replacements.set("word/styles.xml", serializeXml(stylesDoc))
  if (numberingDoc) replacements.set("word/numbering.xml", serializeXml(numberingDoc))
  replacements.set("word/document.xml", serializeXml(documentDoc))
  // Only serialize theme back when we actually mutated it; otherwise keep the
  // original theme1.xml bytes untouched to avoid spurious whitespace changes.
  if (themeMutated && themeDoc) {
    replacements.set("word/theme/theme1.xml", serializeXml(themeDoc))
  }
  // Make sure numbering.xml is referenced from [Content_Types].xml and
  // word/_rels/document.xml.rels when newly created. Routes through the
  // SHARED ContentTypes and body PartRels accumulators so the final
  // flushTo (below) preserves these additions; the older path that
  // mutated the staged [Content_Types].xml / rels strings directly raced
  // with the registry's flushTo and got silently clobbered.
  const numberingTouched = !!config.numbering || !!templateImport?.numIdRemap.size
  if (numberingTouched) {
    const sharedCT = bodyAssetRegistry.getContentTypes()
    const sharedRels = bodyAssetRegistry.getPartRels()
    sharedCT.ensureOverride(
      "/word/numbering.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    )
    if (!sharedRels.hasRelTo("numbering.xml")) {
      sharedRels.appendRel(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "numbering.xml",
      )
    }
  }
  // Cross-references emitted in this run — flip settings.xml's
  // <w:updateFields> flag so Word resolves each REF on next open without
  // the user manually pressing Ctrl+A then F9. MUST run before the
  // bodyAssetRegistry flushes below: when source lacks settings.xml the
  // fabrication path appends settings Override + Relationship to the
  // shared accumulators, and those additions need to be in the
  // accumulator state at flush time.
  if (crossRefsTouched) {
    await ensureUpdateFieldsFlag(reader, replacements, bodyAssetRegistry)
  }
  // Body asset registry flushes its binary media (word/media/imageN.*) + the
  // body's rels (word/_rels/document.xml.rels). HF / settings-fabrication /
  // numbering Override registrations all appended to these accumulators
  // earlier in the pipeline; this single flush carries every accumulated
  // entry. ContentTypes is shared across body and HF parts — flushed once
  // immediately after.
  bodyAssetRegistry.flushTo(replacements)
  bodyAssetRegistry.getContentTypes().flushTo(replacements)
  // 8b. Dry-run skips write + post-write validation; the agent iterates on
  // the in-memory report alone. Otherwise: write, then run the comprehensive
  // bundle check (XML well-formedness, CT_* schema, whitespace preservation,
  // cross-part references, content types, relationship Ids).
  //
  // Baseline-diff validation: errors already present in the source file are
  // reported as warnings, not fatal. Only errors *introduced* by this apply
  // run cause the output to be deleted (unless --allow-validation-warnings
  // is set, in which case new errors are also non-fatal).
  if (!config.dryRun) {
    // Capture source baseline before writing output.
    const baselineErrors = await validateDocxFile(source)
    // Multiset (count per key) instead of a plain Set so that N identical
    // errors in the source only absorb N matching errors in the output.
    // A Set would let 1 baseline entry mask any number of output errors
    // with the same normalized key, silently swallowing introduced errors.
    const baselineCounts = new Map<string, number>()
    for (const e of baselineErrors) {
      const k = normalizeValidationError(e)
      baselineCounts.set(k, (baselineCounts.get(k) ?? 0) + 1)
    }

    await reader.copyAndModify(output, replacements)
    const outputErrors = await validateDocxFile(output)

    // Consume baseline slots greedily: each output error checks if a slot
    // remains for its key and decrements if so (pre-existing), otherwise
    // counts as new (introduced by this run).
    const available = new Map(baselineCounts)
    const newErrors: (typeof outputErrors)[number][] = []
    const preExistingErrors: (typeof outputErrors)[number][] = []
    for (const e of outputErrors) {
      const k = normalizeValidationError(e)
      const n = available.get(k) ?? 0
      if (n > 0) {
        available.set(k, n - 1)
        preExistingErrors.push(e)
      } else {
        newErrors.push(e)
      }
    }

    if (preExistingErrors.length > 0) {
      console.error(
        `Validation: ${preExistingErrors.length} pre-existing error(s) carried through from source (non-fatal).`,
      )
    }

    if (newErrors.length > 0) {
      const lines = newErrors.slice(0, 20).map((e) => `  ${e.part}: ${e.message}`)
      const more = newErrors.length > 20 ? `\n  …${newErrors.length - 20} more` : ""
      if (config.allowValidationWarnings) {
        console.error(
          `Validation: ${newErrors.length} new error(s) introduced by this run (--allow-validation-warnings set; output kept):\n${lines.join("\n")}${more}`,
        )
      } else {
        if (existsSync(output)) {
          try {
            unlinkSync(output)
          } catch {}
        }
        console.error(
          `Validation FAILED — ${newErrors.length} new error(s) introduced by this run:\n${lines.join("\n")}${more}\n` +
            `  (Use --allow-validation-warnings to keep the output despite new errors.)`,
        )
        process.exit(1)
      }
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

  // Dry-run only: best-effort scan of `edits[]` Blocks for typed-prefix
  // residue when the Block's styleId is bound to an auto-numbering scheme.
  // Catches inserts typed by hand; chrome retags are already covered by
  // `unstrippedByStyle`. Skip on real apply — agent already saw it.
  const manualNumberingDetected = config.dryRun
    ? detectManualNumbering(config.edits, new Set(numLvlTextByStyle.keys()))
    : undefined

  // Dry-run only: echo the text of every excluded paragraph so the agent can
  // verify the indices still point at what they intended. exclude entries are
  // bare numbers — easy to drift silently when document order shifts.
  const excludeSamples =
    config.dryRun && excludeSet.size > 0
      ? [...excludeSet]
          .sort((a, b) => a - b)
          .map((idx) => {
            const p = parsed.paragraphs.find((q) => q.index === idx)
            const text = p?.text.trim() ?? ""
            return {
              index: idx,
              snippet: text ? text.slice(0, 60) + (text.length > 60 ? "…" : "") : "(empty)",
            }
          })
      : undefined

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
    manualNumberingDetected,
    excludeSamples,
    numberingBindings,
    numberingAllocation: installedSchemes.map((s, i) => ({
      schemeIndex: i,
      numId: s.numId,
      explicit: s.numIdExplicit,
    })),
    templateImport,
    editsPreview,
    captionsPreview,
    panguWarnings: panguWarnings.length > 0 ? panguWarnings : undefined,
    pageSetup: pageSetupReport,
    headerFooter: headerFooterReport,
    headerFooterBinding: headerFooterBindingReport,
    totalParagraphs: config.dryRun
      ? (parsed.paragraphs[parsed.paragraphs.length - 1]?.index ?? parsed.paragraphs.length)
      : undefined,
  })

  // Dry-run also invokes runEditOps now (so the cross-ref pipeline can see
  // pending caption fills from edits[]) — the message would mislead the agent
  // into thinking changes hit disk. The Edits Preview block in the report
  // already covers the dry-run case.
  if (editsApplied > 0 && !config.dryRun) {
    console.error(
      `\nEdit pass: ${editsApplied} op(s) applied${editsTrackChanges ? " (track-changes)" : ""}. New paragraphs participated in pattern_rules / bulk_rules cleanup uniformly with the original chrome.`,
    )
  }
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
/** Build styleId → set of `<childName>` (`pPr` or `rPr`) child localNames the
 * style's cascade declares (style + every basedOn ancestor). The rules pass
 * uses this to strip a paragraph's direct properties only when the new style
 * actually provides values for them; otherwise stripping would fall back to
 * docDefaults / Normal and lose chrome-baked values (line spacing, font
 * size, bold, etc.). */
function buildStyleChildCascade(
  stylesDoc: Document,
  childName: "pPr" | "rPr",
): Map<string, Set<string>> {
  const w = NS.w
  const styles = getChildrenNS(stylesDoc.documentElement!, w, "style")
  const directChildren = new Map<string, Set<string>>()
  const basedOn = new Map<string, string | null>()
  for (const s of styles) {
    const id = wAttr(s, "styleId")
    if (!id) continue
    const direct = new Set<string>()
    const container = firstChildNS(s, w, childName)
    if (container) {
      for (const c of getChildren(container)) {
        if (c.namespaceURI === w) direct.add(c.localName!)
      }
    }
    directChildren.set(id, direct)
    const bo = firstChildNS(s, w, "basedOn")
    basedOn.set(id, bo ? wAttr(bo, "val") : null)
  }
  const cascade = new Map<string, Set<string>>()
  const walk = (id: string, seen: Set<string>): Set<string> => {
    const cached = cascade.get(id)
    if (cached) return cached
    if (seen.has(id)) return new Set()
    seen.add(id)
    const merged = new Set(directChildren.get(id) ?? [])
    const parent = basedOn.get(id)
    if (parent && directChildren.has(parent)) {
      for (const n of walk(parent, seen)) merged.add(n)
    }
    cascade.set(id, merged)
    return merged
  }
  for (const id of directChildren.keys()) walk(id, new Set())
  return cascade
}

/** Build `outlineParagraphs: Map<Element, { styleName, rendered }>` for the
 * caption simulator. Walks the body, finds paragraphs whose pStyle matches
 * a styleId referenced by some caption's chapterPrefix, looks up its
 * rendered counter number, returns the pair. Caption simulator uses this
 * to resolve STYLEREF chapter prefixes. */
function buildOutlineParagraphsMap(
  documentDoc: Document,
  numberingDoc: Document | null,
  stylesDoc: Document | null,
  styleIdToName: Map<string, string>,
): Map<Element, { styleName: string; rendered: string }> {
  const out = new Map<Element, { styleName: string; rendered: string }>()
  if (styleIdToName.size === 0) return out
  const counters = simulateNumberingCounters(documentDoc, numberingDoc, stylesDoc)
  const w = NS.w
  for (const [paraEl, rendered] of counters) {
    const pPr = firstChildNS(paraEl, w, "pPr")
    if (!pPr) continue
    const pStyle = firstChildNS(pPr, w, "pStyle")
    if (!pStyle) continue
    const styleId = wAttr(pStyle, "val")
    if (!styleId) continue
    const styleName = styleIdToName.get(styleId)
    if (!styleName) continue
    out.set(paraEl, { styleName, rendered: rendered.number })
  }
  return out
}
