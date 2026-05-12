import { unlinkSync, existsSync, copyFileSync, mkdirSync } from "node:fs"
import { dirname, resolve } from "node:path"
import { DocxReader, serializeXml } from "@lib/xml/reader.ts"
import { importTemplateStyles, type ImportResult } from "@lib/apply/template-import.ts"
import { StyleResolver, applyThemeFontOverrides } from "@lib/parse/style-resolver.ts"
import { DocumentParser } from "@lib/parse/document-parser.ts"
import { Fingerprinter } from "@lib/parse/fingerprint.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import {
  blankNumberingDoc,
  blankStylesDoc,
  ensureNumberingContentType,
  ensureNumberingRelationship,
} from "@lib/xml/docx-plumbing.ts"
import { applyToBody } from "@lib/apply/para-mutation.ts"
import { walkIndexedParagraphs } from "@lib/edit/locator.ts"
import { analyzeVsDirect } from "@lib/shared/vs-direct.ts"
import type { ParsedParagraph } from "@lib/parse/types.ts"
import {
  attachNumberingToStyle,
  injectNumbering,
  resolveSuff,
} from "@lib/apply/numbering-mutation.ts"
import { applyListRestartPass } from "@lib/apply/list-restart.ts"
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
import type { ImageAssetRegistry } from "@lib/edit/image-asset.ts"
import { simulateNumberingCounters, extractParagraphText } from "@lib/apply/numbering-counter.ts"
import { ensureUpdateFieldsFlag } from "@lib/apply/settings-mutation.ts"
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
  // Collect all collisions before failing — agents iterating with --dry-run
  // benefit from seeing every conflict at once with the suggested override
  // styleId, instead of fixing one and rerunning to discover the next.
  const styleNameConflicts: Array<{
    def: StyleConfigEntry
    collidingId: string
    existingName: string
  }> = []
  for (const def of resolvedStyles) {
    const key = canonicalNameKey(def.name)
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
      const aliasNote =
        c.existingName !== c.def.name ? ` (locale alias of existing "${c.existingName}")` : ""
      lines.push(
        `  styles[].id="${c.def.id}" name="${c.def.name}"${aliasNote} → already used by source styleId="${c.collidingId}"`,
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
    levels: ReadonlyArray<{ level: number; styleId: string; restart: "continuous" | "perInstance" }>
    numId: string
    abstractNumId: string
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
      const { numId, abstractNumId } = injectNumbering(numberingDoc, scheme)
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
  let imageRegistry: ImageAssetRegistry | null = null
  let editsApplied = 0
  let editsTrackChanges = false
  let editsPreview: EditsPreviewEntry[] = []
  let editTouchedIndices: Set<number> | undefined
  let crossRefsTouched = false
  let listRestartApplied = false
  if (config.edits && config.edits.length > 0) {
    if (config.dryRun) {
      // Dry-run: resolve locators + blocker check, but don't mutate. Lets the
      // change report show predicted edit effect alongside style + rule effect,
      // and lets implicit-keep accounting subtract paragraphs the edits will
      // replace/delete (otherwise those read as false-positive coverage gaps).
      const preview = previewEditOps({
        documentDoc,
        parsedParagraphs: parsed.paragraphs,
        edits: config.edits,
      })
      editsPreview = preview.entries
      editTouchedIndices = preview.replacedOrDeletedIndices
    } else {
      const result = await runEditOps({
        documentDoc,
        parsedParagraphs: parsed.paragraphs,
        reader,
        edits: config.edits,
        trackChanges: config.trackChanges ?? false,
        stylesDoc,
        sections: parsed.sections,
      })
      imageRegistry = result.imageRegistry
      editsApplied = result.report.applied
      editsTrackChanges = result.report.trackChanges

      // Run the list-restart pass BEFORE the cross-ref simulator so the
      // simulator's counter values reflect what Word will actually render.
      // For continuous schemes (the default) this pass is a no-op; for
      // perInstance schemes it forks numIds + writes <w:startOverride>, and
      // the simulator below honors those overrides when initializing
      // counters. Without this ordering, dry-run placeholder text would
      // disagree with the live render.
      if (installedSchemes.length > 0) {
        applyListRestartPass(documentDoc, numberingDoc, installedSchemes)
        listRestartApplied = true
      }

      // Cross-reference post-pass. (1) Simulate numbering counters against the
      // current document state — gives us the rendered label / number text per
      // target paragraph. (2) Backfill each pending REF's placeholder run so
      // Word users see correct text before the first F9 / "update fields"
      // prompt. (3) Wrap target paragraphs with <w:bookmarkStart>/<w:bookmarkEnd>
      // pairs so REF can resolve. (4) Flip settings.xml's updateFields flag.
      const { bookmarkAllocator, pendingBackfills } = result.crossRefs
      if (bookmarkAllocator.hasAllocations()) {
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
          const rec = bookmarkAllocator.resolveByName(pending.targetName)
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

      // Re-parse documentDoc — paragraph indices have shifted since edits inserted
      // new paragraphs. The rules pass below uses parsed.paragraphs directly, so it
      // needs the post-edit state.
      const reParser = new DocumentParser(documentDoc, resolver, numberingDoc)
      parsed = reParser.parse()
      fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver)
      hashToLetter.clear()
      for (const s of fpResult.summary) {
        hashToLetter.set(s.hash, s.label)
      }
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
  // Already invoked above when cross-refs ran (must precede the simulator);
  // this fallback covers configs without edits / cross-refs.
  if (installedSchemes.length > 0 && !listRestartApplied) {
    applyListRestartPass(documentDoc, numberingDoc, installedSchemes)
    listRestartApplied = true
  }

  // 7c. Target-set + vs-direct analysis (dry-run only). For each declared
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

  // 8. Serialize and write
  const replacements = new Map<string, string | Uint8Array>()
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
  // Image registry from edit-pass flushes its staged binaries + rels +
  // content-type updates. No-op when no images were embedded.
  if (imageRegistry) imageRegistry.flushTo(replacements)
  // Cross-references emitted in this run — flip settings.xml's
  // <w:updateFields> flag so Word resolves each REF on next open without
  // the user manually pressing Ctrl+A then F9.
  if (crossRefsTouched) {
    await ensureUpdateFieldsFlag(reader, replacements)
  }
  // 8b. Dry-run skips write + post-write validation; the agent iterates on
  // the in-memory report alone. Otherwise: write, then run the comprehensive
  // bundle check (XML well-formedness, CT_* schema, whitespace preservation,
  // cross-part references, content types, relationship Ids).
  if (!config.dryRun) {
    await reader.copyAndModify(output, replacements)
    const errors = await validateDocxFile(output)
    if (errors.length > 0) {
      if (existsSync(output)) {
        try {
          unlinkSync(output)
        } catch {}
      }
      const lines = errors.slice(0, 20).map((e) => `  ${e.part}: ${e.message}`)
      const more = errors.length > 20 ? `\n  …${errors.length - 20} more` : ""
      console.error(`Validation FAILED:\n${lines.join("\n")}${more}`)
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
    templateImport,
    editsPreview,
  })

  if (editsApplied > 0) {
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
