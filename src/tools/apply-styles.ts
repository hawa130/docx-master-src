import { readFileSync, unlinkSync, existsSync, copyFileSync } from "node:fs"
import { resolve } from "node:path"
import JSZip from "jszip"
import { DOMParser } from "@xmldom/xmldom"
import { DocxReader, serializeXml } from "../core/reader.ts"
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

interface ApplyConfig {
  source: string
  output: string
  styles: StyleConfigEntry[]
  numbering?: NumberingConfig
  assignments?: AssignmentEntry[]
  bulk_rules?: BulkRule[]
  exclude?: number[]
}

interface FlagRecord {
  paraIndex: number
  reason: string
}

interface RestyleStat {
  styleId: string
  count: number
}

async function main() {
  const configPath = process.argv[2]
  if (!configPath) {
    console.error("Usage: node scripts/apply_styles.js <config.json>")
    process.exit(1)
  }
  let config: ApplyConfig
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"))
  } catch (err) {
    console.error(`Cannot read config: ${(err as Error).message}`)
    process.exit(1)
  }

  if (!config.source || !config.output) {
    console.error("config.source and config.output are required")
    process.exit(1)
  }
  const source = resolve(config.source)
  const output = resolve(config.output)
  if (source === output) {
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
  // 1. Copy source → output
  copyFileSync(source, output)

  // 2. Open the COPY as our working zip
  const reader = await DocxReader.open(output)
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

  // 4. Resolve fromParagraph references, then inject styles into styles.xml
  const resolvedStyles = config.styles.map((def) => resolveStyleDef(def, parsed.paragraphs))
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
    const newNumId = injectNumbering(numberingDoc, config.numbering)
    // Update the relevant styles' numPr
    for (const lvl of config.numbering.levels) {
      numIdMap.set(lvl.styleId, newNumId)
      attachNumberingToStyle(stylesDoc, lvl.styleId, newNumId, lvl.level)
    }
  }

  // 6. Build action map for paragraphs
  const excludeSet = new Set(config.exclude ?? [])
  const assignmentMap = new Map<number, AssignmentEntry>()
  for (const a of config.assignments ?? []) assignmentMap.set(a.para, a)
  const bulkMap = new Map<string, string>()
  for (const b of config.bulk_rules ?? []) bulkMap.set(b.fingerprint, b.style)

  const restyleStats = new Map<string, number>()
  const flags: FlagRecord[] = []
  let manualNumberingRemoved: Map<string, number> = new Map()

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
  const ctx: ApplyContext = {
    excludeSet,
    assignmentMap,
    bulkMap,
    paragraphs: parsed.paragraphs,
    restyleStats,
    flags,
    manualNumberingRemoved,
    numLvlTextByStyle,
    config,
  }
  applyToBody(documentDoc, ctx)

  // 8. Serialize and write
  const replacements = new Map<string, string>()
  replacements.set("word/styles.xml", serializeXml(stylesDoc))
  if (numberingDoc) replacements.set("word/numbering.xml", serializeXml(numberingDoc))
  replacements.set("word/document.xml", serializeXml(documentDoc))
  // Make sure numbering.xml is referenced from [Content_Types].xml when newly created
  if (config.numbering) {
    await ensureNumberingContentType(reader, replacements)
    await ensureNumberingRelationship(reader, replacements)
  }
  // Use original source for the base zip
  await reader.copyAndModify(output, replacements)

  // 9. Validate by re-parsing modified entries
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

  // 10. Print report
  printReport({
    injected,
    updated,
    restyleStats,
    flags,
    manualNumberingRemoved,
    derivedFrom,
    output,
  })
}

interface ApplyContext {
  excludeSet: Set<number>
  assignmentMap: Map<number, AssignmentEntry>
  bulkMap: Map<string, string>
  paragraphs: ParsedParagraph[]
  restyleStats: Map<string, number>
  flags: FlagRecord[]
  manualNumberingRemoved: Map<string, number>
  numLvlTextByStyle: Map<string, string[]>
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

  if (a) {
    action = a.action
    targetStyle = a.style
    reason = a.reason
  } else {
    const bulkStyle = ctx.bulkMap.get(para.fingerprint)
    if (bulkStyle) {
      action = "restyle"
      targetStyle = bulkStyle
    }
  }

  if (action === "flag" && reason) {
    ctx.flags.push({ paraIndex: para.index, reason })
    return
  }
  if (action !== "restyle" || !targetStyle) return

  // apply restyle
  setParagraphStyle(pEl, targetStyle)
  stripConflictingDirectFormatting(pEl, targetStyle, ctx)
  ctx.restyleStats.set(targetStyle, (ctx.restyleStats.get(targetStyle) ?? 0) + 1)

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
    throw new Error(
      `style "${def.id}": fromParagraph #${def.fromParagraph} not found (paragraph indices are 1-based and must refer to indexed paragraphs — content inside data/form tables is not indexed)`,
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
  derivedFrom: Map<string, number>
  output: string
}) {
  const lines: string[] = []
  lines.push("=== Change Report ===")
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
  if (args.flags.length > 0) {
    lines.push(`Flagged (not modified): ${args.flags.length}`)
    for (const f of args.flags) {
      lines.push(`  #${f.paraIndex}: ${f.reason}`)
    }
    lines.push("")
  }
  lines.push("Validation: PASS")
  lines.push(`Output: ${args.output}`)
  console.log(lines.join("\n"))
}

main()
