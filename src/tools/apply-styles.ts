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
  spaceBefore?: number // pt
  spaceAfter?: number // pt
  firstLineIndent?: string | number | null
  hangingIndent?: string | number | null
  outlineLevel?: number
}

interface NumberingConfig {
  levels: Array<{
    level: number
    format: string
    text: string
    styleId: string
    start?: number
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

  // 4. Inject styles into styles.xml (modify or create)
  const injected: string[] = []
  const updated: string[] = []
  for (const def of config.styles) {
    const result = upsertStyle(stylesDoc, def)
    if (result === "created") injected.push(def.id)
    else updated.push(def.id)
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

  // Build a map of styleId → numbering lvlText (to support manual prefix stripping)
  const numLvlTextByStyle = new Map<string, string>()
  if (config.numbering) {
    for (const lvl of config.numbering.levels) {
      numLvlTextByStyle.set(lvl.styleId, lvl.text)
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
  numLvlTextByStyle: Map<string, string>
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

  // Manual numbering text stripping
  const lvlText = ctx.numLvlTextByStyle.get(targetStyle)
  if (lvlText) {
    const removed = removeManualNumberingPrefix(pEl, lvlText)
    if (removed) {
      ctx.manualNumberingRemoved.set(
        lvlText,
        (ctx.manualNumberingRemoved.get(lvlText) ?? 0) + 1,
      )
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

function stripConflictingDirectFormatting(
  pEl: Element,
  _styleId: string,
  _ctx: ApplyContext,
) {
  const w = NS.w
  const pPr = firstChildNS(pEl, w, "pPr")
  if (!pPr) return
  // remove direct paragraph formatting that the style now controls
  const conflictNames = new Set([
    "jc",
    "spacing",
    "ind",
    "outlineLvl",
  ])
  for (const c of Array.from(getChildren(pPr))) {
    if (c.namespaceURI === w && conflictNames.has(c.localName!)) {
      pPr.removeChild(c)
    }
  }
  // remove paragraph-mark rPr direct font/size/bold/italic
  const paraRPr = firstChildNS(pPr, w, "rPr")
  if (paraRPr) {
    stripRunRPrConflicts(paraRPr)
    if (getChildren(paraRPr).length === 0) {
      pPr.removeChild(paraRPr)
    }
  }
  // for each run, strip conflicting rPr
  for (const r of getChildrenNS(pEl, w, "r")) {
    const rPr = firstChildNS(r, w, "rPr")
    if (!rPr) continue
    stripRunRPrConflicts(rPr)
    if (getChildren(rPr).length === 0) {
      r.removeChild(rPr)
    }
  }
}

function stripRunRPrConflicts(rPr: Element) {
  const w = NS.w
  const conflictNames = new Set([
    "rFonts",
    "sz",
    "szCs",
    "b",
    "bCs",
    "i",
    "iCs",
    "color",
  ])
  for (const c of Array.from(getChildren(rPr))) {
    if (c.namespaceURI === w && conflictNames.has(c.localName!)) {
      rPr.removeChild(c)
    }
  }
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
      if (def.lineSpacing < 10) {
        spacing.setAttributeNS(w, "w:line", String(Math.round(def.lineSpacing * 240)))
        spacing.setAttributeNS(w, "w:lineRule", "auto")
      } else {
        spacing.setAttributeNS(w, "w:line", String(Math.round(def.lineSpacing * 20)))
        spacing.setAttributeNS(w, "w:lineRule", "exact")
      }
    }
    pPr.appendChild(spacing)
  }
  if (def.firstLineIndent != null || def.hangingIndent != null) {
    const ind = stylesDoc.createElementNS(w, "w:ind")
    if (def.firstLineIndent != null && def.firstLineIndent !== 0) {
      ind.setAttributeNS(w, "w:firstLine", String(parseIndent(def.firstLineIndent)))
    }
    if (def.hangingIndent != null && def.hangingIndent !== 0) {
      ind.setAttributeNS(w, "w:hanging", String(parseIndent(def.hangingIndent)))
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

function parseIndent(v: string | number): number {
  if (typeof v === "number") return Math.round(v * 20) // pt → twips
  const m = v.trim().match(/^(-?\d+(?:\.\d+)?)\s*(char|chars|pt)?$/i)
  if (!m) return 0
  const n = parseFloat(m[1]!)
  const unit = (m[2] || "").toLowerCase()
  if (unit.startsWith("char")) return Math.round(n * 240)
  return Math.round(n * 20)
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
  output: string
}) {
  const lines: string[] = []
  lines.push("=== Change Report ===")
  lines.push(`Styles injected: ${args.injected.length} (${args.injected.join(", ")})`)
  lines.push(`Styles updated:  ${args.updated.length} (${args.updated.join(", ")})`)
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
