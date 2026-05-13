/**
 * Numbering counter simulator.
 *
 * Walks the post-edit `document.xml` in body order, maintaining per-`numId`
 * per-level counters, and renders each numbered paragraph's `lvlText` with
 * the actual counter values substituted. The output map keys by paragraph
 * Element so callers (REF placeholder backfill) can look up exactly what
 * Word will render once it updates fields on open.
 *
 * Why a simulator: Word renders auto-numbering at view time from `<w:numPr>`
 * + `numbering.xml`. The XML in `document.xml` doesn't contain the rendered
 * digits anywhere — running the counter machinery offline is the only way
 * to compute the placeholder text a REF field should display BEFORE Word
 * updates fields. Without this, the placeholder either stays empty (Word
 * shows blank until F9) or carries an incorrect guess.
 *
 * Limitations:
 *   - Counter reset behavior modelled simply: a higher-level increment
 *     resets every lower level to its `start` value. Matches Word's default
 *     `lvlRestart` (= -1, restart-on-any-higher-level). Custom
 *     `lvlRestart` values are not honored — the placeholder may diverge in
 *     that case (Word still updates correctly on F9).
 *   - `isLgl` is honored: cross-level placeholders rendered as decimal
 *     regardless of the referenced level's `numFmt`.
 *   - Exotic `numFmt` values (e.g. `cardinalText`, `ordinalText`) fall back
 *     to decimal. Common formats covered: decimal, upper/lowerLetter,
 *     upper/lowerRoman, chineseCounting, chineseCountingThousand,
 *     ideographTraditional, bullet, none.
 */

import { NS } from "@lib/parse/types.ts"
import {
  firstChildNS,
  getChildren,
  getChildrenNS,
  wAttr,
  wVal,
  walkBodyParagraphs,
} from "@lib/xml/xml-utils.ts"

export interface RenderedNumbering {
  /** Rendered lvlText with %N placeholders substituted. */
  label: string
  /** Bottom level's own counter value, rendered per its numFmt. */
  number: string
}

interface AbstractLevel {
  level: number
  numFmt: string
  lvlText: string
  start: number
  isLgl: boolean
}

interface AbstractNum {
  levels: Map<number, AbstractLevel>
}

/** Build a `paragraphElement → RenderedNumbering` map for every paragraph
 * carrying `<w:numPr>` — including paragraphs that inherit numbering from
 * a style cascade. Paragraphs without any numbering binding are absent. */
export function simulateNumberingCounters(
  documentDoc: Document,
  numberingDoc: Document | null,
  stylesDoc: Document | null = null,
): Map<Element, RenderedNumbering> {
  const out = new Map<Element, RenderedNumbering>()
  const w = NS.w
  if (!numberingDoc) return out

  const numIdToAbstract = buildNumIdMap(numberingDoc)
  const numIdStartOverrides = buildNumStartOverrides(numberingDoc)
  const abstractById = buildAbstractMap(numberingDoc)
  const styleNumPr = stylesDoc ? buildStyleNumPrMap(stylesDoc) : new Map()

  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return out

  // counters[numId][level] = current counter value. Initialized lazily.
  const counters = new Map<string, Map<number, number>>()
  const ensureCounters = (numId: string) => {
    let m = counters.get(numId)
    if (!m) {
      m = new Map()
      counters.set(numId, m)
    }
    return m
  }

  // Walk all <w:p> in document order via the canonical body walker — same
  // traversal scope as `injectChapterCounters` (chapter-SEQ injection) and
  // `standardizeCaptions` (caption re-emit). Descends through tbl/tr/tc so
  // captions / numbered paragraphs inside tables advance counters here too.
  for (const pEl of walkBodyParagraphs(body)) {
    const binding = resolveParagraphNumbering(pEl, styleNumPr)
    if (!binding) continue
    const { numId, level } = binding
    if (numId === "0") continue

    const abstractId = numIdToAbstract.get(numId)
    if (abstractId === undefined) continue
    const abstract = abstractById.get(abstractId)
    if (!abstract) continue
    const lvlDef = abstract.levels.get(level)
    if (!lvlDef) continue

    // Increment this level's counter; reset every lower level to its start.
    // Per-numId <w:lvlOverride><w:startOverride> wins over the abstract's
    // <w:start> — that's how the perInstance restart pass communicates
    // "fresh counter for this list instance" via fork-and-override.
    const cmap = ensureCounters(numId)
    const overrideForNumId = numIdStartOverrides.get(numId)
    const startFor = (lvl: number, abstractStart: number): number =>
      overrideForNumId?.get(lvl) ?? abstractStart
    const cur = cmap.get(level) ?? startFor(level, lvlDef.start) - 1
    cmap.set(level, cur + 1)
    for (const [lvl, def] of abstract.levels) {
      if (lvl > level) cmap.set(lvl, startFor(lvl, def.start) - 1)
    }

    out.set(pEl, {
      label: renderLvlText(lvlDef, abstract, cmap),
      number: renderCounter(lvlDef.numFmt, cmap.get(level) ?? lvlDef.start),
    })
  }
  return out
}

interface NumBinding {
  numId: string
  level: number
}

/** Resolve a paragraph's effective numbering binding: direct numPr wins;
 * otherwise the paragraph's pStyle is traced through the style cascade
 * (built once in `styleNumPr`) until a numId surfaces. */
function resolveParagraphNumbering(
  pEl: Element,
  styleNumPr: Map<string, NumBinding>,
): NumBinding | null {
  const pPr = firstChildNS(pEl, NS.w, "pPr")
  if (!pPr) return null
  const direct = firstChildNS(pPr, NS.w, "numPr")
  if (direct) {
    const numIdEl = firstChildNS(direct, NS.w, "numId")
    const ilvlEl = firstChildNS(direct, NS.w, "ilvl")
    const numId = numIdEl ? wVal(numIdEl) : null
    if (numId) return { numId, level: ilvlEl ? parseInt(wVal(ilvlEl) || "0", 10) : 0 }
  }
  const pStyle = firstChildNS(pPr, NS.w, "pStyle")
  if (!pStyle) return null
  const styleId = wVal(pStyle)
  if (!styleId) return null
  return styleNumPr.get(styleId) ?? null
}

/** Walk stylesDoc to map every styleId to its cascade-resolved numbering
 * binding (direct numPr or inherited via basedOn). Built once per
 * simulator run; lookups are O(1) thereafter. */
function buildStyleNumPrMap(stylesDoc: Document): Map<string, NumBinding> {
  const w = NS.w
  const out = new Map<string, NumBinding>()
  const root = stylesDoc.documentElement
  if (!root) return out
  const styles = new Map<string, Element>()
  for (const s of getChildrenNS(root, w, "style")) {
    const id = wAttr(s, "styleId")
    if (id) styles.set(id, s)
  }
  const resolve = (id: string, seen: Set<string>): NumBinding | null => {
    if (seen.has(id)) return null
    seen.add(id)
    const el = styles.get(id)
    if (!el) return null
    const pPr = firstChildNS(el, w, "pPr")
    if (pPr) {
      const numPr = firstChildNS(pPr, w, "numPr")
      if (numPr) {
        const numIdEl = firstChildNS(numPr, w, "numId")
        const ilvlEl = firstChildNS(numPr, w, "ilvl")
        const numId = numIdEl ? wVal(numIdEl) : null
        if (numId) return { numId, level: ilvlEl ? parseInt(wVal(ilvlEl) || "0", 10) : 0 }
      }
    }
    const basedOn = firstChildNS(el, w, "basedOn")
    if (basedOn) {
      const parent = wVal(basedOn)
      if (parent) return resolve(parent, seen)
    }
    return null
  }
  for (const id of styles.keys()) {
    const r = resolve(id, new Set())
    if (r) out.set(id, r)
  }
  return out
}

/** Read `<w:lvlOverride><w:startOverride>` on every `<w:num>` to produce
 * `numId → level → startValue`. Used by the simulator so a forked numId
 * (perInstance restart) starts counting from its override instead of the
 * abstractNum's `<w:start>`. */
function buildNumStartOverrides(numberingDoc: Document): Map<string, Map<number, number>> {
  const out = new Map<string, Map<number, number>>()
  const w = NS.w
  for (const num of getChildrenNS(numberingDoc.documentElement, w, "num")) {
    const numId = wAttr(num, "numId")
    if (!numId) continue
    const perLevel = new Map<number, number>()
    for (const ovr of getChildrenNS(num, w, "lvlOverride")) {
      const ilvl = wAttr(ovr, "ilvl")
      if (ilvl === null) continue
      const startOvr = firstChildNS(ovr, w, "startOverride")
      if (!startOvr) continue
      const val = wAttr(startOvr, "val")
      if (val === null) continue
      perLevel.set(parseInt(ilvl, 10), parseInt(val, 10))
    }
    if (perLevel.size > 0) out.set(numId, perLevel)
  }
  return out
}

function buildNumIdMap(numberingDoc: Document): Map<string, string> {
  const out = new Map<string, string>()
  const w = NS.w
  for (const num of getChildrenNS(numberingDoc.documentElement, w, "num")) {
    const numId = wAttr(num, "numId")
    if (!numId) continue
    const absRef = firstChildNS(num, w, "abstractNumId")
    if (!absRef) continue
    const absId = wAttr(absRef, "val") ?? wVal(absRef)
    if (absId) out.set(numId, absId)
  }
  return out
}

function buildAbstractMap(numberingDoc: Document): Map<string, AbstractNum> {
  const out = new Map<string, AbstractNum>()
  const w = NS.w
  for (const abs of getChildrenNS(numberingDoc.documentElement, w, "abstractNum")) {
    const id = wAttr(abs, "abstractNumId")
    if (!id) continue
    const levels = new Map<number, AbstractLevel>()
    for (const lvl of getChildrenNS(abs, w, "lvl")) {
      const ilvl = wAttr(lvl, "ilvl")
      if (ilvl === null) continue
      const level = parseInt(ilvl, 10)
      const numFmtEl = firstChildNS(lvl, w, "numFmt")
      const lvlTextEl = firstChildNS(lvl, w, "lvlText")
      const startEl = firstChildNS(lvl, w, "start")
      const isLglEl = firstChildNS(lvl, w, "isLgl")
      levels.set(level, {
        level,
        numFmt: (numFmtEl && wAttr(numFmtEl, "val")) || "decimal",
        lvlText: (lvlTextEl && wAttr(lvlTextEl, "val")) || "",
        start: parseInt((startEl && wAttr(startEl, "val")) || "1", 10),
        isLgl: !!isLglEl,
      })
    }
    out.set(id, { levels })
  }
  return out
}

function renderLvlText(
  lvl: AbstractLevel,
  abstract: AbstractNum,
  counters: Map<number, number>,
): string {
  if (lvl.numFmt === "bullet" || lvl.numFmt === "none") return lvl.lvlText
  return lvl.lvlText.replace(/%(\d)/g, (_match, digit: string) => {
    const refLevel = parseInt(digit, 10) - 1
    const counter = counters.get(refLevel)
    if (counter === undefined) return ""
    const refLvlDef = abstract.levels.get(refLevel)
    const fmt = lvl.isLgl ? "decimal" : (refLvlDef?.numFmt ?? "decimal")
    return renderCounter(fmt, counter)
  })
}

function renderCounter(numFmt: string, value: number): string {
  switch (numFmt) {
    case "decimal":
      return String(value)
    case "decimalZero":
      return value < 10 ? `0${value}` : String(value)
    case "upperLetter":
      return toAlphaCounter(value, true)
    case "lowerLetter":
      return toAlphaCounter(value, false)
    case "upperRoman":
      return toRoman(value).toUpperCase()
    case "lowerRoman":
      return toRoman(value).toLowerCase()
    case "chineseCounting":
    case "chineseCountingThousand":
      return toChineseCounting(value)
    case "ideographTraditional":
      return toHeavenlyStem(value)
    case "bullet":
    case "none":
      return ""
    default:
      // Unknown numFmts: fall back to decimal. Word will re-render on F9.
      return String(value)
  }
}

function toAlphaCounter(n: number, upper: boolean): string {
  // Spreadsheet-column style: A..Z, AA..ZZ, AAA...
  let s = ""
  let m = n
  while (m > 0) {
    const r = (m - 1) % 26
    s = String.fromCharCode((upper ? 65 : 97) + r) + s
    m = Math.floor((m - 1) / 26)
  }
  return s || (upper ? "A" : "a")
}

function toRoman(n: number): string {
  if (n <= 0) return ""
  const pairs: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ]
  let out = ""
  let m = n
  for (const [v, r] of pairs) {
    while (m >= v) {
      out += r
      m -= v
    }
  }
  return out
}

const CHINESE_DIGITS = ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
function toChineseCounting(n: number): string {
  if (n === 0) return "〇"
  if (n < 10) return CHINESE_DIGITS[n]!
  if (n < 20) return n === 10 ? "十" : `十${CHINESE_DIGITS[n - 10]!}`
  if (n < 100) {
    const tens = Math.floor(n / 10)
    const ones = n % 10
    return `${CHINESE_DIGITS[tens]!}十${ones === 0 ? "" : CHINESE_DIGITS[ones]!}`
  }
  // > 99: fall back to digit-by-digit. Real Word uses 百 / 千 forms; this
  // is the simplification covering the realistic counter range. Word
  // re-renders on F9 so the placeholder mismatch is cosmetic.
  return String(n)
    .split("")
    .map((d) => CHINESE_DIGITS[parseInt(d, 10)]!)
    .join("")
}

const HEAVENLY_STEMS = ["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"]
function toHeavenlyStem(n: number): string {
  if (n < 1) return ""
  return HEAVENLY_STEMS[(n - 1) % 10]!
}

/** Extract the visible text from a paragraph's `<w:t>` descendants — used
 * for the "full" display option (REF without switches renders the
 * bookmark's text content). Skips text inside fldChar=begin/end pairs to
 * avoid recursive reference text leaking in. */
export function extractParagraphText(pEl: Element): string {
  const w = NS.w
  const parts: string[] = []
  let inField = 0
  const walk = (el: Element) => {
    for (const c of getChildren(el)) {
      if (c.namespaceURI === w && c.localName === "fldChar") {
        const type = wAttr(c, "fldCharType")
        if (type === "begin") inField++
        else if (type === "end" && inField > 0) inField--
        continue
      }
      if (c.namespaceURI === w && c.localName === "t" && inField === 0) {
        parts.push(c.textContent || "")
        continue
      }
      walk(c)
    }
  }
  walk(pEl)
  return parts.join("")
}
