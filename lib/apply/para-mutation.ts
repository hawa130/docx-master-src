import { summarizeTable } from "@lib/parse/table-classifier.ts"
import { NS, type ParsedParagraph } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, textContent, wAttr } from "@lib/xml/xml-utils.ts"
import type { ApplyContext, CompiledPatternRule, RestyleSample } from "@lib/config/config-types.ts"

/* ------------- paragraph processing ------------- */

export function applyToBody(documentDoc: Document, ctx: ApplyContext) {
  const root = documentDoc.documentElement
  const body = firstChildNS(root, NS.w, "body")
  if (!body) return
  // Build a map of paragraph index → ParsedParagraph for action lookup
  const byIdx = new Map<number, ParsedParagraph>()
  for (const p of ctx.paragraphs) byIdx.set(p.index, p)

  // Walk in same order as DocumentParser to assign indices; modify in place
  let nextIdx = 1
  const traverseChildren = (parentEl: Element | Document, _insideLayout: boolean) => {
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
        // data tables: leave content untouched
      }
    }
  }
  traverseChildren(body, false)
}

function processOneParagraph(pEl: Element, para: ParsedParagraph, ctx: ApplyContext) {
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
  if (action !== "restyle" || !targetStyle) {
    // Paragraph fell through all rules — implicit keep. Track by fingerprint
    // so the change report can show which fingerprints were not covered;
    // makes "where did the rest go?" verification cheap.
    if (!a) {
      // Skip recording paragraphs the edits[] pass will replace or delete:
      // they won't survive past the edit phase, so showing them in
      // implicit-keep would read as false-positive coverage gaps.
      if (ctx.editTouchedIndices?.has(para.index)) return
      const isEmpty = para.text.trim().length === 0
      const cur = ctx.implicitKeepByFingerprint.get(para.fingerprint) ?? {
        empty: 0,
        nonEmpty: 0,
        nonEmptySamples: [],
      }
      if (isEmpty) cur.empty += 1
      else {
        cur.nonEmpty += 1
        // Keep up to 2 short samples per fingerprint so the agent can spot-
        // check coverage without running inspect_range. Two is enough to
        // confirm a kind ("年 月 日" / "Page 3 of 12" → form chrome) without
        // bloating the report.
        if (cur.nonEmptySamples.length < 2) {
          const snippet = para.text.trim().slice(0, 30)
          cur.nonEmptySamples.push(snippet + (para.text.trim().length > 30 ? "…" : ""))
        }
      }
      ctx.implicitKeepByFingerprint.set(para.fingerprint, cur)
    }
    return
  }

  // apply restyle
  const oldPStyle = para.styleId
  setParagraphStyle(pEl, targetStyle)
  stripConflictingDirectFormatting(
    pEl,
    ctx.stylePPrCascade.get(targetStyle) ?? new Set(),
    ctx.styleRPrCascade.get(targetStyle) ?? new Set(),
  )
  ctx.restyleStats.set(targetStyle, (ctx.restyleStats.get(targetStyle) ?? 0) + 1)

  // Record sample for the change report (cap per style to keep output bounded).
  // `thisSample` refers ONLY to the sample created for THIS paragraph; if the
  // cap was already reached and we didn't push, it stays null. This prevents
  // notes from later paragraphs leaking onto the last sample of the previous
  // bunch — earlier code used existingSamples[length-1] which was wrong for
  // every paragraph past the cap.
  const existingSamples = ctx.samples.get(targetStyle) ?? []
  let thisSample: RestyleSample | null = null
  if (existingSamples.length < ctx.samplesPerStyleCap) {
    const via: RestyleSample["via"] = a ? "assignment" : matchedPattern ? "pattern" : "bulk"
    thisSample = {
      paraIndex: para.index,
      oldStyle: oldPStyle,
      newStyle: targetStyle,
      textPreview: para.text.slice(0, 60) + (para.text.length > 60 ? "…" : ""),
      via,
      patternSource: matchedPattern?.rule.source,
      notes: [],
    }
    existingSamples.push(thisSample)
    ctx.samples.set(targetStyle, existingSamples)
  }

  if (matchedPattern) {
    const key = matchedPattern.rule.source
    ctx.patternMatchStats.set(key, (ctx.patternMatchStats.get(key) ?? 0) + 1)
    if (matchedPattern.rule.stripMatch) {
      const removed = removeRegexPrefix(pEl, matchedPattern.rule.regex)
      if (removed) {
        ctx.patternStripStats.set(key, (ctx.patternStripStats.get(key) ?? 0) + 1)
        if (thisSample) thisSample.notes.push(`stripped pattern /${key}/`)
      }
    }
  }

  const lvlPatterns = ctx.numLvlTextByStyle.get(targetStyle)
  if (lvlPatterns && lvlPatterns.length > 0) {
    let stripped = false
    for (const pat of lvlPatterns) {
      if (removeManualNumberingPrefix(pEl, pat)) {
        ctx.manualNumberingRemoved.set(pat, (ctx.manualNumberingRemoved.get(pat) ?? 0) + 1)
        let perStyle = ctx.manualNumberingByStyle.get(targetStyle)
        if (!perStyle) {
          perStyle = new Map()
          ctx.manualNumberingByStyle.set(targetStyle, perStyle)
        }
        perStyle.set(pat, (perStyle.get(pat) ?? 0) + 1)
        if (thisSample) thisSample.notes.push(`stripped manual prefix "${pat}"`)
        stripped = true
        break
      }
    }
    // None of the level's stripPrefixPatterns covered this paragraph's
    // leading text. Record a leading-text sample so the agent can see what
    // the doc actually contains — recognising whether it's a missed prefix
    // shape ("1. ..." that needs adding to stripPrefixPatterns) or a clean
    // heading without manual numbering (already correct, no action). The
    // script doesn't classify the shape; agents are far better at reading
    // typed text and the script's hardcoded shape list would always lag
    // real-doc variety (附录 A, I., a), 任务一, ⒈, ...).
    //
    // Skip when the paragraph was already stripped by a pattern_rule
    // (matchedPattern.stripMatch): the prefix is gone, so lvlPatterns
    // finding nothing is expected, not a missed shape. Reporting it would
    // double-count and read as a false-positive coverage gap.
    const alreadyStrippedByPatternRule = !!matchedPattern?.rule.stripMatch
    if (!stripped && !alreadyStrippedByPatternRule) {
      let perStyle = ctx.unstrippedByStyle.get(targetStyle)
      if (!perStyle) {
        perStyle = { count: 0, samples: [] }
        ctx.unstrippedByStyle.set(targetStyle, perStyle)
      }
      perStyle.count += 1
      if (perStyle.samples.length < 3) {
        const snippet = para.text.trim().slice(0, 40)
        perStyle.samples.push(snippet + (para.text.trim().length > 40 ? "…" : ""))
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
  "vertAlign",
] as const

function stripConflictingDirectFormatting(
  pEl: Element,
  stylePPrProvides: Set<string>,
  styleRPrProvides: Set<string>,
) {
  const w = NS.w
  const pPr = firstChildNS(pEl, w, "pPr")
  if (pPr) {
    // Strip direct paragraph-level overrides ONLY for properties the new
    // style's cascade actually provides — those are the conflicts the style
    // takes over. Properties the style cascade doesn't declare stay on the
    // paragraph: they came from chrome convention (e.g. template-prescribed
    // `<w:spacing line=400 lineRule=exact/>` baked into every chrome
    // paragraph), and dropping them would silently fall back to docDefaults
    // and lose the convention.
    const directConflicts = new Set(["jc", "spacing", "ind", "outlineLvl"])
    for (const c of Array.from(getChildren(pPr))) {
      if (
        c.namespaceURI === w &&
        directConflicts.has(c.localName!) &&
        stylePPrProvides.has(c.localName!)
      ) {
        pPr.removeChild(c)
      }
    }
    const paraRPr = firstChildNS(pPr, w, "rPr")
    if (paraRPr) {
      removeRPrConflicts(paraRPr, styleRPrProvides)
      if (getChildren(paraRPr).length === 0) pPr.removeChild(paraRPr)
    }
  }

  // Run-level rPr: only strip a property when (a) ALL runs in the paragraph
  // carry the same value for it AND (b) the new style's cascade actually
  // declares that property. (a) is the uniform-vs-mixed rule (mixed-runs
  // formatting like a colored numbering prefix or bold lead phrase is
  // intentional and must survive). (b) is the same gate as the pPr-level
  // strip: stripping a uniform run property the style cascade doesn't
  // declare would fall back to docDefaults / Normal cascade and silently
  // lose chrome's run-level conventions (font, size, bold). styleRPrProvides
  // is the cascade set passed in.
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
    if (vals.size <= 1 && styleRPrProvides.has(name)) uniformToStrip.add(name)
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

function removeRPrConflicts(rPr: Element, styleProvides: Set<string>) {
  const w = NS.w
  const conflicts = new Set<string>(RPR_CONFLICT_NAMES)
  for (const c of Array.from(getChildren(rPr))) {
    if (c.namespaceURI === w && conflicts.has(c.localName!) && styleProvides.has(c.localName!)) {
      rPr.removeChild(c)
    }
  }
}

function rPrChildSignature(el: Element, name: string): string {
  if (name === "rFonts") {
    return ["ascii", "hAnsi", "eastAsia", "cs"].map((a) => `${a}=${wAttr(el, a) ?? ""}`).join("|")
  }
  // toggles (b, bCs, i, iCs): absence==off; presence==on unless val="0"/"false"
  if (name === "b" || name === "bCs" || name === "i" || name === "iCs") {
    const v = wAttr(el, "val")
    return v === "0" || v === "false" ? "off" : "on"
  }
  return wAttr(el, "val") ?? ""
}

/**
 * Replace a leading match in the paragraph's leading text with the regex
 * stripped. Returns true on a hit. Accumulates text from consecutive w:t
 * runs so that prefixes split across runs (e.g. "五" + "、Title", which
 * Word emits when a paragraph was hand-edited) still match. Removed
 * characters are distributed across the runs they came from; subsequent
 * runs are emptied as needed and the first remaining run keeps the tail.
 */
function stripLeadingMatch(pEl: Element, re: RegExp): boolean {
  const w = NS.w
  const tEls: Element[] = []
  let combined = ""
  for (const run of getChildrenNS(pEl, w, "r")) {
    const tEl = firstChildNS(run, w, "t")
    if (!tEl) continue
    tEls.push(tEl)
    combined += textContent(tEl)
    const m = combined.match(re)
    if (!m) continue
    let remaining = m[0].length
    for (const t of tEls) {
      const txt = textContent(t)
      while (t.firstChild) t.removeChild(t.firstChild)
      if (remaining >= txt.length) {
        remaining -= txt.length
        t.appendChild(t.ownerDocument!.createTextNode(""))
      } else {
        t.appendChild(t.ownerDocument!.createTextNode(txt.slice(remaining)))
        remaining = 0
      }
      t.setAttribute("xml:space", "preserve")
    }
    return true
  }
  return false
}

function removeRegexPrefix(pEl: Element, regex: RegExp): boolean {
  const re = regex.source.startsWith("^") ? regex : new RegExp("^" + regex.source, regex.flags)
  return stripLeadingMatch(pEl, re)
}

function removeManualNumberingPrefix(pEl: Element, lvlText: string): boolean {
  // Build a regex from lvlText: replace each %N placeholder with a generic
  // numeric / CJK-numeral capture (matches "1.1", "第三章", etc.).
  const pattern = lvlText
    .replace(/[.*+?^${}()|[\]\\]/g, (m) => "\\" + m)
    .replace(/%\d/g, "(?:\\d+|[一二三四五六七八九十百千]+)")
  return stripLeadingMatch(pEl, new RegExp("^\\s*" + pattern + "\\s*"))
}
