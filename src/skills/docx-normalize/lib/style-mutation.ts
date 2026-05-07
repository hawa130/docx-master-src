import { NS, type ParsedParagraph } from "@core/types.ts"
import {
  firstChildNS,
  getChildren,
  getChildrenNS,
  wAttr,
} from "@core/xml-utils.ts"
import type { StyleConfigEntry } from "./types.ts"

/* ------------- fromParagraph resolution ------------- */

export function resolveStyleDef(
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

/** pPr / rPr children that this function manages (writes from def). Anything
 * else found in an existing style's pPr/rPr is preserved untouched. The pPr
 * list is critical: numPr (auto-numbering binding), keepNext, pBdr, shd,
 * adjustRightInd, etc. are all preserved when overriding an existing style. */
const PPR_MANAGED_CHILDREN = new Set(["spacing", "ind", "jc", "outlineLvl"])
const RPR_MANAGED_CHILDREN = new Set([
  "rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color",
])

export function upsertStyle(stylesDoc: Document, def: StyleConfigEntry): "created" | "updated" {
  const w = NS.w
  const root = stylesDoc.documentElement!
  const existing = getChildrenNS(root, w, "style").find((s) => wAttr(s, "styleId") === def.id)
  let target: Element
  let result: "created" | "updated"
  if (existing) {
    target = existing
    result = "updated"
  } else {
    target = stylesDoc.createElementNS(w, "w:style")
    target.setAttributeNS(w, "w:type", "paragraph")
    target.setAttributeNS(w, "w:styleId", def.id)
    root.appendChild(target)
    result = "created"
  }

  // name: required and idempotent. When the element already exists, update
  // its w:val in place — otherwise removing-and-re-appending would push it
  // to the end of <w:style>, but OOXML's schema (ECMA-376 §17.7.4) requires
  // <w:name> to be the FIRST child. Word is lenient enough to load
  // mis-ordered styles, but stricter validators (and other docx libraries)
  // reject them.
  let nameEl = firstChildNS(target, w, "name")
  if (nameEl) {
    nameEl.setAttributeNS(w, "w:val", def.name)
  } else {
    nameEl = stylesDoc.createElementNS(w, "w:name")
    nameEl.setAttributeNS(w, "w:val", def.name)
    target.insertBefore(nameEl, target.firstChild)
  }

  // basedOn: only touched when def explicitly provides it; otherwise the
  // existing style's basedOn (and inheritance chain) is preserved. Avoids
  // silently flattening the cascade when an agent overrides a style without
  // re-specifying its parent. Same in-place update pattern as name to keep
  // the element in its original DOM position (basedOn must come before pPr
  // per the OOXML schema).
  if (def.basedOn) {
    let bo = firstChildNS(target, w, "basedOn")
    if (bo) {
      bo.setAttributeNS(w, "w:val", def.basedOn)
    } else {
      bo = stylesDoc.createElementNS(w, "w:basedOn")
      bo.setAttributeNS(w, "w:val", def.basedOn)
      // OOXML: basedOn comes right after name. Insert there.
      const afterName = nameEl.nextSibling
      if (afterName) target.insertBefore(bo, afterName)
      else target.appendChild(bo)
    }
  }

  // pPr: mutate in place. Remove only the children listed in
  // PPR_MANAGED_CHILDREN (the visible paragraph properties this function
  // writes), then append the new ones built from `def`. Existing children we
  // don't manage — most importantly `numPr` (numbering binding), but also
  // keepNext, pBdr, shd, adjustRightInd, etc. — stay intact. Without this,
  // overriding an existing heading style would silently drop its
  // auto-numbering reference.
  let pPr = firstChildNS(target, w, "pPr")
  if (pPr) {
    for (const c of Array.from(getChildren(pPr))) {
      if (c.namespaceURI === w && PPR_MANAGED_CHILDREN.has(c.localName!)) {
        pPr.removeChild(c)
      }
    }
  }
  const pPrAdditions: Element[] = []
  if (def.outlineLevel !== undefined) {
    const ol = stylesDoc.createElementNS(w, "w:outlineLvl")
    ol.setAttributeNS(w, "w:val", String(def.outlineLevel))
    pPrAdditions.push(ol)
  }
  if (def.alignment) {
    const jc = stylesDoc.createElementNS(w, "w:jc")
    jc.setAttributeNS(w, "w:val", def.alignment)
    pPrAdditions.push(jc)
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
    pPrAdditions.push(spacing)
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
    pPrAdditions.push(ind)
  }
  if (pPrAdditions.length > 0) {
    if (!pPr) {
      pPr = stylesDoc.createElementNS(w, "w:pPr")
      target.appendChild(pPr)
    }
    for (const c of pPrAdditions) pPr.appendChild(c)
  }

  // rPr: same mutate-in-place pattern. Removes only the run properties this
  // function manages (font, size, weight, italic, color); preserves anything
  // else the existing rPr carried (lang, w, kern, etc.).
  let rPr = firstChildNS(target, w, "rPr")
  if (rPr) {
    for (const c of Array.from(getChildren(rPr))) {
      if (c.namespaceURI === w && RPR_MANAGED_CHILDREN.has(c.localName!)) {
        rPr.removeChild(c)
      }
    }
  }
  const rPrAdditions: Element[] = []
  if (def.font || def.fontEastAsia) {
    const rFonts = stylesDoc.createElementNS(w, "w:rFonts")
    const ascii = def.font ?? def.fontEastAsia ?? ""
    const ea = def.fontEastAsia ?? def.font ?? ""
    if (ascii) {
      rFonts.setAttributeNS(w, "w:ascii", ascii)
      rFonts.setAttributeNS(w, "w:hAnsi", ascii)
    }
    if (ea) rFonts.setAttributeNS(w, "w:eastAsia", ea)
    rPrAdditions.push(rFonts)
  }
  if (def.size !== undefined) {
    const sz = stylesDoc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(Math.round(def.size * 2)))
    rPrAdditions.push(sz)
    const szCs = stylesDoc.createElementNS(w, "w:szCs")
    szCs.setAttributeNS(w, "w:val", String(Math.round(def.size * 2)))
    rPrAdditions.push(szCs)
  }
  if (def.bold) {
    rPrAdditions.push(stylesDoc.createElementNS(w, "w:b"))
    rPrAdditions.push(stylesDoc.createElementNS(w, "w:bCs"))
  }
  if (def.italic) {
    rPrAdditions.push(stylesDoc.createElementNS(w, "w:i"))
    rPrAdditions.push(stylesDoc.createElementNS(w, "w:iCs"))
  }
  if (def.color) {
    const color = stylesDoc.createElementNS(w, "w:color")
    color.setAttributeNS(w, "w:val", def.color)
    rPrAdditions.push(color)
  }
  if (rPrAdditions.length > 0) {
    if (!rPr) {
      rPr = stylesDoc.createElementNS(w, "w:rPr")
      target.appendChild(rPr)
    }
    for (const c of rPrAdditions) rPr.appendChild(c)
  }

  return result
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

/**
 * Move the styles named in `orderedIds` to the top of styles.xml's <w:style>
 * entry list, preserving the given order. Styles not in the list keep their
 * relative positions. Non-style children of the root (<w:docDefaults>,
 * <w:latentStyles>) stay above all <w:style> entries — they're not touched.
 *
 * Each id appears at most once in `orderedIds`; duplicates are deduped (a
 * style can be both template-imported and re-declared in config.styles[]).
 */
export function reorderAgentTouchedStylesFirst(
  stylesDoc: Document,
  orderedIds: string[],
): void {
  const w = NS.w
  const root = stylesDoc.documentElement
  if (!root) return

  const seen = new Set<string>()
  const dedupedIds = orderedIds.filter((id) => {
    if (seen.has(id)) return false
    seen.add(id)
    return true
  })
  const targetSet = new Set(dedupedIds)

  // Find the touched style elements and the first non-touched <w:style> (our
  // insertion anchor). Iterate once.
  const touchedById = new Map<string, Element>()
  let firstUntouched: Element | null = null
  for (const el of getChildrenNS(root, w, "style")) {
    const id = wAttr(el, "styleId")
    if (id && targetSet.has(id)) {
      touchedById.set(id, el)
    } else if (!firstUntouched) {
      firstUntouched = el
    }
  }

  // Detach touched elements (they'll be re-inserted in `dedupedIds` order).
  for (const el of touchedById.values()) {
    if (el.parentNode === root) root.removeChild(el)
  }

  // Re-insert in the requested order, just before the first untouched style.
  // If the doc has no other styles, append at the end of root (which sits
  // after docDefaults / latentStyles per OOXML schema).
  for (const id of dedupedIds) {
    const el = touchedById.get(id)
    if (!el) continue
    if (firstUntouched) {
      root.insertBefore(el, firstUntouched)
    } else {
      root.appendChild(el)
    }
  }
}
