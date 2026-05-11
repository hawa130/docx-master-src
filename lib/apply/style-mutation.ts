import { NS, type ParsedParagraph } from "@lib/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml-utils.ts"
import type { StyleConfigEntry } from "./config-types.ts"
import type { StyleResolver } from "./style-resolver.ts"
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, insertChildInOrder } from "./xml-order.ts"

/**
 * Insert a freshly-created `<w:pPr>` into a `<w:style>` at the schema-correct
 * position. CT_Style requires `pPr` before `rPr` / `tblPr` / `trPr` / `tcPr` /
 * `tblStylePr`. A naive `appendChild` would land it after an existing `rPr`,
 * which Word's strict validator rejects with "file needs repair".
 */
export function insertPPrIntoStyle(styleEl: Element, pPr: Element): void {
  const w = NS.w
  for (const name of ["rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]) {
    const el = firstChildNS(styleEl, w, name)
    if (el) {
      styleEl.insertBefore(pPr, el)
      return
    }
  }
  styleEl.appendChild(pPr)
}

/* ------------- fromParagraph resolution ------------- */

export function resolveStyleDef(
  def: StyleConfigEntry,
  paragraphs: ParsedParagraph[],
): StyleConfigEntry {
  // Mode B (no fromParagraph): the schema accepts an `overrides` block on
  // every style entry, but downstream emission reads top-level fields only —
  // un-merged overrides drop silently. Spreading them up keeps Mode A and
  // Mode B symmetric so agents can place fields in either location.
  if (def.fromParagraph === undefined) {
    return def.overrides ? { ...def, ...def.overrides } : def
  }
  const para = paragraphs.find((p) => p.index === def.fromParagraph)
  if (!para) {
    const indices = paragraphs.map((p) => p.index)
    const minIdx = indices[0] ?? 0
    const maxIdx = indices[indices.length - 1] ?? 0
    const closest = paragraphs.reduce(
      (best, p) =>
        Math.abs(p.index - def.fromParagraph!) < Math.abs(best.index - def.fromParagraph!)
          ? p
          : best,
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

/**
 * Capture the cascade-resolved display fields for an existing styleId BEFORE
 * this apply run mutates it. Returns null when the styleId is absent from
 * source — the agent is installing it fresh, so there's no prior state to
 * diff against. Shape mirrors `extractDisplayFields` (report.ts) so the two
 * can be compared field-by-field for the dry-run Δ line.
 */
export function extractPriorDisplayFields(
  resolver: StyleResolver,
  styleId: string,
): Record<string, unknown> | null {
  if (!resolver.getStyleDefinition(styleId)) return null
  const { rPr, pPr } = resolver.resolveStyleChain(styleId)
  const out: Record<string, unknown> = {}
  if (rPr.fontEastAsia) out.fontCJK = rPr.fontEastAsia
  const latin = rPr.fontAscii ?? rPr.fontHAnsi
  if (latin) out.fontLatin = latin
  if (rPr.size !== undefined) out.size = rPr.size / 2
  if (rPr.bold !== undefined) out.bold = rPr.bold
  if (rPr.italic !== undefined) out.italic = rPr.italic
  if (rPr.color && rPr.color !== "auto") out.color = rPr.color
  if (pPr.alignment) out.alignment = pPr.alignment
  if (pPr.spaceBefore !== undefined) out.spaceBefore = pPr.spaceBefore / 20
  if (pPr.spaceAfter !== undefined) out.spaceAfter = pPr.spaceAfter / 20
  if (pPr.lineSpacing !== undefined) {
    const rule = pPr.lineRule || "auto"
    if (rule === "auto") {
      out.lineSpacing = pPr.lineSpacing / 240
    } else {
      out.lineSpacing = pPr.lineSpacing / 20
      out.lineRule = rule
    }
  }
  if (pPr.firstLineIndentChars !== undefined) {
    out.firstLineIndent = `${pPr.firstLineIndentChars / 100}char`
  } else if (pPr.firstLineIndent !== undefined) {
    out.firstLineIndent = `${pPr.firstLineIndent / 20}pt`
  }
  if (pPr.outlineLevel !== undefined) out.outlineLevel = pPr.outlineLevel
  return out
}

function paragraphToStyleEntry(p: ParsedParagraph): Partial<StyleConfigEntry> {
  const r = p.rPr
  const pp = p.pPr
  const out: Partial<StyleConfigEntry> = {}
  const latin = r.fontAscii ?? r.fontHAnsi
  if (latin) out.fontLatin = latin
  if (r.fontEastAsia && r.fontEastAsia !== latin) out.fontCJK = r.fontEastAsia
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

/**
 * Resolve a style cross-reference (basedOn / link / next val) to an existing
 * `w:styleId`. ECMA-376 says these refs must target a styleId; agents and
 * defaults often write the conventional name ("Normal", "heading 1") instead.
 * Falls back to case-insensitive `w:name` match when the literal val isn't a
 * styleId. Returns null when neither matches (caller drops the ref rather
 * than leaving it dangling).
 *
 * `selfId` is the styleId of the style currently being written; self-refs
 * (`basedOn === id`) collapse the cascade and are silently dropped.
 */
function resolveStyleRef(stylesDoc: Document, val: string, selfId?: string): string | null {
  const w = NS.w
  const styles = getChildrenNS(stylesDoc.documentElement!, w, "style")
  for (const s of styles) {
    if (wAttr(s, "styleId") === val) {
      return val === selfId ? null : val
    }
  }
  const target = val.toLowerCase()
  for (const s of styles) {
    const id = wAttr(s, "styleId")
    if (id === selfId) continue
    const nameEl = firstChildNS(s, w, "name")
    if (nameEl && (wAttr(nameEl, "val") || "").toLowerCase() === target) {
      return id ?? null
    }
  }
  return null
}

/** pPr / rPr children that this function clears before re-writing from `def`.
 * Anything else found in an existing style's pPr/rPr is preserved untouched.
 * The pPr list is critical: numPr (auto-numbering binding), keepNext, pBdr,
 * shd, adjustRightInd, etc. are all preserved when overriding an existing
 * style.
 *
 * Scope: only the children writeable from `StyleConfigEntry`. Narrower than
 * fragment-emit's RPR_MANAGED_LOCAL_NAMES (run-level rPr from `RunFormat`,
 * which has u / strike). The two sets must NOT be unified — expanding this
 * one without adding the corresponding write paths would silently drop
 * existing style attrs (e.g. <w:u>) that the engine has no way to put back. */
const PPR_MANAGED_CHILDREN = new Set(["spacing", "ind", "jc", "outlineLvl"])
const RPR_MANAGED_CHILDREN = new Set(["rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color"])

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
  //
  // The val MUST be the target style's `w:styleId`, not its `w:name`. In
  // Chinese-localized docs the Normal style typically has styleId="a" with
  // name="Normal" — emitting `basedOn val="Normal"` then dangles to a
  // non-existent id and Word prompts to repair. resolveStyleRef rewrites
  // a name match to the matching styleId; if neither styleId nor name
  // resolves, the basedOn is dropped rather than left dangling.
  if (def.basedOn) {
    const resolved = resolveStyleRef(stylesDoc, def.basedOn, def.id)
    let bo = firstChildNS(target, w, "basedOn")
    if (resolved) {
      if (bo) {
        bo.setAttributeNS(w, "w:val", resolved)
      } else {
        bo = stylesDoc.createElementNS(w, "w:basedOn")
        bo.setAttributeNS(w, "w:val", resolved)
        const afterName = nameEl.nextSibling
        if (afterName) target.insertBefore(bo, afterName)
        else target.appendChild(bo)
      }
    } else if (bo) {
      target.removeChild(bo)
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
      const ls = parseLineSpacing(def.lineSpacing)
      const rule = def.lineRule ?? (ls.explicitPt || ls.value >= 10 ? "exact" : "auto")
      if (rule === "auto") {
        spacing.setAttributeNS(w, "w:line", String(Math.round(ls.value * 240)))
        spacing.setAttributeNS(w, "w:lineRule", "auto")
      } else {
        spacing.setAttributeNS(w, "w:line", String(Math.round(ls.value * 20)))
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
      insertPPrIntoStyle(target, pPr)
    }
    // Build order is local convenience (outlineLevel → jc → spacing → ind);
    // CT_PPr's actual schema order is spacing → ind → jc → outlineLvl. Insert
    // each at its schema-correct position rather than appending blind.
    for (const c of pPrAdditions) insertChildInOrder(pPr, c, PPR_CHILD_ORDER)
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
  if (def.fontLatin || def.fontCJK) {
    const rFonts = stylesDoc.createElementNS(w, "w:rFonts")
    const ascii = def.fontLatin ?? def.fontCJK ?? ""
    const ea = def.fontCJK ?? def.fontLatin ?? ""
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
    for (const c of rPrAdditions) insertChildInOrder(rPr, c, RPR_CHILD_ORDER)
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

/** Normalize a lineSpacing config value to `{ value, explicitPt }`.
 *  - number stays as-is; explicitPt=false (downstream applies the legacy
 *    `<10 → multiplier`, `>=10 → pt` magnitude heuristic).
 *  - "Npt" string (e.g. "20pt") parses to N with explicitPt=true so the
 *    downstream rule defaults to "exact" regardless of magnitude — agent
 *    who explicitly typed "pt" gets pt even at small numbers like 1.5pt.
 *  Throws on unparseable strings so the error fires at config-read time. */
export function parseLineSpacing(v: string | number): { value: number; explicitPt: boolean } {
  if (typeof v === "number") return { value: v, explicitPt: false }
  const m = v.trim().match(/^(-?\d+(?:\.\d+)?)\s*pt$/i)
  if (!m) {
    throw new Error(
      `lineSpacing "${v}": must be a number (e.g. 1.5 multiplier / 20 pt) or "Npt" string (e.g. "20pt").`,
    )
  }
  return { value: parseFloat(m[1]!), explicitPt: true }
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
export function reorderAgentTouchedStylesFirst(stylesDoc: Document, orderedIds: string[]): void {
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
