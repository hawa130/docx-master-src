import {
  NS,
  type ComputedParaStyle,
  type ComputedRunStyle,
  type ParsedParagraph,
} from "@lib/parse/types.ts"
import { firstChildNS, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import type { StyleConfigEntry } from "@lib/config/config-types.ts"
import type { StyleResolver } from "@lib/parse/style-resolver.ts"
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, insertChildInOrder } from "@lib/xml/xml-order.ts"
import {
  charUnitsToString,
  halfPtToPt,
  type LineSpacingInput,
  lineSpacingToConfig,
  parseIndent,
  parseLineSpacing,
  toHalfPt,
  toTwips,
  twipsToPtString,
} from "@lib/shared/units.ts"

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
  // Top-level fields and `overrides` are spread into the same merged shape
  // in both modes. Mode B (no fromParagraph): top-level + overrides on top
  // of nothing. Mode A (fromParagraph): extracted typography from the donor
  // paragraph, then top-level, then overrides (each layer wins over the
  // previous). Keeps the schema symmetric — an agent placing `outlineLevel`
  // at top level in Mode A used to silently drop; now it merges.
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
  // Top-level keys that aren't typography fields (id / name / fromParagraph /
  // basedOn / overrides) are skipped from the spread — they're handled
  // explicitly below or are non-mergeable. Everything else (outlineLevel,
  // size, alignment, ...) layers on top of extracted.
  // Underscore-prefixed names tell tsc + oxlint these destructured fields
  // are intentionally discarded; only `topLevel` is read below.
  const {
    id: _id,
    name: _name,
    fromParagraph: _fp,
    basedOn: _bo,
    overrides: _ov,
    ...topLevel
  } = def
  return {
    basedOn: "Normal",
    ...extracted,
    ...topLevel,
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
  if (rPr.size !== undefined) out.size = halfPtToPt(rPr.size)
  if (rPr.bold !== undefined) out.bold = rPr.bold
  if (rPr.italic !== undefined) out.italic = rPr.italic
  if (rPr.color && rPr.color !== "auto") out.color = rPr.color
  if (rPr.vertAlign) out.vertAlign = rPr.vertAlign
  if (pPr.alignment) out.alignment = pPr.alignment
  if (pPr.spaceBefore !== undefined) out.spaceBefore = twipsToPtString(pPr.spaceBefore)
  if (pPr.spaceAfter !== undefined) out.spaceAfter = twipsToPtString(pPr.spaceAfter)
  if (pPr.lineSpacing !== undefined) {
    out.lineSpacing = lineSpacingToConfig({
      mode: pPrLineRuleToMode(pPr.lineRule),
      value: pPr.lineSpacing,
    })
  }
  if (pPr.firstLineIndentChars !== undefined) {
    out.firstLineIndent = charUnitsToString(pPr.firstLineIndentChars)
  } else if (pPr.firstLineIndent !== undefined) {
    out.firstLineIndent = twipsToPtString(pPr.firstLineIndent)
  }
  if (pPr.outlineLevel !== undefined) out.outlineLevel = pPr.outlineLevel
  return out
}

/** Resolve the OOXML `lineRule` attribute string (or missing) to the
 *  units.ts LineSpacingParsed mode. Defaults to "auto" matching Word's
 *  ECMA-376 default when the attribute is absent. */
function pPrLineRuleToMode(lineRule: string | undefined): "auto" | "exact" | "atLeast" {
  if (lineRule === "exact" || lineRule === "atLeast") return lineRule
  return "auto"
}

function paragraphToStyleEntry(p: ParsedParagraph): Partial<StyleConfigEntry> {
  return computedStyleToEntry(p.rPr, p.pPr)
}

/** Convert a ComputedRunStyle + ComputedParaStyle pair into the field
 * shape used by `StyleConfigEntry`. Shared between Mode A's `fromParagraph`
 * extraction (cascade-merged values) and the dry-run vs-target-direct
 * classifier (per-paragraph direct values). Both produce the same field
 * names + units so they can be diffed against an agent-declared style. */
export function computedStyleToEntry(
  r: ComputedRunStyle,
  pp: ComputedParaStyle,
): Partial<StyleConfigEntry> {
  const out: Partial<StyleConfigEntry> = {}
  const latin = r.fontAscii ?? r.fontHAnsi
  if (latin) out.fontLatin = latin
  if (r.fontEastAsia && r.fontEastAsia !== latin) out.fontCJK = r.fontEastAsia
  if (r.size !== undefined) out.size = halfPtToPt(r.size)
  if (r.bold) out.bold = true
  if (r.italic) out.italic = true
  if (r.color && r.color !== "auto") out.color = r.color
  if (r.vertAlign) out.vertAlign = r.vertAlign
  if (pp.alignment) out.alignment = pp.alignment as StyleConfigEntry["alignment"]
  if (pp.spaceBefore !== undefined) out.spaceBefore = twipsToPtString(pp.spaceBefore)
  if (pp.spaceAfter !== undefined) out.spaceAfter = twipsToPtString(pp.spaceAfter)
  if (pp.lineSpacing !== undefined) {
    out.lineSpacing = lineSpacingToConfig({
      mode: pPrLineRuleToMode(pp.lineRule),
      value: pp.lineSpacing,
    }) as StyleConfigEntry["lineSpacing"]
  }
  // Preserve character-based indent semantics when the source paragraph used
  // `firstLineChars` / `hangingChars` (Word writes these for "首行缩进 N 字符").
  // Round-tripping through pt would drop the font-size auto-scale.
  if (pp.firstLineIndentChars !== undefined) {
    out.firstLineIndent = charUnitsToString(pp.firstLineIndentChars)
  } else if (pp.firstLineIndent !== undefined) {
    out.firstLineIndent = twipsToPtString(pp.firstLineIndent)
  }
  if (pp.hangingIndentChars !== undefined) {
    out.hangingIndent = charUnitsToString(pp.hangingIndentChars)
  } else if (pp.hangingIndent !== undefined) {
    out.hangingIndent = twipsToPtString(pp.hangingIndent)
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

/** Return (or create) a child element with the given local name under
 * `parent`. When creating, the new element is NOT inserted — caller must
 * place it via `insertChildInOrder` or equivalent so schema order is
 * respected. When the element already exists it is returned as-is for
 * attribute-level mutation. */
function getOrCreateNS(parent: Element, doc: Document, ns: string, localName: string): Element {
  const existing = firstChildNS(parent, ns, localName)
  if (existing) return existing
  return doc.createElementNS(ns, `w:${localName}`)
}

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

  // pPr: field-merge in place. Only touch the children the def explicitly
  // declares. Unmanaged children (numPr, keepNext, pBdr, shd, adjustRightInd,
  // etc.) and managed children the def omits stay intact.
  //
  // Compound elements (spacing, ind) are get-or-created and only the
  // attributes for declared fields are set — partial declarations (e.g. only
  // spaceBefore) preserve the element's other attributes intact.
  const hasPPrChange =
    def.outlineLevel !== undefined ||
    def.alignment !== undefined ||
    def.spaceBefore !== undefined ||
    def.spaceAfter !== undefined ||
    def.lineSpacing !== undefined ||
    def.firstLineIndent != null ||
    def.hangingIndent != null
  let pPr = firstChildNS(target, w, "pPr")
  if (hasPPrChange) {
    if (!pPr) {
      pPr = stylesDoc.createElementNS(w, "w:pPr")
      insertPPrIntoStyle(target, pPr)
    }
    if (def.outlineLevel !== undefined) {
      const ol = getOrCreateNS(pPr, stylesDoc, w, "outlineLvl")
      ol.setAttributeNS(w, "w:val", String(def.outlineLevel))
      if (!ol.parentNode) insertChildInOrder(pPr, ol, PPR_CHILD_ORDER)
    }
    if (def.alignment !== undefined) {
      const jc = getOrCreateNS(pPr, stylesDoc, w, "jc")
      jc.setAttributeNS(w, "w:val", def.alignment)
      if (!jc.parentNode) insertChildInOrder(pPr, jc, PPR_CHILD_ORDER)
    }
    if (
      def.spaceBefore !== undefined ||
      def.spaceAfter !== undefined ||
      def.lineSpacing !== undefined
    ) {
      const spacing = getOrCreateNS(pPr, stylesDoc, w, "spacing")
      if (def.spaceBefore !== undefined)
        spacing.setAttributeNS(w, "w:before", String(toTwips(def.spaceBefore, "spaceBefore")))
      if (def.spaceAfter !== undefined)
        spacing.setAttributeNS(w, "w:after", String(toTwips(def.spaceAfter, "spaceAfter")))
      if (def.lineSpacing !== undefined) {
        const ls = parseLineSpacing(def.lineSpacing as LineSpacingInput, "lineSpacing")
        spacing.setAttributeNS(w, "w:line", String(ls.value))
        spacing.setAttributeNS(w, "w:lineRule", ls.mode)
      }
      if (!spacing.parentNode) insertChildInOrder(pPr, spacing, PPR_CHILD_ORDER)
    }
    if (def.firstLineIndent != null || def.hangingIndent != null) {
      const ind = getOrCreateNS(pPr, stylesDoc, w, "ind")
      if (def.firstLineIndent != null && def.firstLineIndent !== 0) {
        const r = parseIndent(def.firstLineIndent)
        if (r) {
          if (r.kind === "char") {
            ind.setAttributeNS(w, "w:firstLineChars", String(r.value))
          } else {
            ind.setAttributeNS(w, "w:firstLine", String(r.value))
          }
        }
      }
      if (def.hangingIndent != null && def.hangingIndent !== 0) {
        const r = parseIndent(def.hangingIndent)
        if (r) {
          if (r.kind === "char") {
            ind.setAttributeNS(w, "w:hangingChars", String(r.value))
          } else {
            ind.setAttributeNS(w, "w:hanging", String(r.value))
          }
        }
      }
      if (!ind.parentNode) insertChildInOrder(pPr, ind, PPR_CHILD_ORDER)
    }
  }

  // rPr: same field-merge pattern. Only touch the attributes / elements the
  // def explicitly declares; anything else the existing rPr carries stays.
  //
  // rFonts: get-or-create the element, then set only the font-slot attributes
  // for the declared fields. fontLatin → ascii + hAnsi; fontCJK → eastAsia.
  // When only one is declared the other slot is left alone so an existing
  // CJK or Latin font in the source style is not clobbered.
  const hasRPrChange =
    def.fontLatin !== undefined ||
    def.fontCJK !== undefined ||
    def.size !== undefined ||
    def.bold !== undefined ||
    def.italic !== undefined ||
    def.color !== undefined ||
    def.vertAlign !== undefined
  let rPr = firstChildNS(target, w, "rPr")
  if (hasRPrChange) {
    if (!rPr) {
      rPr = stylesDoc.createElementNS(w, "w:rPr")
      target.appendChild(rPr)
    }
    if (def.fontLatin !== undefined || def.fontCJK !== undefined) {
      const rFonts = getOrCreateNS(rPr, stylesDoc, w, "rFonts")
      if (def.fontLatin !== undefined) {
        rFonts.setAttributeNS(w, "w:ascii", def.fontLatin)
        rFonts.setAttributeNS(w, "w:hAnsi", def.fontLatin)
      }
      if (def.fontCJK !== undefined) {
        rFonts.setAttributeNS(w, "w:eastAsia", def.fontCJK)
      }
      if (!rFonts.parentNode) insertChildInOrder(rPr, rFonts, RPR_CHILD_ORDER)
    }
    if (def.size !== undefined) {
      const halfPt = toHalfPt(def.size, "size")
      const sz = getOrCreateNS(rPr, stylesDoc, w, "sz")
      sz.setAttributeNS(w, "w:val", String(halfPt))
      if (!sz.parentNode) insertChildInOrder(rPr, sz, RPR_CHILD_ORDER)
      const szCs = getOrCreateNS(rPr, stylesDoc, w, "szCs")
      szCs.setAttributeNS(w, "w:val", String(halfPt))
      if (!szCs.parentNode) insertChildInOrder(rPr, szCs, RPR_CHILD_ORDER)
    }
    if (def.bold !== undefined) {
      const bEl = firstChildNS(rPr, w, "b")
      const bCsEl = firstChildNS(rPr, w, "bCs")
      if (def.bold) {
        if (!bEl) insertChildInOrder(rPr, stylesDoc.createElementNS(w, "w:b"), RPR_CHILD_ORDER)
        if (!bCsEl)
          insertChildInOrder(rPr, stylesDoc.createElementNS(w, "w:bCs"), RPR_CHILD_ORDER)
      } else {
        if (bEl) rPr.removeChild(bEl)
        if (bCsEl) rPr.removeChild(bCsEl)
      }
    }
    if (def.italic !== undefined) {
      const iEl = firstChildNS(rPr, w, "i")
      const iCsEl = firstChildNS(rPr, w, "iCs")
      if (def.italic) {
        if (!iEl) insertChildInOrder(rPr, stylesDoc.createElementNS(w, "w:i"), RPR_CHILD_ORDER)
        if (!iCsEl)
          insertChildInOrder(rPr, stylesDoc.createElementNS(w, "w:iCs"), RPR_CHILD_ORDER)
      } else {
        if (iEl) rPr.removeChild(iEl)
        if (iCsEl) rPr.removeChild(iCsEl)
      }
    }
    if (def.color !== undefined) {
      const color = getOrCreateNS(rPr, stylesDoc, w, "color")
      color.setAttributeNS(w, "w:val", def.color)
      if (!color.parentNode) insertChildInOrder(rPr, color, RPR_CHILD_ORDER)
    }
    if (def.vertAlign !== undefined) {
      const va = getOrCreateNS(rPr, stylesDoc, w, "vertAlign")
      va.setAttributeNS(w, "w:val", def.vertAlign)
      if (!va.parentNode) insertChildInOrder(rPr, va, RPR_CHILD_ORDER)
    }
  }

  return result
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
