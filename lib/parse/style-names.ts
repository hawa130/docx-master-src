/**
 * Build a `styleId → StyleInfo` resolver from a styles.xml `Document`.
 *
 * Used by `emitInlineStyleRef` to drive a 3-way dispatch on the
 * referenced style:
 *   1. style has `<w:outlineLvl>` → emit `STYLEREF N` (locale-neutral)
 *   2. style is a custom name (not in the built-in localizable set) →
 *      emit `STYLEREF "<name>"` (Word doesn't translate custom names)
 *   3. style is a built-in localizable non-outline style (Title /
 *      Caption / Subtitle / etc.) → caller should reject, since
 *      `STYLEREF "<English-name>"` silently fails in non-EN Word UIs
 *
 * Returns `undefined` for unknown styleIds — caller decides whether
 * to throw or fall back.
 */

import { firstChildNS, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import { NS } from "@lib/parse/types.ts"
import { isBuiltInLocalizable } from "@lib/parse/builtin-styles.ts"

const w = NS.w

export interface StyleInfo {
  /** `<w:name w:val="..."/>` value, falling back to styleId when the
   *  style omits a name element. */
  name: string
  /** 1-9 when the style declares `<w:pPr><w:outlineLvl w:val="N"/>`
   *  (OOXML's 0-indexed N+1). `undefined` for non-outline styles. */
  outlineLevel: number | undefined
  /** True when `name` matches a Word built-in style whose display form
   *  is locale-translated by Word (e.g. "Title" ↔ "标题"). Used by
   *  STYLEREF emit to refuse locale-unsafe field codes. */
  isBuiltInLocalizable: boolean
}

export function buildStyleResolver(
  stylesDoc: Document | null,
): (styleId: string) => StyleInfo | undefined {
  if (!stylesDoc) return () => undefined
  const root = stylesDoc.documentElement
  if (!root) return () => undefined
  const map = new Map<string, StyleInfo>()
  for (const styleEl of getChildrenNS(root, w, "style")) {
    const id = wAttr(styleEl, "styleId")
    if (!id) continue
    const nameEl = firstChildNS(styleEl, w, "name")
    // Falls back to styleId when `<w:name>` is missing — rare but
    // schema-legal; using the id is closer to "any sensible default"
    // than returning undefined for a style that clearly exists.
    const name = nameEl ? (wAttr(nameEl, "val") ?? id) : id
    let outlineLevel: number | undefined
    const pPr = firstChildNS(styleEl, w, "pPr")
    if (pPr) {
      const lvlEl = firstChildNS(pPr, w, "outlineLvl")
      if (lvlEl) {
        const v = wAttr(lvlEl, "val")
        const n = v !== null && v !== undefined ? parseInt(v, 10) : NaN
        // OOXML stores 0-indexed; STYLEREF takes 1-indexed.
        if (Number.isFinite(n) && n >= 0 && n <= 8) outlineLevel = n + 1
      }
    }
    map.set(id, {
      name,
      outlineLevel,
      isBuiltInLocalizable: isBuiltInLocalizable(name),
    })
  }
  return (id) => map.get(id)
}
