/**
 * Build a `styleId → display name` resolver from a styles.xml `Document`.
 *
 * STYLEREF (ECMA-376 §17.16.5.61) takes the style's *display name*, not
 * the styleId. The OOXML engine emits inline-field nodes that carry a
 * styleId (the agent-facing handle); the emitter looks up the name at
 * write time through this resolver. Two callers wire it: edit-engine
 * (body edits[]) and header-footer-mutation (HF content).
 *
 * Returns `undefined` for unknown styleIds — caller decides whether to
 * throw or fall back.
 */

import { firstChildNS, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import { NS } from "@lib/parse/types.ts"

const w = NS.w

export function buildStyleNameResolver(
  stylesDoc: Document | null,
): (styleId: string) => string | undefined {
  if (!stylesDoc) return () => undefined
  const root = stylesDoc.documentElement
  if (!root) return () => undefined
  const map = new Map<string, string>()
  for (const styleEl of getChildrenNS(root, w, "style")) {
    const id = wAttr(styleEl, "styleId")
    if (!id) continue
    const nameEl = firstChildNS(styleEl, w, "name")
    // Falls back to styleId when `<w:name>` is missing — rare but
    // schema-legal; using the id is closer to "any sensible default" than
    // returning undefined for a style that clearly exists.
    const name = nameEl ? (wAttr(nameEl, "val") ?? id) : id
    map.set(id, name)
  }
  return (id) => map.get(id)
}
