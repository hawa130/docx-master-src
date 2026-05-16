/**
 * DOM helpers narrowed for MathML tree walking. Skips text/comment nodes
 * (MathML is element-structural — whitespace text between elements is
 * not semantic and Word's OMML reader rejects extra text nodes at
 * structural levels like inside <m:nary>).
 */

import { MML_NS, M_NS, W_NS } from "./constants.ts"

export function elementChildren(parent: Element): Element[] {
  const out: Element[] = []
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) out.push(n as Element)
  }
  return out
}

export function isMmlElement(n: Node | null, localName?: string): n is Element {
  if (!n || n.nodeType !== 1) return false
  const el = n as Element
  if (el.namespaceURI !== MML_NS) return false
  return localName === undefined || el.localName === localName
}

/** Concatenated text content of an Element, MathML-flavored — trims
 *  surrounding whitespace because MathML leaf text is conventionally
 *  exactly the symbol/identifier with no padding, but some serializers
 *  emit a stray newline inside <mo>+</mo>. */
export function mmlText(el: Element): string {
  return (el.textContent ?? "").trim()
}

/** Single named attribute, undefined when absent. Works for MathML
 *  attributes which are typically unnamespaced. */
export function attr(el: Element, name: string): string | undefined {
  return el.hasAttribute(name) ? el.getAttribute(name)! : undefined
}

/** Create an element in the OMML math namespace using the `m:` prefix
 *  so xmldom emits readable `<m:nary>` instead of `<nary xmlns="…">`. */
export function mEl(doc: Document, localName: string): Element {
  return doc.createElementNS(M_NS, `m:${localName}`)
}

/** Create an element in the wordprocessingml namespace, used only for
 *  the `<w:rPr><w:rFonts/></w:rPr>` font-selection block that lives
 *  inside every OMML run. */
export function wEl(doc: Document, localName: string): Element {
  return doc.createElementNS(W_NS, `w:${localName}`)
}

/** Set a `m:val` attribute on a "property" element (e.g. `<m:chr m:val="∑"/>`).
 *  All OMML property attributes use this single shape. */
export function setMVal(el: Element, value: string): void {
  el.setAttributeNS(M_NS, "m:val", value)
}
