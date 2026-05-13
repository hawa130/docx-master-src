import { NS } from "@lib/parse/types.ts"

function isElement(node: Node | null | undefined): node is Element {
  return !!node && node.nodeType === 1
}

export function getChildren(parent: Element | Document | null): Element[] {
  if (!parent) return []
  const out: Element[] = []
  const children = parent.childNodes
  for (let i = 0; i < children.length; i++) {
    const n = children[i]
    if (isElement(n)) out.push(n)
  }
  return out
}

export function getChildrenNS(
  parent: Element | Document | null,
  ns: string,
  localName: string,
): Element[] {
  return getChildren(parent).filter((e) => e.namespaceURI === ns && e.localName === localName)
}

export function firstChildNS(
  parent: Element | Document | null,
  ns: string,
  localName: string,
): Element | null {
  const all = getChildren(parent)
  for (const e of all) {
    if (e.namespaceURI === ns && e.localName === localName) return e
  }
  return null
}

export function descendantsNS(
  parent: Element | Document | null,
  ns: string,
  localName: string,
): Element[] {
  if (!parent) return []
  const out: Element[] = []
  walk(parent, (n) => {
    if (n.namespaceURI === ns && n.localName === localName) out.push(n)
  })
  return out
}

function walk(node: Element | Document, fn: (e: Element) => void) {
  const children = node.childNodes
  for (let i = 0; i < children.length; i++) {
    const c = children[i]
    if (isElement(c)) {
      fn(c)
      walk(c, fn)
    }
  }
}

export function attr(el: Element | null, ns: string, name: string): string | null {
  if (!el) return null
  const v = el.getAttributeNS(ns, name)
  if (v !== null && v !== "") return v
  // fall back to non-namespaced lookup (common for w: attributes)
  const v2 = el.getAttribute(`w:${name}`)
  if (v2) return v2
  const v3 = el.getAttribute(name)
  return v3 || null
}

export function wAttr(el: Element | null, name: string): string | null {
  return attr(el, NS.w, name)
}

export function wVal(el: Element | null): string | null {
  return wAttr(el, "val")
}

export function textContent(el: Element): string {
  let out = ""
  const children = el.childNodes
  for (let i = 0; i < children.length; i++) {
    const n = children[i]
    if (!n) continue
    if (n.nodeType === 3 || n.nodeType === 4) out += n.nodeValue || ""
    else if (isElement(n)) out += textContent(n)
  }
  return out
}

/** Boolean toggle parsing: present with no val, or val="1"/"true" → true; val="0"/"false" → false; missing → undefined */
export function parseToggle(el: Element | null): boolean | undefined {
  if (!el) return undefined
  const v = wVal(el)
  if (v === null || v === undefined) return true
  if (v === "0" || v === "false") return false
  return true
}

/** Depth-first walk yielding every `<w:p>` reachable from `root`,
 * descending through `<w:tbl>` / `<w:tr>` / `<w:tc>` containers. Used
 * by the caption pipeline (counter sim, standardize re-emit,
 * edit-caption resolver) and the inspection / migration tools — they
 * all need the same shape of traversal. Distinct from the
 * `walkIndexedParagraphs` in `lib/edit/locator.ts` which is layout-
 * table aware and returns indexed pairs. */
export function* walkBodyParagraphs(root: Element): Generator<Element> {
  for (const child of getChildren(root)) {
    if (child.namespaceURI !== NS.w) continue
    if (child.localName === "p") {
      yield child
    } else if (child.localName === "tbl" || child.localName === "tr" || child.localName === "tc") {
      yield* walkBodyParagraphs(child)
    }
  }
}

/** Build a `<w:r>` containing a single `<w:t xml:space="preserve">text</w:t>`.
 * Idiom shared by caption emit, edit-caption op, and standardize re-emit
 * for literal decoration runs (prefix / suffix / separators / body
 * text). Preserve-space ensures leading / trailing whitespace survives
 * XML serialization. */
export function buildPlainTextRun(ownerDoc: Document, text: string): Element {
  const r = ownerDoc.createElementNS(NS.w, "w:r")
  const t = ownerDoc.createElementNS(NS.w, "w:t")
  t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
  t.textContent = text
  r.appendChild(t)
  return r
}
