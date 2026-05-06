import { NS } from "./types.ts"

function isElement(node: any): node is Element {
  return node && node.nodeType === 1
}

export function getChildren(parent: Element | Document | null): Element[] {
  if (!parent) return []
  const out: Element[] = []
  const children = (parent as any).childNodes
  if (!children) return out
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
  return getChildren(parent).filter(
    (e) => e.namespaceURI === ns && e.localName === localName,
  )
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
  walk(parent as any, (n) => {
    if (n.namespaceURI === ns && n.localName === localName) out.push(n)
  })
  return out
}

function walk(node: any, fn: (e: Element) => void) {
  const children = node.childNodes
  if (!children) return
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
  const children = (el as any).childNodes
  if (!children) return out
  for (let i = 0; i < children.length; i++) {
    const n = children[i]
    if (n.nodeType === 3 || n.nodeType === 4) out += n.nodeValue || ""
    else if (n.nodeType === 1) out += textContent(n)
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
