/**
 * Cross-document style import: copy `<w:style>` definitions from a template
 * docx into the target's styles.xml, transitively pulling in basedOn
 * ancestors and migrating any numbering references so the imported styles
 * actually work in the new document context.
 *
 * Why this lives outside apply-styles.ts: the logic is self-contained and
 * has its own data flow (open template, walk styles + numbering, mutate
 * target docs). Keeping it isolated keeps apply-styles' main pipeline
 * readable.
 */
import { DocxReader } from "./reader.ts"
import { NS } from "./types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr, wVal } from "./xml-utils.ts"

export interface ImportResult {
  /** Style IDs actually imported (includes transitive basedOn ancestors). */
  imported: string[]
  /** Old template numId → new source numId remapping (for the report). */
  numIdRemap: Map<string, string>
  /** basedOn ancestors that were auto-pulled (for the report). */
  pulledAncestors: string[]
}

/**
 * Imports the named styles from `templatePath` into the existing
 * `targetStylesDoc`, also injecting any required abstractNum/num entries
 * into `targetNumberingDoc`. Mutates both target docs in place.
 */
export async function importTemplateStyles(
  templatePath: string,
  styleIds: string[],
  targetStylesDoc: Document,
  targetNumberingDoc: Document,
  options: { importNumbering?: boolean } = {},
): Promise<ImportResult> {
  const importNumbering = options.importNumbering ?? true
  const reader = await DocxReader.open(templatePath)
  const tplStylesDoc = await reader.readXml("word/styles.xml")
  if (!tplStylesDoc) {
    throw new Error(`Template ${templatePath} has no word/styles.xml`)
  }
  const tplNumberingDoc = importNumbering ? await reader.readXml("word/numbering.xml") : null

  const tplStyles = collectStyles(tplStylesDoc)
  const targetStyles = collectStyles(targetStylesDoc)

  // Resolve transitive basedOn closure. If a style is based on another that
  // doesn't exist in the target, pull that one too — otherwise the imported
  // style would silently inherit from nothing.
  const toImport = new Set<string>()
  const pulledAncestors = new Set<string>()
  const queue = [...styleIds]
  while (queue.length > 0) {
    const id = queue.shift()!
    if (toImport.has(id)) continue
    const tpl = tplStyles.get(id)
    if (!tpl) {
      // ID not in template at all — let it fail loudly so the agent
      // notices the typo or missing style.
      throw new Error(
        `template.styles: "${id}" is not defined in ${templatePath}'s styles.xml.\n  Available: [${[...tplStyles.keys()].sort().join(", ")}]`,
      )
    }
    toImport.add(id)
    const basedOn = firstChildNS(tpl, NS.w, "basedOn")
    if (basedOn) {
      const parent = wAttr(basedOn, "val")
      if (parent && !targetStyles.has(parent) && !toImport.has(parent)) {
        // Parent not in target and not already queued: pull it in.
        // Skip if it's a Word built-in we're confident the host already has
        // ("Normal" / "DefaultParagraphFont" — both auto-created by Word).
        if (!isBuiltinAlwaysPresent(parent)) {
          pulledAncestors.add(parent)
          queue.push(parent)
        }
      }
    }
  }

  // Find numId references in the to-import set and migrate them.
  const numIdRemap = new Map<string, string>()
  if (importNumbering && tplNumberingDoc) {
    const referencedNumIds = new Set<string>()
    for (const id of toImport) {
      const style = tplStyles.get(id)!
      const pPr = firstChildNS(style, NS.w, "pPr")
      if (!pPr) continue
      const numPr = firstChildNS(pPr, NS.w, "numPr")
      if (!numPr) continue
      const numIdEl = firstChildNS(numPr, NS.w, "numId")
      if (!numIdEl) continue
      const v = wAttr(numIdEl, "val")
      if (v) referencedNumIds.add(v)
    }
    for (const oldNumId of referencedNumIds) {
      const fresh = migrateNumIdToTarget(oldNumId, tplNumberingDoc, targetNumberingDoc)
      if (fresh !== null) numIdRemap.set(oldNumId, fresh)
    }
  }

  // Inject (or replace) each style in the target. We clone the template
  // node into the target document so namespaces and ownership are correct.
  for (const id of toImport) {
    const tplNode = tplStyles.get(id)!
    const cloned = importStyleNode(tplNode, targetStylesDoc, numIdRemap)
    upsertStyleNode(targetStylesDoc, id, cloned)
  }

  return {
    imported: [...toImport],
    numIdRemap,
    pulledAncestors: [...pulledAncestors],
  }
}

function collectStyles(stylesDoc: Document): Map<string, Element> {
  const out = new Map<string, Element>()
  const root = stylesDoc.documentElement
  if (!root) return out
  for (const s of getChildrenNS(root, NS.w, "style")) {
    const id = wAttr(s, "styleId")
    if (id) out.set(id, s)
  }
  return out
}

function isBuiltinAlwaysPresent(styleId: string): boolean {
  // Word always provides these even in a "blank" document; safe to skip.
  // We're not 100% guaranteed but it's a sane default. Anything else gets
  // auto-pulled to be safe.
  return styleId === "Normal" || styleId === "DefaultParagraphFont"
}

/**
 * Clone a <w:style> element from the template document into the target,
 * remapping any numPr/numId references via numIdRemap.
 */
function importStyleNode(
  src: Element,
  targetDoc: Document,
  numIdRemap: Map<string, string>,
): Element {
  const cloned = deepCloneIntoDocument(src, targetDoc)
  // Walk for numPr/numId children and remap
  const w = NS.w
  const visit = (el: Element) => {
    if (el.namespaceURI === w && el.localName === "numId") {
      const v = wAttr(el, "val")
      if (v && numIdRemap.has(v)) {
        el.setAttributeNS(w, "w:val", numIdRemap.get(v)!)
      }
    }
    for (const c of getChildren(el)) visit(c)
  }
  visit(cloned)
  return cloned
}

function deepCloneIntoDocument(src: Element, targetDoc: Document): Element {
  const ns = src.namespaceURI
  const tag = src.tagName // includes prefix (e.g. "w:style")
  const out = ns ? targetDoc.createElementNS(ns, tag) : targetDoc.createElement(tag)
  // copy attributes
  if (src.attributes) {
    for (let i = 0; i < src.attributes.length; i++) {
      const a = src.attributes.item(i)!
      if (a.namespaceURI) {
        out.setAttributeNS(a.namespaceURI, a.name, a.value)
      } else {
        out.setAttribute(a.name, a.value)
      }
    }
  }
  for (const child of Array.from(src.childNodes)) {
    if (child.nodeType === 1) {
      // ELEMENT_NODE
      out.appendChild(deepCloneIntoDocument(child as Element, targetDoc))
    } else if (child.nodeType === 3 || child.nodeType === 4) {
      // TEXT / CDATA
      out.appendChild(targetDoc.createTextNode(child.nodeValue || ""))
    }
  }
  return out
}

/**
 * Replace any existing <w:style> with the same styleId, otherwise append.
 */
function upsertStyleNode(stylesDoc: Document, styleId: string, node: Element) {
  const root = stylesDoc.documentElement!
  const existing = getChildrenNS(root, NS.w, "style").find((s) => wAttr(s, "styleId") === styleId)
  if (existing) {
    root.replaceChild(node, existing)
  } else {
    root.appendChild(node)
  }
}

/**
 * Copy the abstractNum + num pair referenced by `oldNumId` from the template
 * numbering doc into the target. Returns the new numId in the target doc,
 * or null if the template doesn't actually have that numId.
 */
function migrateNumIdToTarget(
  oldNumId: string,
  tplNumberingDoc: Document,
  targetNumberingDoc: Document,
): string | null {
  const w = NS.w
  const tplRoot = tplNumberingDoc.documentElement!
  const targetRoot = targetNumberingDoc.documentElement!

  // Find template num
  const tplNum = getChildrenNS(tplRoot, w, "num").find((n) => wAttr(n, "numId") === oldNumId)
  if (!tplNum) return null
  const absRef = firstChildNS(tplNum, w, "abstractNumId")
  if (!absRef) return null
  const oldAbsId = wVal(absRef)
  if (!oldAbsId) return null
  // Find template abstractNum
  const tplAbs = getChildrenNS(tplRoot, w, "abstractNum").find(
    (a) => wAttr(a, "abstractNumId") === oldAbsId,
  )
  if (!tplAbs) return null

  // Pick fresh IDs in the target
  const existingAbsIds = getChildrenNS(targetRoot, w, "abstractNum").map((e) =>
    parseInt(wAttr(e, "abstractNumId") || "0", 10),
  )
  const existingNumIds = getChildrenNS(targetRoot, w, "num").map((e) =>
    parseInt(wAttr(e, "numId") || "0", 10),
  )
  const newAbsId = (existingAbsIds.length ? Math.max(...existingAbsIds) : -1) + 1
  const newNumId = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1

  // Clone abstractNum, retag its abstractNumId
  const clonedAbs = deepCloneIntoDocument(tplAbs, targetNumberingDoc)
  clonedAbs.setAttributeNS(w, "w:abstractNumId", String(newAbsId))
  // abstractNum must precede num children
  const firstNum = getChildrenNS(targetRoot, w, "num")[0]
  if (firstNum) targetRoot.insertBefore(clonedAbs, firstNum)
  else targetRoot.appendChild(clonedAbs)

  // Clone num and retag its numId + abstractNumId reference
  const clonedNum = deepCloneIntoDocument(tplNum, targetNumberingDoc)
  clonedNum.setAttributeNS(w, "w:numId", String(newNumId))
  const newAbsRef = firstChildNS(clonedNum, w, "abstractNumId")
  if (newAbsRef) newAbsRef.setAttributeNS(w, "w:val", String(newAbsId))
  targetRoot.appendChild(clonedNum)

  return String(newNumId)
}
