/**
 * Hyperlink inline node — emits `<w:hyperlink>` wrapping one or more runs.
 *
 *   `{ link: "https://x", text: "..." }`  → `<w:hyperlink r:id="rIdN" w:history="1">`
 *   `{ link: "#fig-x",   text: "..." }`  → `<w:hyperlink w:anchor="fig-x">`
 *
 * External links register a Relationship via the shared ImageAssetRegistry
 * (TargetMode="External"); internal `#anchor` links use the w:anchor
 * attribute directly with no rId allocation. The visible text is
 * agent-supplied — Word doesn't compute it like REF/SEQ.
 *
 * Visual style: the run carries the `Hyperlink` character style (Word's
 * blue + underlined convention). `ensureHyperlinkCharStyle` injects the
 * style into styles.xml on first use when it isn't already present.
 */

import { NS } from "@lib/parse/types.ts"
import { getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import { RPR_CHILD_ORDER, insertChildInOrder } from "@lib/xml/xml-order.ts"
import { buildRPrChildren } from "@lib/edit/fragment-emit.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"
import type { ImageAssetRegistry } from "@lib/edit/image-asset.ts"

const w = NS.w
const HYPERLINK_STYLE_ID = "Hyperlink"

/** Parse a link string into either an internal anchor name or an external
 *  URI. Schema's refine already enforces `#`-prefix format for anchors;
 *  this is a runtime split, not a re-validation. */
export function parseLinkTarget(link: string): { kind: "anchor"; name: string } | { kind: "url"; href: string } {
  if (link.startsWith("#")) return { kind: "anchor", name: link.slice(1) }
  return { kind: "url", href: link }
}

/** Build a `<w:hyperlink>` element containing one run carrying the visible
 *  text. The run is bound to the `Hyperlink` character style; any agent-
 *  supplied `format` layers on top via run rPr. */
export function emitHyperlinkNode(
  ownerDoc: Document,
  link: string,
  text: string,
  format: RunFormat | undefined,
  assetRegistry: ImageAssetRegistry,
): Element {
  const target = parseLinkTarget(link)
  const hyper = ownerDoc.createElementNS(w, "w:hyperlink")
  if (target.kind === "anchor") {
    hyper.setAttributeNS(w, "w:anchor", target.name)
  } else {
    const { rId } = assetRegistry.registerExternalLink(target.href)
    hyper.setAttributeNS(NS.r, "r:id", rId)
    // history=1 matches Word's default and makes the link styled as visited
    // on subsequent clicks. Cosmetic but matches "looks like a hyperlink".
    hyper.setAttributeNS(w, "w:history", "1")
  }

  const run = ownerDoc.createElementNS(w, "w:r")
  const rPr = ownerDoc.createElementNS(w, "w:rPr")
  // rStyle first per CT_RPr schema order — buildRPrChildren handles its own
  // children's ordering but doesn't know about rStyle.
  const rStyle = ownerDoc.createElementNS(w, "w:rStyle")
  rStyle.setAttributeNS(w, "w:val", HYPERLINK_STYLE_ID)
  rPr.appendChild(rStyle)
  if (format) {
    for (const c of buildRPrChildren(format, ownerDoc)) {
      insertChildInOrder(rPr, c, RPR_CHILD_ORDER)
    }
  }
  run.appendChild(rPr)

  const t = ownerDoc.createElementNS(w, "w:t")
  t.setAttribute("xml:space", "preserve")
  t.textContent = text
  run.appendChild(t)
  hyper.appendChild(run)
  return hyper
}

/** Inject the `Hyperlink` character style into stylesDoc when missing. Word
 *  treats `Hyperlink` as a built-in id; the canonical name "Hyperlink" and
 *  the conventional 0563C1 / single-underline rPr match what Word
 *  generates when you paste a URL. Idempotent — call once per apply, no-op
 *  when the style already exists. */
export function ensureHyperlinkCharStyle(stylesDoc: Document): boolean {
  const root = stylesDoc.documentElement
  if (!root) return false
  for (const s of getChildrenNS(root, w, "style")) {
    if (wAttr(s, "styleId") === HYPERLINK_STYLE_ID) return false
  }
  const style = stylesDoc.createElementNS(w, "w:style")
  style.setAttributeNS(w, "w:type", "character")
  style.setAttributeNS(w, "w:styleId", HYPERLINK_STYLE_ID)

  const name = stylesDoc.createElementNS(w, "w:name")
  name.setAttributeNS(w, "w:val", "Hyperlink")
  style.appendChild(name)

  const basedOn = stylesDoc.createElementNS(w, "w:basedOn")
  basedOn.setAttributeNS(w, "w:val", "DefaultParagraphFont")
  style.appendChild(basedOn)

  const uiPriority = stylesDoc.createElementNS(w, "w:uiPriority")
  uiPriority.setAttributeNS(w, "w:val", "99")
  style.appendChild(uiPriority)

  style.appendChild(stylesDoc.createElementNS(w, "w:unhideWhenUsed"))

  const rPr = stylesDoc.createElementNS(w, "w:rPr")
  const color = stylesDoc.createElementNS(w, "w:color")
  color.setAttributeNS(w, "w:val", "0563C1")
  color.setAttributeNS(w, "w:themeColor", "hyperlink")
  rPr.appendChild(color)
  const u = stylesDoc.createElementNS(w, "w:u")
  u.setAttributeNS(w, "w:val", "single")
  rPr.appendChild(u)
  style.appendChild(rPr)

  // Append at end of styles list (no schema-order constraints for sibling
  // <w:style> entries; doc order is presentation only).
  root.appendChild(style)
  return true
}

/** Quick existence check used by callers that want to skip the injection
 *  branch entirely (e.g. when no hyperlink-bearing edits[] op exists). */
export function hasHyperlinkCharStyle(stylesDoc: Document): boolean {
  const root = stylesDoc.documentElement
  if (!root) return false
  for (const s of getChildrenNS(root, w, "style")) {
    if (wAttr(s, "styleId") === HYPERLINK_STYLE_ID) return true
  }
  return false
}

