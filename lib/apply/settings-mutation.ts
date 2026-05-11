/**
 * `word/settings.xml` mutations.
 *
 * One concern only at this point: the `<w:updateFields>` flag that tells
 * Word to update every field (including REF cross-references) when the
 * doc next opens. Without this, an inserted cross-reference shows its
 * placeholder text until the user manually triggers Ctrl+A → F9.
 *
 * Three cases:
 *   1. settings.xml exists and already has `<w:updateFields w:val="true"/>` → no-op
 *   2. settings.xml exists but lacks the element (or has val="false") → mutate / append
 *   3. settings.xml doesn't exist → fabricate a minimal one + register it in
 *      [Content_Types].xml and word/_rels/document.xml.rels
 *
 * Case 3 is rare — every Word-generated docx ships with settings.xml — but
 * common enough in hand-built fixtures that handling it cleanly is worth
 * 30 extra lines.
 */

import { DOMParser } from "@xmldom/xmldom"
import { DocxReader, serializeXml } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, wAttr } from "@lib/xml/xml-utils.ts"

const w = NS.w
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
const SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
const SETTINGS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"

const MINIMAL_SETTINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:settings xmlns:w="${w}">` +
  `<w:updateFields w:val="true"/>` +
  `</w:settings>`

/**
 * Ensure word/settings.xml carries `<w:updateFields w:val="true"/>`, and
 * stage the appropriate replacements / new-part registration into the
 * replacement map the caller is about to flush.
 *
 * Idempotent: re-running on an already-flagged doc returns without
 * touching anything.
 */
export async function ensureUpdateFieldsFlag(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
): Promise<void> {
  const settingsDoc = await reader.readXml("word/settings.xml")
  if (settingsDoc) {
    const root = settingsDoc.documentElement
    if (root) {
      let updateFields = firstChildNS(root, w, "updateFields")
      if (updateFields) {
        if (wAttr(updateFields, "val") === "true") return // already flagged
        updateFields.setAttributeNS(w, "w:val", "true")
      } else {
        updateFields = settingsDoc.createElementNS(w, "w:updateFields")
        updateFields.setAttributeNS(w, "w:val", "true")
        insertSettingsChildInOrder(root, updateFields)
      }
      replacements.set("word/settings.xml", serializeXml(settingsDoc))
      return
    }
  }

  // Path 3: fabricate. Parse the minimal stub into a Document so the
  // serializer produces consistent output formatting with the rest of the
  // docx, then register the new part in Content_Types and the doc rels.
  const parser = new DOMParser({ onError: () => {} } as any)
  const fabricated = parser.parseFromString(MINIMAL_SETTINGS_XML, "text/xml") as unknown as Document
  replacements.set("word/settings.xml", serializeXml(fabricated))
  await registerSettingsContentType(reader, replacements)
  await registerSettingsRelationship(reader, replacements)
}

/** Per ECMA-376 §17.15.1, CT_Settings has a long fixed child order.
 * `<w:updateFields>` sits late in the sequence — every element listed here
 * comes AFTER it, so we insert before the first occurrence of any of these
 * to stay schema-valid. The list intentionally only enumerates the tail of
 * CT_Settings (post-updateFields elements); pre-updateFields elements
 * don't need to be named — if none of the tail elements are present,
 * appending at the end works.  */
const SETTINGS_AFTER_UPDATEFIELDS = new Set([
  "hdrShapeDefaults",
  "footnotePr",
  "endnotePr",
  "compat",
  "docVars",
  "rsids",
  "mathPr",
  "uiCompat97To2003",
  "attachedSchema",
  "themeFontLang",
  "clrSchemeMapping",
  "doNotIncludeSubdocsInStats",
  "doNotAutoCompressPictures",
  "forceUpgrade",
  "captions",
  "readModeInkLockDown",
  "smartTagType",
  "schemaLibrary",
  "shapeDefaults",
  "doNotEmbedSmartTags",
  "decimalSymbol",
  "listSeparator",
])

function insertSettingsChildInOrder(root: Element, newEl: Element): void {
  const children = root.childNodes
  for (let i = 0; i < children.length; i++) {
    const c = children[i]
    if (!c || (c as any).nodeType !== 1) continue
    const el = c as Element
    if (el.namespaceURI === w && SETTINGS_AFTER_UPDATEFIELDS.has(el.localName!)) {
      root.insertBefore(newEl, el)
      return
    }
  }
  root.appendChild(newEl)
}

/** Add the settings Override entry to [Content_Types].xml when not present. */
async function registerSettingsContentType(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
): Promise<void> {
  const path = "[Content_Types].xml"
  const ctSerialized = replacements.get(path)
  const ctDoc =
    typeof ctSerialized === "string"
      ? (new DOMParser({ onError: () => {} } as any).parseFromString(
          ctSerialized,
          "text/xml",
        ) as unknown as Document)
      : await reader.readXml(path)
  if (!ctDoc) return
  const root = ctDoc.documentElement
  if (!root) return
  // Skip if an Override already targets /word/settings.xml.
  const overrides = root.getElementsByTagNameNS(CT_NS, "Override")
  for (let i = 0; i < overrides.length; i++) {
    if (overrides[i]!.getAttribute("PartName") === "/word/settings.xml") return
  }
  const override = ctDoc.createElementNS(CT_NS, "Override")
  override.setAttribute("PartName", "/word/settings.xml")
  override.setAttribute("ContentType", SETTINGS_CT)
  root.appendChild(override)
  replacements.set(path, serializeXml(ctDoc))
}

/** Add the settings relationship to word/_rels/document.xml.rels when not present. */
async function registerSettingsRelationship(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
): Promise<void> {
  const path = "word/_rels/document.xml.rels"
  const relsSerialized = replacements.get(path)
  const relsDoc =
    typeof relsSerialized === "string"
      ? (new DOMParser({ onError: () => {} } as any).parseFromString(
          relsSerialized,
          "text/xml",
        ) as unknown as Document)
      : await reader.readXml(path)
  if (!relsDoc) return
  const root = relsDoc.documentElement
  if (!root) return
  const rels = root.getElementsByTagNameNS(PKG_REL, "Relationship")
  let maxId = 0
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i]!
    if (rel.getAttribute("Type") === SETTINGS_REL_TYPE) return // already linked
    const id = rel.getAttribute("Id") || ""
    const m = id.match(/^rId(\d+)$/)
    if (m) {
      const n = parseInt(m[1]!, 10)
      if (n > maxId) maxId = n
    }
  }
  const rel = relsDoc.createElementNS(PKG_REL, "Relationship")
  rel.setAttribute("Id", `rId${maxId + 1}`)
  rel.setAttribute("Type", SETTINGS_REL_TYPE)
  rel.setAttribute("Target", "settings.xml")
  root.appendChild(rel)
  replacements.set(path, serializeXml(relsDoc))
}
