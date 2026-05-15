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

import { DocxReader, parseXml, serializeXml } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, wAttr } from "@lib/xml/xml-utils.ts"

const w = NS.w
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
const SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
const SETTINGS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"

/** Empty stub for the fabrication path. Each mutator adds its own children
 *  via `insertSettingsChildInOrder`; leaving the stub child-free avoids
 *  baking in flags the caller didn't request. */
const MINIMAL_SETTINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:settings xmlns:w="${w}"></w:settings>`

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
  await mutateSettings(reader, replacements, (root) => {
    let updateFields = firstChildNS(root, w, "updateFields")
    if (updateFields) {
      if (wAttr(updateFields, "val") === "true") return false // already flagged
      updateFields.setAttributeNS(w, "w:val", "true")
      return true
    }
    updateFields = root.ownerDocument!.createElementNS(w, "w:updateFields")
    updateFields.setAttributeNS(w, "w:val", "true")
    insertSettingsChild(root, updateFields, SETTINGS_AFTER_UPDATEFIELDS)
    return true
  })
}

/**
 * Set or clear `<w:evenAndOddHeaders/>` in settings.xml based on
 * `enabled`. The flag is presence-only (no @val); Word treats the empty
 * element as enabling distinct even/odd headers and footers. Pair with
 * sectPrs that carry header/footer references with `w:type="even"`.
 *
 * Called by the HF pipeline as the source of truth: with HF config
 * declared, the flag is enabled iff at least one surface declared an
 * `even` variant. The fabrication path (no settings.xml in source)
 * is skipped when `enabled=false` to avoid creating settings.xml just
 * to remove a flag that wasn't there.
 *
 * Idempotent: re-run on an already-correct doc no-ops.
 */
export async function setEvenAndOddHeadersFlag(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
  enabled: boolean,
): Promise<void> {
  await mutateSettings(
    reader,
    replacements,
    (root) => {
      const existing = firstChildNS(root, w, "evenAndOddHeaders")
      if (enabled) {
        if (existing) return false
        const el = root.ownerDocument!.createElementNS(w, "w:evenAndOddHeaders")
        insertSettingsChild(root, el, SETTINGS_AFTER_EVENANDODDHEADERS)
        return true
      }
      if (!existing) return false
      root.removeChild(existing)
      return true
    },
    { fabricateOnMissing: enabled },
  )
}

/** Shared settings.xml mutation flow. Loads (or reuses an in-flight
 *  replacement of) settings.xml, hands the root to `mutate`, and stages
 *  the result. Fabricates a minimal settings.xml + registers the part
 *  in Content_Types and the body rels when none exists.
 *
 *  `mutate` returns true when it changed the doc; false (already in
 *  desired state) skips re-serialization but still falls through to the
 *  fabricate path if no settings.xml existed at all. */
async function mutateSettings(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
  mutate: (root: Element) => boolean,
  opts: { fabricateOnMissing?: boolean } = {},
): Promise<void> {
  const fabricateOnMissing = opts.fabricateOnMissing ?? true
  // Honour pending in-flight changes from earlier ensure calls in this run
  // so the second mutator sees the first's edits (e.g. updateFields then
  // evenAndOddHeaders both target the same file).
  const pending = replacements.get("word/settings.xml")
  let settingsDoc: Document | null = null
  if (typeof pending === "string") {
    settingsDoc = parseXml(pending)
  } else {
    settingsDoc = await reader.readXml("word/settings.xml")
  }

  if (settingsDoc) {
    const root = settingsDoc.documentElement
    if (root) {
      const changed = mutate(root)
      if (changed) replacements.set("word/settings.xml", serializeXml(settingsDoc))
      return
    }
  }

  // Path 3: fabricate from the minimal stub, then let mutate set its
  // contribution on top. Skipped when `fabricateOnMissing=false`
  // (e.g. clearing a flag that wasn't there — no need to materialise
  // settings.xml just to keep it empty).
  if (!fabricateOnMissing) return
  const fabricated = parseXml(MINIMAL_SETTINGS_XML)
  const root = fabricated.documentElement
  if (root) mutate(root)
  replacements.set("word/settings.xml", serializeXml(fabricated))
  await registerSettingsContentType(reader, replacements)
  await registerSettingsRelationship(reader, replacements)
}

/** Per ECMA-376 §17.15.1, CT_Settings has a long fixed child order.
 * Rather than enumerate the full ~80-element sequence, each ensure
 * function declares the set of element local-names that come AFTER its
 * own insertion — `insertSettingsChild` then places the new element
 * before the first existing child found in that set. Pre-element
 * children stay untouched; trailing position is preserved when none of
 * the AFTER elements are present. */
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

/** Elements that come AFTER `<w:evenAndOddHeaders>` in CT_Settings.
 * Superset of SETTINGS_AFTER_UPDATEFIELDS plus the run of mid-tail
 * elements between defaultTableStyle (position ~47) and updateFields
 * (position ~74). Without these, an insert against a source doc that
 * already carries updateFields (which most do) misplaces the new
 * evenAndOddHeaders AFTER updateFields, producing schema-invalid XML. */
const SETTINGS_AFTER_EVENANDODDHEADERS = new Set([
  "bookFoldRevPrinting",
  "bookFoldPrinting",
  "bookFoldPrintingSheets",
  "drawingGridHorizontalSpacing",
  "drawingGridVerticalSpacing",
  "displayHorizontalDrawingGridEvery",
  "displayVerticalDrawingGridEvery",
  "doNotUseMarginsForDrawingGridOrigin",
  "drawingGridHorizontalOrigin",
  "drawingGridVerticalOrigin",
  "doNotShadeFormData",
  "noPunctuationKerning",
  "characterSpacingControl",
  "printTwoOnOne",
  "strictFirstAndLastChars",
  "noLineBreaksAfter",
  "noLineBreaksBefore",
  "savePreviewPicture",
  "doNotValidateAgainstSchema",
  "saveInvalidXml",
  "ignoreMixedContent",
  "alwaysShowPlaceholderText",
  "doNotDemarcateInvalidXml",
  "saveXmlDataOnly",
  "useXSLTWhenSaving",
  "saveThroughXslt",
  "showXMLTags",
  "alwaysMergeEmptyNamespace",
  "updateStyles",
  "updateFields",
  ...SETTINGS_AFTER_UPDATEFIELDS,
])

function insertSettingsChild(root: Element, newEl: Element, afterSet: Set<string>): void {
  for (const el of getChildren(root)) {
    if (el.namespaceURI === w && afterSet.has(el.localName!)) {
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
    typeof ctSerialized === "string" ? parseXml(ctSerialized) : await reader.readXml(path)
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
    typeof relsSerialized === "string" ? parseXml(relsSerialized) : await reader.readXml(path)
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
