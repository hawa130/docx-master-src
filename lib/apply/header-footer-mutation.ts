/**
 * Header / footer part generation.
 *
 * For each declared `headerFooter.<surface>.<variant>` (surface ∈ {header,
 * footer}, variant ∈ {default, first, even}), we materialise one new
 * `word/headerN.xml` or `word/footerN.xml` part:
 *
 *   1. emit Block[] content via `fragment-emit.emitBlock`, wrapping in a
 *      `<w:hdr>` or `<w:ftr>` root element with the namespaces the
 *      emitted runs may use (w / r / wp / a / pic / m / mc)
 *   2. each part owns its own `PartRels` so images-in-HF resolve against
 *      `word/_rels/<part>.rels` (per-part rId namespace, per ECMA-376)
 *   3. images / external hyperlinks in HF blocks register through a
 *      per-part `DocxAssetRegistry`; the shared `ContentTypes` accumulator
 *      threads through every registry so `Default Extension` + `Override
 *      PartName` entries land in one file
 *   4. register an `<Override>` for the HF part in `[Content_Types].xml`
 *      and a `<Relationship>` in the BODY's rels (document.xml.rels) so
 *      the sectPr binding step (`applyHeaderFooterBinding`) has a rId
 *      to plug into `<w:headerReference>` / `<w:footerReference>`
 *
 * Part naming starts at the highest existing index + 1, so re-running
 * apply on a doc that already has headers/footers appends new parts
 * rather than colliding. Old parts are left orphaned in the archive —
 * sectPr binding (B3) replaces every section's references, so Word
 * never reads them. v2 may opt to GC orphans.
 *
 * Triggers (`hasFirst` / `hasEven` in the report) feed B3: `<w:titlePg/>`
 * on each sectPr when any surface has a `first` variant, and
 * `<w:evenAndOddHeaders/>` in settings.xml when any surface has `even`.
 */

import type { DocxReader } from "@lib/xml/reader.ts"
import { parseXml, serializeXml } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import { SECT_PR_CHILD_ORDER, insertChildInOrder } from "@lib/xml/xml-order.ts"
import type { ApplyConfig } from "@lib/config/config-types.ts"
import type { Block } from "@lib/config/edit-types.ts"
import { emitBlock, type EmitContext } from "@lib/edit/fragment-emit.ts"
import { PartRels } from "@lib/edit/part-rels.ts"
import { ContentTypes } from "@lib/edit/content-types.ts"
import { DocxAssetRegistry } from "@lib/edit/asset-registry.ts"
import { emitHyperlinkNode } from "@lib/edit/hyperlink.ts"

type HeaderFooterConfig = NonNullable<ApplyConfig["headerFooter"]>
type Surface = "header" | "footer"
type Variant = "default" | "first" | "even"

const VARIANT_ORDER: readonly Variant[] = ["default", "first", "even"] as const
const SURFACE_ORDER: readonly Surface[] = ["header", "footer"] as const

const HEADER_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
const FOOTER_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
const HEADER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
const FOOTER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

export interface HeaderFooterPartRecord {
  surface: Surface
  variant: Variant
  /** rId allocated in the BODY's rels, used by sectPr binding for
   *  `<w:headerReference r:id="..."/>`. */
  rId: string
  /** Full part name with leading slash, e.g. `/word/header1.xml`. The
   *  leading slash matches `[Content_Types].xml` Override convention; the
   *  archive entry path (without slash) is `partName.slice(1)`. */
  partName: string
  blockCount: number
  /** True when the part body contains at least one external hyperlink. The
   *  apply pipeline uses this to ensure the `Hyperlink` character style is
   *  injected into styles.xml even when no body edits[] declare one. */
  hasHyperlinks: boolean
}

export interface HeaderFooterReport {
  parts: HeaderFooterPartRecord[]
  /** Any surface declared a `first` variant — drives `<w:titlePg/>` on
   *  every sectPr in B3. */
  hasFirst: boolean
  /** Any surface declared an `even` variant — drives
   *  `<w:evenAndOddHeaders/>` in settings.xml in B3. */
  hasEven: boolean
}

/** Skeleton XML for a fresh `<w:hdr>` or `<w:ftr>` part. Carries every
 *  namespace a body paragraph might bring along (drawingML, math, mc) so
 *  serialization doesn't have to splice declarations onto descendants. */
function buildHfSkeleton(surface: Surface): Document {
  const root = surface === "header" ? "w:hdr" : "w:ftr"
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<${root} xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}"` +
    ` xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:m="${NS.m}" xmlns:mc="${NS.mc}"/>`
  return parseXml(xml)
}

/** Scan the source archive for existing `word/header<N>.xml` /
 *  `word/footer<N>.xml` entries. Returns the next free index — new parts
 *  start at this number and increment per-emission, preventing collision
 *  with parts left over from a prior apply. */
function nextFreePartIndex(reader: DocxReader): number {
  let max = 0
  for (const entry of reader.listEntries()) {
    const m = entry.match(/^word\/(?:header|footer)(\d+)\.xml$/)
    if (m) {
      const n = parseInt(m[1]!, 10)
      if (n > max) max = n
    }
  }
  return max + 1
}

/** Materialise one (surface, variant) pair: emit blocks, register
 *  rels/content-types, stage the XML + binary additions. Mutates
 *  `replacements` and the shared `contentTypes` / `bodyPartRels` in place;
 *  returns the metadata the report and sectPr binding need. */
function emitOnePart(args: {
  surface: Surface
  variant: Variant
  blocks: Block[]
  partIndex: number
  reader: DocxReader
  bodyPartRels: PartRels
  contentTypes: ContentTypes
  replacements: Map<string, string | Uint8Array>
}): HeaderFooterPartRecord {
  const { surface, variant, blocks, partIndex, bodyPartRels, contentTypes, replacements } = args
  const partFileName = `${surface}${partIndex}.xml`
  const partPath = `word/${partFileName}` // archive entry path
  const partNameSlash = `/${partPath}` // Override convention
  const relsPath = `word/_rels/${partFileName}.rels`

  // Fresh per-part rels — new file, so we don't `PartRels.open` from the
  // reader (no existing entry). Construct on empty skeleton.
  const partRels = new PartRels(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  )
  const partRegistry = new DocxAssetRegistry(partRels, contentTypes)

  // Build the XML document.
  const doc = buildHfSkeleton(surface)
  const root = doc.documentElement!

  let hasHyperlinks = false
  const ctx: EmitContext = {
    emitImage: (src, w, h, alt, ownerDoc) => {
      const { rId } = partRegistry.registerImage(src)
      return partRegistry.buildDrawing(rId, w, h, alt, ownerDoc)
    },
    emitHyperlink: (link, text, format, ownerDoc) => {
      // parseLinkTarget inside emitHyperlinkNode discriminates anchor vs url;
      // url path consumes the registry (rels rId), anchor path doesn't.
      if (!link.startsWith("#")) hasHyperlinks = true
      return emitHyperlinkNode(ownerDoc, link, text, format, partRegistry)
    },
    // adoptAnchor / emitRef / captions intentionally absent — HF excludes
    // caption / equation blocks at schema; ParagraphBlock.anchor inside HF
    // surfaces a clean fragment-emit error so the agent learns the
    // exclusion concretely.
  }

  for (const block of blocks) {
    root.appendChild(emitBlock(block, doc, ctx))
  }

  // Stage XML + per-part binaries + per-part rels.
  replacements.set(partPath, serializeXml(doc))
  partRegistry.flushTo(replacements, relsPath)

  // Register the part in shared Content_Types and the BODY's rels.
  contentTypes.ensureOverride(
    partNameSlash,
    surface === "header" ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE,
  )
  const { rId } = bodyPartRels.appendRel(
    surface === "header" ? HEADER_REL_TYPE : FOOTER_REL_TYPE,
    partFileName, // target relative to word/_rels/document.xml.rels base
  )

  return {
    surface,
    variant,
    rId,
    partName: partNameSlash,
    blockCount: blocks.length,
    hasHyperlinks,
  }
}

/** Emit every declared HF part, registering each in Content_Types and the
 *  body rels. The caller is responsible for: (a) flushing the body's own
 *  `PartRels` (since this function appends to it) and (b) flushing the
 *  shared `ContentTypes` (single flush at end of apply, like the body's
 *  asset registry — multi-registry race is avoided by the orchestrator
 *  owning the single flush). */
export async function applyHeaderFooter(
  reader: DocxReader,
  config: HeaderFooterConfig,
  bodyPartRels: PartRels,
  contentTypes: ContentTypes,
  replacements: Map<string, string | Uint8Array>,
): Promise<HeaderFooterReport> {
  const parts: HeaderFooterPartRecord[] = []
  let hasFirst = false
  let hasEven = false
  let partIndex = nextFreePartIndex(reader)

  for (const surface of SURFACE_ORDER) {
    const surfaceCfg = config[surface]
    if (!surfaceCfg) continue
    for (const variant of VARIANT_ORDER) {
      const blocks = surfaceCfg[variant]
      if (blocks === undefined) continue
      if (variant === "first") hasFirst = true
      if (variant === "even") hasEven = true

      const record = emitOnePart({
        surface,
        variant,
        blocks: blocks as Block[],
        partIndex,
        reader,
        bodyPartRels,
        contentTypes,
        replacements,
      })
      parts.push(record)
      partIndex++
    }
  }

  return { parts, hasFirst, hasEven }
}

/* ------------- sectPr binding ------------- */

const w = NS.w

/** Collect every sectPr in body order — paragraph-embedded sectPrs first
 *  (sections 1..N-1) then the body-trailing one (section N). Mirrors
 *  `collectSectPrs` in page-setup-mutation; duplicated here rather than
 *  exported to keep page-setup's surface narrow (the HF binding pass
 *  doesn't need the indexing semantics page-setup exposes). */
function collectSectPrs(body: Element): Element[] {
  const out: Element[] = []
  for (const child of getChildren(body)) {
    if (child.namespaceURI !== w) continue
    if (child.localName === "p") {
      const pPr = firstChildNS(child, w, "pPr")
      if (pPr) {
        const sectPr = firstChildNS(pPr, w, "sectPr")
        if (sectPr) out.push(sectPr)
      }
    } else if (child.localName === "sectPr") {
      out.push(child)
    }
  }
  return out
}

/** Remove any pre-existing `<w:headerReference>` / `<w:footerReference>`
 *  children from a sectPr. We replace wholesale rather than merge — v1
 *  rule: HF declarations apply to EVERY section (decision 2). Leftover
 *  references from a prior apply would mix old + new and confuse Word
 *  about which part wins for each variant. */
function stripExistingReferences(sectPr: Element): void {
  const stale: Element[] = []
  for (const c of getChildren(sectPr)) {
    if (c.namespaceURI !== w) continue
    if (c.localName === "headerReference" || c.localName === "footerReference") {
      stale.push(c)
    }
  }
  for (const el of stale) sectPr.removeChild(el)
}

/** Idempotent set/clear of `<w:titlePg/>` on a sectPr based on `enabled`.
 *  HF config is the source of truth: with the block declared, titlePg
 *  is set when any surface has a `first` variant, cleared otherwise.
 *  Returns true when the DOM actually changed; false on no-op. */
function setTitlePg(sectPr: Element, doc: Document, enabled: boolean): boolean {
  const existing = firstChildNS(sectPr, w, "titlePg")
  if (enabled) {
    if (existing) return false
    const el = doc.createElementNS(w, "w:titlePg")
    insertChildInOrder(sectPr, el, SECT_PR_CHILD_ORDER)
    return true
  }
  if (!existing) return false
  sectPr.removeChild(existing)
  return true
}

export interface HeaderFooterBindingReport {
  /** Number of sectPrs we wrote references onto. */
  sectionCount: number
  /** True when at least one sectPr's `<w:titlePg/>` flag was actually
   *  mutated this run — either set (because `first` was declared and the
   *  flag was missing) or cleared (because `first` is not declared and
   *  the source had a leftover flag from a prior apply). */
  titlePgApplied: boolean
}

/** Plug HF part records into every sectPr in `documentDoc`.
 *
 *   - clears pre-existing headerReference / footerReference children
 *   - appends one `<w:headerReference w:type="..." r:id="..."/>` per
 *     declared header variant, one `<w:footerReference>` per footer
 *     variant
 *   - sets `<w:titlePg/>` on every sectPr when any HF surface declared
 *     `first` (decision 7)
 *
 *  evenAndOddHeaders activation lives in settings.xml — see
 *  `setEvenAndOddHeadersFlag` in settings-mutation.
 */
/* ------------- Header / Footer paragraph styles ------------- */

/** Inject the built-in `Header` and `Footer` paragraph styles into
 *  stylesDoc when missing. Word treats these styleIds as well-known —
 *  the canonical name "header" / "footer", uiPriority="99",
 *  basedOn="Normal" match what Word generates when you insert a header
 *  via the UI.
 *
 *  Deliberately minimal: no tab stops (the classic split layout uses a
 *  3-column borderless table — see `references/header-footer.md`) and
 *  no bottom border (Word's default Header style ships one, but it's a
 *  divisive cosmetic choice — agents wanting it declare paraFormat).
 *
 *  Idempotent. Returns true when at least one style was injected. */
export function ensureHeaderFooterStyles(stylesDoc: Document): boolean {
  const root = stylesDoc.documentElement
  if (!root) return false
  const existing = new Set<string>()
  for (const s of getChildrenNS(root, w, "style")) {
    const id = wAttr(s, "styleId")
    if (id) existing.add(id)
  }
  let injected = false
  for (const { id, name } of [
    { id: "Header", name: "header" },
    { id: "Footer", name: "footer" },
  ]) {
    if (existing.has(id)) continue
    const style = stylesDoc.createElementNS(w, "w:style")
    style.setAttributeNS(w, "w:type", "paragraph")
    style.setAttributeNS(w, "w:styleId", id)

    const nm = stylesDoc.createElementNS(w, "w:name")
    nm.setAttributeNS(w, "w:val", name)
    style.appendChild(nm)

    const basedOn = stylesDoc.createElementNS(w, "w:basedOn")
    basedOn.setAttributeNS(w, "w:val", "Normal")
    style.appendChild(basedOn)

    const uiPriority = stylesDoc.createElementNS(w, "w:uiPriority")
    uiPriority.setAttributeNS(w, "w:val", "99")
    style.appendChild(uiPriority)

    style.appendChild(stylesDoc.createElementNS(w, "w:unhideWhenUsed"))

    root.appendChild(style)
    injected = true
  }
  return injected
}

/* ------------- sectPr binding (continued) ------------- */

export function applyHeaderFooterBinding(
  documentDoc: Document,
  report: HeaderFooterReport,
): HeaderFooterBindingReport {
  const body = firstChildNS(documentDoc.documentElement!, w, "body")
  if (!body) return { sectionCount: 0, titlePgApplied: false }
  const sectPrs = collectSectPrs(body)
  if (sectPrs.length === 0) return { sectionCount: 0, titlePgApplied: false }

  let titlePgApplied = false
  for (const sectPr of sectPrs) {
    stripExistingReferences(sectPr)
    for (const part of report.parts) {
      const ref = documentDoc.createElementNS(
        w,
        part.surface === "header" ? "w:headerReference" : "w:footerReference",
      )
      ref.setAttributeNS(w, "w:type", part.variant)
      ref.setAttributeNS(NS.r, "r:id", part.rId)
      insertChildInOrder(sectPr, ref, SECT_PR_CHILD_ORDER)
    }
    if (setTitlePg(sectPr, documentDoc, report.hasFirst)) titlePgApplied = true
  }
  return { sectionCount: sectPrs.length, titlePgApplied }
}

