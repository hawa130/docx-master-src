/**
 * Document-asset registry — handles image and external-hyperlink registration
 * for a single docx part (body or any header / footer). The class composes:
 *
 *   - `PartRels` — owns this part's `_rels/<part>.rels` (rId allocation +
 *     Relationship appends)
 *   - `ContentTypes` — shared `[Content_Types].xml` accumulator across all
 *     parts in the apply run (singleton, passed in by the orchestrator)
 *   - image-specific state: media binaries staged for the writer, docPr id
 *     counter, source-path / href dedup caches
 *
 * For images, registration touches three parts besides the host XML:
 *   1. word/media/imageN.<ext>      — the binary file (per-instance state)
 *   2. <relsPath>                   — relationship from rId to media path
 *   3. [Content_Types].xml          — Default entry for the file extension
 *
 * For external hyperlinks: just the rels entry (TargetMode="External").
 * Internal hyperlinks (`#anchor` form) skip the registry entirely — they
 * use `<w:hyperlink w:anchor="name">` with no rId.
 *
 * One instance serves one part. The body of document.xml gets one
 * registry pointing at `word/_rels/document.xml.rels`; each header /
 * footer part gets its own registry pointing at `word/_rels/headerN.xml.rels`
 * (etc.). All registries share one ContentTypes accumulator.
 *
 * Inline drawing XML structure follows ECMA-376 §20.4.2.8 (DrawingML
 * picture). Sizes are in EMU (1 pt = 12,700 EMU).
 */

import { existsSync, readFileSync } from "node:fs"
import { extname, resolve as resolvePath } from "node:path"
import type { DocxReader } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"
import { type Length, toEmu } from "@lib/shared/units.ts"
import { PartRels } from "@lib/edit/part-rels.ts"
import { ContentTypes } from "@lib/edit/content-types.ts"

const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
const REL_TYPE_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

export class DocxAssetRegistry {
  /** archivePath → bytes. New media binaries staged for the writer. */
  private binaryAdditions = new Map<string, Uint8Array>()
  private nextDocPrId = 1
  /** Cache: srcPath → already-registered rId + ext. Lets two image blocks
   *  pointing at the same file share one media entry. */
  private byAbsPath = new Map<string, { rId: string; ext: string }>()
  /** Cache: external href → already-registered rId. Multiple hyperlinks to
   *  the same target share one Relationship entry — matches what Word writes
   *  when you paste the same URL twice. */
  private byExternalHref = new Map<string, string>()

  constructor(
    private readonly partRels: PartRels,
    private readonly contentTypes: ContentTypes,
  ) {}

  /** Expose the shared ContentTypes accumulator. The orchestrator
   *  (apply-styles) flushes it once at the end of the apply run, and HF
   *  part registries get constructed with the same instance so all of
   *  the apply's content-type additions land in one file. */
  getContentTypes(): ContentTypes {
    return this.contentTypes
  }

  /** Convenience: open the body-part registry from a docx reader. Loads
   *  `word/_rels/document.xml.rels` and `[Content_Types].xml`. For header /
   *  footer parts, the orchestrator constructs `PartRels` and the shared
   *  `ContentTypes` directly and passes them in. */
  static async open(reader: DocxReader): Promise<DocxAssetRegistry> {
    const partRels = await PartRels.open(reader, "word/_rels/document.xml.rels")
    const contentTypes = await ContentTypes.open(reader)
    return new DocxAssetRegistry(partRels, contentTypes)
  }

  /** Returns the rId to use in <a:blip r:embed=...>. Reads the file, picks
   * a fresh archive path (word/media/imageN.ext), stages the binary, adds
   * the rel + content type if needed. Throws on missing source file. */
  registerImage(srcPath: string): { rId: string; ext: string } {
    const abs = resolvePath(srcPath)
    if (!existsSync(abs)) {
      throw new Error(`image not found: ${abs}`)
    }
    const cached = this.byAbsPath.get(abs)
    if (cached) return { rId: cached.rId, ext: cached.ext }
    const rawExt = extname(abs).slice(1).toLowerCase()
    if (!rawExt) throw new Error(`image source has no extension: ${abs}`)
    // ext stays as the source's file suffix ("jpg" not "jpeg") so the
    // archivePath, the [Content_Types].xml Default entry, and the cached
    // return value all agree. mimeForExt internally maps "jpg" / "jpeg" /
    // "tif" / "tiff" to the canonical MIME, so no normalization needed.
    const archivePath = this.uniqueMediaPath(rawExt)
    const bytes = new Uint8Array(readFileSync(abs))
    this.binaryAdditions.set(archivePath, bytes)
    this.contentTypes.ensureDefault(rawExt, mimeForExt(rawExt))
    const { rId } = this.partRels.appendRel(
      REL_TYPE_IMAGE,
      archivePath.replace(/^word\//, ""),
    )
    this.byAbsPath.set(abs, { rId, ext: rawExt })
    return { rId, ext: rawExt }
  }

  /** Register an external hyperlink target. Returns the rId for the
   *  `<w:hyperlink r:id="...">` element. Identical hrefs share one rId
   *  (matches Word's deduplication). Internal `#anchor` links don't go
   *  through this path — they use `<w:hyperlink w:anchor="...">` with
   *  no rId. */
  registerExternalLink(href: string): { rId: string } {
    const cached = this.byExternalHref.get(href)
    if (cached) return { rId: cached }
    const { rId } = this.partRels.appendRel(REL_TYPE_HYPERLINK, href, "External")
    this.byExternalHref.set(href, rId)
    return { rId }
  }

  /** Build a single inline <w:drawing> element ready to live inside a <w:r>.
   * EMU-converts the dimensions; `alt` populates docPr name + descr. */
  buildDrawing(
    rId: string,
    width: Length,
    height: Length,
    alt: string | undefined,
    ownerDoc: Document,
  ): Element {
    const cx = String(toEmu(width, "image.width"))
    const cy = String(toEmu(height, "image.height"))
    const docPrId = String(this.nextDocPrId++)
    const altText = alt ?? `Image ${docPrId}`

    const drawing = ownerDoc.createElementNS(NS.w, "w:drawing")
    const inline = ownerDoc.createElementNS(NS.wp, "wp:inline")
    inline.setAttribute("distT", "0")
    inline.setAttribute("distB", "0")
    inline.setAttribute("distL", "0")
    inline.setAttribute("distR", "0")
    drawing.appendChild(inline)

    const extent = ownerDoc.createElementNS(NS.wp, "wp:extent")
    extent.setAttribute("cx", cx)
    extent.setAttribute("cy", cy)
    inline.appendChild(extent)

    const effectExtent = ownerDoc.createElementNS(NS.wp, "wp:effectExtent")
    effectExtent.setAttribute("l", "0")
    effectExtent.setAttribute("t", "0")
    effectExtent.setAttribute("r", "0")
    effectExtent.setAttribute("b", "0")
    inline.appendChild(effectExtent)

    const docPr = ownerDoc.createElementNS(NS.wp, "wp:docPr")
    docPr.setAttribute("id", docPrId)
    docPr.setAttribute("name", altText)
    if (alt) docPr.setAttribute("descr", alt)
    inline.appendChild(docPr)

    const cNvGraphicFramePr = ownerDoc.createElementNS(NS.wp, "wp:cNvGraphicFramePr")
    const graphicFrameLocks = ownerDoc.createElementNS(NS.a, "a:graphicFrameLocks")
    graphicFrameLocks.setAttribute("noChangeAspect", "1")
    cNvGraphicFramePr.appendChild(graphicFrameLocks)
    inline.appendChild(cNvGraphicFramePr)

    const graphic = ownerDoc.createElementNS(NS.a, "a:graphic")
    inline.appendChild(graphic)

    const graphicData = ownerDoc.createElementNS(NS.a, "a:graphicData")
    graphicData.setAttribute("uri", PIC_NS)
    graphic.appendChild(graphicData)

    const pic = ownerDoc.createElementNS(PIC_NS, "pic:pic")
    graphicData.appendChild(pic)

    const nvPicPr = ownerDoc.createElementNS(PIC_NS, "pic:nvPicPr")
    pic.appendChild(nvPicPr)
    const cNvPr = ownerDoc.createElementNS(PIC_NS, "pic:cNvPr")
    cNvPr.setAttribute("id", docPrId)
    cNvPr.setAttribute("name", altText)
    nvPicPr.appendChild(cNvPr)
    nvPicPr.appendChild(ownerDoc.createElementNS(PIC_NS, "pic:cNvPicPr"))

    const blipFill = ownerDoc.createElementNS(PIC_NS, "pic:blipFill")
    pic.appendChild(blipFill)
    const blip = ownerDoc.createElementNS(NS.a, "a:blip")
    blip.setAttributeNS(NS.r, "r:embed", rId)
    blipFill.appendChild(blip)
    const stretch = ownerDoc.createElementNS(NS.a, "a:stretch")
    stretch.appendChild(ownerDoc.createElementNS(NS.a, "a:fillRect"))
    blipFill.appendChild(stretch)

    const spPr = ownerDoc.createElementNS(PIC_NS, "pic:spPr")
    pic.appendChild(spPr)
    const xfrm = ownerDoc.createElementNS(NS.a, "a:xfrm")
    spPr.appendChild(xfrm)
    const off = ownerDoc.createElementNS(NS.a, "a:off")
    off.setAttribute("x", "0")
    off.setAttribute("y", "0")
    xfrm.appendChild(off)
    const ext = ownerDoc.createElementNS(NS.a, "a:ext")
    ext.setAttribute("cx", cx)
    ext.setAttribute("cy", cy)
    xfrm.appendChild(ext)
    const prstGeom = ownerDoc.createElementNS(NS.a, "a:prstGeom")
    prstGeom.setAttribute("prst", "rect")
    spPr.appendChild(prstGeom)
    prstGeom.appendChild(ownerDoc.createElementNS(NS.a, "a:avLst"))

    return drawing
  }

  /** Stage every mutation (modified [Content_Types].xml, modified rels,
   * binary media files) into the writer's replacement map. No-op when
   * nothing was registered. Hyperlinks-only flushes write the rels alone;
   * image flushes additionally write content types + media binaries.
   *
   * `relsPath` defaults to the body part's rels file. Header / footer
   * registries pass their own part rels path. */
  flushTo(
    replacements: Map<string, string | Uint8Array>,
    relsPath = "word/_rels/document.xml.rels",
  ): void {
    this.partRels.flushTo(replacements, relsPath)
    for (const [path, bytes] of this.binaryAdditions) {
      replacements.set(path, bytes)
    }
    // contentTypes is shared across all registries in the apply run; the
    // orchestrator flushes it once at the end. Don't write it here — would
    // race with other registries holding the same accumulator.
  }

  /* ------------- internals ------------- */

  private uniqueMediaPath(ext: string): string {
    while (true) {
      const n = this.partRels.nextImageNum()
      this.partRels.advanceImageNum(n)
      const path = `word/media/image${n}.${ext}`
      if (!this.binaryAdditions.has(path)) return path
    }
  }
}

/* ------------- helpers ------------- */

function mimeForExt(ext: string): string {
  switch (ext) {
    case "png":
      return "image/png"
    case "jpeg":
    case "jpg":
      return "image/jpeg"
    case "gif":
      return "image/gif"
    case "bmp":
      return "image/bmp"
    case "tiff":
    case "tif":
      return "image/tiff"
    case "svg":
      return "image/svg+xml"
    case "webp":
      return "image/webp"
    default:
      return `image/${ext}`
  }
}

