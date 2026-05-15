/**
 * Document-asset registry — wires image blocks AND external hyperlinks into
 * the docx zip. Both kinds of assets register entries in
 * `word/_rels/document.xml.rels`, so they share one writer to avoid two
 * paths racing on the same file. The class is still named after its first
 * client (images); the hyperlink path was added in the inline-primitives
 * phase.
 *
 * For images, registration touches three parts besides word/document.xml:
 *   1. word/media/imageN.<ext>           — the binary file
 *   2. word/_rels/document.xml.rels      — relationship from rId to media path
 *   3. [Content_Types].xml               — Default entry for the file extension
 *
 * For external hyperlinks: just the rels entry (TargetMode="External").
 * Internal hyperlinks (`#anchor` form) skip the registry entirely — they
 * use `<w:hyperlink w:anchor="name">` with no rId.
 *
 * The registry caches existing rels / extensions so multiple emits in one
 * apply pass coexist with each other and with whatever the source docx
 * already had. After all emits, `flushTo(replacements)` stages every
 * mutation (text + binary) into the writer's replacement map.
 *
 * Inline drawing XML structure follows ECMA-376 §20.4.2.8 (DrawingML
 * picture). Sizes are in EMU (1 pt = 12,700 EMU).
 */

import { existsSync, readFileSync } from "node:fs"
import { extname, resolve as resolvePath } from "node:path"
import { type DocxReader, parseXml } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"
import { type Length, toEmu } from "@lib/shared/units.ts"

const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
const REL_TYPE_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

interface RelEntry {
  id: string
  target: string
  type: string
}

export class ImageAssetRegistry {
  private contentTypesText: string
  private relsText: string
  /** archivePath → bytes. New media binaries staged for the writer. */
  private binaryAdditions = new Map<string, Uint8Array>()
  private existingExtensions = new Set<string>()
  private rels: RelEntry[] = []
  private nextImageNum = 1
  private nextDocPrId = 1
  private nextRId = 1
  /** Set whenever rels are appended (image OR hyperlink). Required because
   *  hyperlink-only flushes still need to write the rels file even though
   *  `binaryAdditions` stays empty. */
  private relsDirty = false
  /** Cache: srcPath → already-registered rId. Lets two image blocks pointing
   * at the same file share one media entry. */
  private byAbsPath = new Map<string, string>()
  /** Cache: external href → already-registered rId. Multiple hyperlinks to
   *  the same target share one Relationship entry — matches what Word writes
   *  when you paste the same URL twice. */
  private byExternalHref = new Map<string, string>()

  private constructor(contentTypesText: string, relsText: string) {
    this.contentTypesText = contentTypesText
    this.relsText = relsText
    this.parseExtensions()
    this.parseRels()
  }

  static async open(reader: DocxReader): Promise<ImageAssetRegistry> {
    const ct = (await reader.readText("[Content_Types].xml")) ?? ""
    const rels = (await reader.readText("word/_rels/document.xml.rels")) ?? defaultRelsXml()
    return new ImageAssetRegistry(ct || defaultContentTypesXml(), rels)
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
    if (cached) {
      const r = this.rels.find((x) => x.id === cached)!
      return { rId: cached, ext: extname(r.target).slice(1).toLowerCase() }
    }
    const rawExt = extname(abs).slice(1).toLowerCase()
    if (!rawExt) throw new Error(`image source has no extension: ${abs}`)
    const ext = rawExt === "jpg" ? "jpeg" : rawExt
    const archivePath = this.uniqueMediaPath(rawExt)
    const bytes = new Uint8Array(readFileSync(abs))
    this.binaryAdditions.set(archivePath, bytes)
    if (!this.existingExtensions.has(rawExt)) {
      this.appendContentTypeDefault(rawExt, mimeForExt(ext))
      this.existingExtensions.add(rawExt)
    }
    const rId = this.allocateRId()
    this.appendRel(rId, REL_TYPE_IMAGE, archivePath.replace(/^word\//, ""))
    this.byAbsPath.set(abs, rId)
    return { rId, ext }
  }

  /** Register an external hyperlink target. Returns the rId for the
   *  `<w:hyperlink r:id="...">` element. Identical hrefs share one rId
   *  (matches Word's deduplication). Internal `#anchor` links don't go
   *  through this path — they use `<w:hyperlink w:anchor="...">` with
   *  no rId. */
  registerExternalLink(href: string): { rId: string } {
    const cached = this.byExternalHref.get(href)
    if (cached) return { rId: cached }
    const rId = this.allocateRId()
    this.appendRel(rId, REL_TYPE_HYPERLINK, href, "External")
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
   * image flushes additionally write content types + media binaries. */
  flushTo(replacements: Map<string, string | Uint8Array>): void {
    if (!this.relsDirty && this.binaryAdditions.size === 0) return
    if (this.relsDirty) {
      replacements.set("word/_rels/document.xml.rels", this.relsText)
    }
    if (this.binaryAdditions.size > 0) {
      replacements.set("[Content_Types].xml", this.contentTypesText)
      for (const [path, bytes] of this.binaryAdditions) {
        replacements.set(path, bytes)
      }
    }
  }

  /* ------------- internals ------------- */

  private parseExtensions(): void {
    const re = /<Default[^>]*\bExtension="([^"]+)"/g
    let m: RegExpMatchArray | null
    while ((m = re.exec(this.contentTypesText)) !== null) {
      this.existingExtensions.add(m[1]!.toLowerCase())
    }
  }

  private parseRels(): void {
    const doc = parseXml(this.relsText)
    const root = doc.documentElement
    if (!root) return
    let maxNum = 0
    for (let i = 0; i < root.childNodes.length; i++) {
      const c = root.childNodes[i]
      if (!c || c.nodeType !== 1) continue
      const el = c as Element
      if (el.localName !== "Relationship") continue
      const id = el.getAttribute("Id") ?? ""
      const target = el.getAttribute("Target") ?? ""
      const type = el.getAttribute("Type") ?? ""
      if (id) {
        this.rels.push({ id, target, type })
        const m = id.match(/^rId(\d+)$/)
        if (m) maxNum = Math.max(maxNum, parseInt(m[1]!, 10))
      }
      const mt = target.match(/^(?:\.\.\/)?media\/image(\d+)\./)
      if (mt) {
        const n = parseInt(mt[1]!, 10) + 1
        if (n > this.nextImageNum) this.nextImageNum = n
      }
    }
    this.nextRId = maxNum + 1
  }

  private allocateRId(): string {
    return `rId${this.nextRId++}`
  }

  private uniqueMediaPath(ext: string): string {
    while (true) {
      const path = `word/media/image${this.nextImageNum}.${ext}`
      this.nextImageNum++
      if (!this.binaryAdditions.has(path)) return path
    }
  }

  private appendContentTypeDefault(ext: string, mime: string): void {
    const insert = `<Default Extension="${ext}" ContentType="${mime}"/>`
    if (this.contentTypesText.includes(insert)) return
    this.contentTypesText = this.contentTypesText.replace("</Types>", `${insert}</Types>`)
  }

  private appendRel(id: string, type: string, target: string, mode?: "External"): void {
    // target stored relative to the rels' part location: rels live in
    // word/_rels/, so a media file at word/media/imageN.ext is "media/imageN.ext".
    // External targets (hyperlink URLs) are stored verbatim; the writer
    // doesn't rewrite them and Word resolves them at click time.
    const modeAttr = mode ? ` TargetMode="${mode}"` : ""
    const escapedTarget = escapeXmlAttr(target)
    const insert = `<Relationship Id="${id}" Type="${type}" Target="${escapedTarget}"${modeAttr}/>`
    this.relsText = this.relsText.replace("</Relationships>", `${insert}</Relationships>`)
    this.rels.push({ id, target, type })
    this.relsDirty = true
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

function defaultContentTypesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`
}

function defaultRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
}

/** Minimal XML-attribute escape — image targets are sanitized media paths
 *  but hyperlink hrefs are agent-supplied URIs that may contain `&` (query
 *  strings), `"`, `<`, `>`. Encode just enough to keep the rels XML valid. */
function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
}
