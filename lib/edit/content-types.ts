/**
 * `[Content_Types].xml` accumulator — one shared instance per apply run.
 *
 * Multiple parts (document.xml + headerN.xml + footerN.xml) may all need
 * entries: `Default` for new media extensions (images), `Override` for
 * named parts (headers, footers, numbering, settings, ...). Routing every
 * change through a single accumulator keeps the file's state coherent
 * across the body-asset registry and per-HF-part registries.
 */

import type { DocxReader } from "@lib/xml/reader.ts"

const DEFAULT_CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`

export class ContentTypes {
  private text: string
  private extensions = new Set<string>()
  private overridePartNames = new Set<string>()
  private dirty = false

  constructor(initialText: string) {
    this.text = initialText
    this.parseExtensions()
    this.parseOverrides()
  }

  static async open(reader: DocxReader): Promise<ContentTypes> {
    const text = (await reader.readText("[Content_Types].xml")) ?? DEFAULT_CONTENT_TYPES_XML
    return new ContentTypes(text || DEFAULT_CONTENT_TYPES_XML)
  }

  /** Ensure a `<Default Extension="..." ContentType="..."/>` entry exists.
   *  Idempotent — duplicate calls for the same extension are no-ops. */
  ensureDefault(ext: string, mime: string): void {
    const key = ext.toLowerCase()
    if (this.extensions.has(key)) return
    const insert = `<Default Extension="${ext}" ContentType="${mime}"/>`
    if (this.text.includes(insert)) {
      this.extensions.add(key)
      return
    }
    this.text = this.text.replace("</Types>", `${insert}</Types>`)
    this.extensions.add(key)
    this.dirty = true
  }

  /** Ensure an `<Override PartName="..." ContentType="..."/>` entry exists.
   *  Used to register headerN.xml / footerN.xml / settings.xml / etc. by
   *  full part name. Idempotent. */
  ensureOverride(partName: string, contentType: string): void {
    if (this.overridePartNames.has(partName)) return
    const insert = `<Override PartName="${partName}" ContentType="${contentType}"/>`
    if (this.text.includes(insert)) {
      this.overridePartNames.add(partName)
      return
    }
    this.text = this.text.replace("</Types>", `${insert}</Types>`)
    this.overridePartNames.add(partName)
    this.dirty = true
  }

  isDirty(): boolean {
    return this.dirty
  }

  /** Stage the modified content-types XML in the writer's replacement map.
   *  No-op when nothing was added. */
  flushTo(replacements: Map<string, string | Uint8Array>): void {
    if (!this.dirty) return
    replacements.set("[Content_Types].xml", this.text)
  }

  /* ------------- internals ------------- */

  private parseExtensions(): void {
    const re = /<Default[^>]*\bExtension="([^"]+)"/g
    let m: RegExpMatchArray | null
    while ((m = re.exec(this.text)) !== null) {
      this.extensions.add(m[1]!.toLowerCase())
    }
  }

  private parseOverrides(): void {
    const re = /<Override[^>]*\bPartName="([^"]+)"/g
    let m: RegExpMatchArray | null
    while ((m = re.exec(this.text)) !== null) {
      this.overridePartNames.add(m[1]!)
    }
  }
}
