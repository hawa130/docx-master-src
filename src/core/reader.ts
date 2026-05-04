import { readFileSync, writeFileSync } from "node:fs"
import JSZip from "jszip"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"

export class DocxReader {
  private zip: JSZip
  private parser = new DOMParser({
    onError: () => {},
  } as any)
  filePath: string
  fileSize: number

  private constructor(filePath: string, zip: JSZip, size: number) {
    this.filePath = filePath
    this.zip = zip
    this.fileSize = size
  }

  static async open(filePath: string): Promise<DocxReader> {
    const buf = readFileSync(filePath)
    const zip = await JSZip.loadAsync(buf)
    return new DocxReader(filePath, zip, buf.length)
  }

  hasEntry(entryPath: string): boolean {
    return this.zip.file(entryPath) !== null
  }

  async readText(entryPath: string): Promise<string | null> {
    const entry = this.zip.file(entryPath)
    if (!entry) return null
    return entry.async("string")
  }

  async readXml(entryPath: string): Promise<Document | null> {
    const text = await this.readText(entryPath)
    if (text === null) return null
    return this.parser.parseFromString(text, "text/xml") as unknown as Document
  }

  async readBinary(entryPath: string): Promise<Buffer | null> {
    const entry = this.zip.file(entryPath)
    if (!entry) return null
    const u8 = await entry.async("uint8array")
    return Buffer.from(u8)
  }

  listEntries(): string[] {
    const names: string[] = []
    this.zip.forEach((path) => names.push(path))
    return names
  }

  /**
   * Copy zip to outputPath, replacing the listed entries with new XML strings.
   * Replacement values are full XML documents serialized to string.
   */
  async copyAndModify(
    outputPath: string,
    replacements: Map<string, string>,
  ): Promise<void> {
    // load fresh zip from original buffer to avoid mutating the open one
    const original = readFileSync(this.filePath)
    const out = await JSZip.loadAsync(original)
    for (const [path, content] of replacements) {
      out.file(path, content)
    }
    const buffer = await out.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    })
    writeFileSync(outputPath, buffer)
  }
}

export const xmlSerializer = new XMLSerializer()

export function serializeXml(doc: Document): string {
  // ensure declaration; @xmldom does not include one by default for top-level Document
  const out = xmlSerializer.serializeToString(doc as any)
  if (out.startsWith("<?xml")) return out
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`
}
