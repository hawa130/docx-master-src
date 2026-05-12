import { readFileSync, writeFileSync } from "node:fs"
import JSZip from "jszip"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"

/** @xmldom/xmldom's `Document` is structurally compatible with TS's
 * `lib.dom` `Document` but typed as a separate nominal type. Every call
 * site that produces a Document via `DOMParser` would otherwise repeat a
 * `as unknown as Document` cast; centralizing here keeps the type-erasure
 * in one place. The same module owns serialization for symmetry. */
const xmlParser = new DOMParser({ onError: () => {} } as ConstructorParameters<typeof DOMParser>[0])

export function parseXml(text: string): Document {
  return xmlParser.parseFromString(text, "text/xml") as unknown as Document
}

export class DocxReader {
  private zip: JSZip
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

  async readText(entryPath: string): Promise<string | null> {
    const entry = this.zip.file(entryPath)
    if (!entry) return null
    return entry.async("string")
  }

  async readXml(entryPath: string): Promise<Document | null> {
    const text = await this.readText(entryPath)
    if (text === null) return null
    return parseXml(text)
  }

  /** Path-only list of every entry in the archive. Used by the validator to
   * enumerate XML / .rels parts without hard-coding the part list. */
  listEntries(): string[] {
    const out: string[] = []
    this.zip.forEach((relativePath: string) => {
      out.push(relativePath)
    })
    return out
  }

  /**
   * Copy zip to outputPath, replacing the listed entries. Values are either
   * XML/text strings (most cases) or binary `Uint8Array` (image assets).
   * JSZip handles both shapes natively. Adding a *new* archive entry uses
   * the same call: `replacements.set("word/media/image1.png", bytes)`.
   */
  async copyAndModify(
    outputPath: string,
    replacements: Map<string, string | Uint8Array>,
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

const xmlSerializer = new XMLSerializer()

export function serializeXml(doc: Document): string {
  // @xmldom omits the XML declaration when serializing a Document; re-add it
  // since Word expects the standard `<?xml ...?>` prelude on every part.
  // The cross-type cast: see parseXml above for the nominal-type rationale.
  const out = xmlSerializer.serializeToString(
    doc as unknown as Parameters<XMLSerializer["serializeToString"]>[0],
  )
  if (out.startsWith("<?xml")) return out
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`
}
