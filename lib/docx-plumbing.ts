import { readFileSync } from "node:fs"
import JSZip from "jszip"
import { DOMParser } from "@xmldom/xmldom"
import type { DocxReader } from "@lib/reader.ts"
import { NS } from "@lib/types.ts"

/* ------------- bootstrap blank docs ------------- */

export function blankStylesDoc(): Document {
  const text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}"></w:styles>`
  return new DOMParser().parseFromString(text, "text/xml") as unknown as Document
}

export function blankNumberingDoc(): Document {
  const text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS.w}"></w:numbering>`
  return new DOMParser().parseFromString(text, "text/xml") as unknown as Document
}

export async function ensureNumberingContentType(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
): Promise<void> {
  const ctText = await reader.readText("[Content_Types].xml")
  if (!ctText) return
  if (ctText.includes("/word/numbering.xml")) return
  const insert = `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
  const updated = ctText.replace("</Types>", `${insert}</Types>`)
  replacements.set("[Content_Types].xml", updated)
}

export async function ensureNumberingRelationship(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
): Promise<void> {
  const path = "word/_rels/document.xml.rels"
  const text = await reader.readText(path)
  if (!text) return
  if (text.includes('Target="numbering.xml"')) return
  // pick a fresh rId
  const ids = Array.from(text.matchAll(/Id="rId(\d+)"/g)).map((m) => parseInt(m[1]!, 10))
  const next = (ids.length ? Math.max(...ids) : 0) + 1
  const insert = `<Relationship Id="rId${next}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`
  const updated = text.replace("</Relationships>", `${insert}</Relationships>`)
  replacements.set(path, updated)
}

/* ------------- validation ------------- */

export async function validateOutput(
  outputPath: string,
  modifiedEntries: string[],
): Promise<{ ok: true } | { ok: false; error: string }> {
  try {
    const buf = readFileSync(outputPath)
    const zip = await JSZip.loadAsync(buf)
    for (const entry of modifiedEntries) {
      const file = zip.file(entry)
      if (!file) continue
      const text = await file.async("string")
      let parseError: string | null = null
      const parser = new DOMParser({
        onError: (level: any, msg: any) => {
          if (level === "error" || level === "fatalError") parseError = String(msg)
        },
      } as any)
      const doc = parser.parseFromString(text, "text/xml")
      if (parseError) return { ok: false, error: `${entry}: ${parseError}` }
      if (!doc || !(doc as any).documentElement) return { ok: false, error: `${entry}: empty doc` }
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: (err as Error).message }
  }
}
