import { type DocxReader, parseXml } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"

/* ------------- bootstrap blank docs ------------- */

export function blankStylesDoc(): Document {
  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${NS.w}"></w:styles>`,
  )
}

export function blankNumberingDoc(): Document {
  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${NS.w}"></w:numbering>`,
  )
}

/** Read the current state of `path` — pending in-flight replacement first
 *  (some earlier subsystem already mutated it during this apply), then
 *  reader (source bytes). Returns null when neither has the entry. */
async function readPending(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
  path: string,
): Promise<string | null> {
  const pending = replacements.get(path)
  if (typeof pending === "string") return pending
  return reader.readText(path)
}

export async function ensureNumberingContentType(
  reader: DocxReader,
  replacements: Map<string, string | Uint8Array>,
): Promise<void> {
  const ctText = await readPending(reader, replacements, "[Content_Types].xml")
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
  const text = await readPending(reader, replacements, path)
  if (!text) return
  if (text.includes('Target="numbering.xml"')) return
  // pick a fresh rId
  const ids = Array.from(text.matchAll(/Id="rId(\d+)"/g)).map((m) => parseInt(m[1]!, 10))
  const next = (ids.length ? Math.max(...ids) : 0) + 1
  const insert = `<Relationship Id="rId${next}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`
  const updated = text.replace("</Relationships>", `${insert}</Relationships>`)
  replacements.set(path, updated)
}
