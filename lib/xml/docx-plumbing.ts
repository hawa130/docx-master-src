import { parseXml } from "@lib/xml/reader.ts"
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
