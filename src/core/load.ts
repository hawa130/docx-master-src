import { basename } from "node:path"
import { statSync } from "node:fs"
import { DocxReader } from "./reader.ts"
import { StyleResolver } from "./style-resolver.ts"
import { DocumentParser } from "./document-parser.ts"
import { Fingerprinter, type FingerprintSummary } from "./fingerprint.ts"
import {
  NS,
  type DocumentElement,
  type NeighborItem,
  type ParsedParagraph,
  type SectionInfo,
} from "./types.ts"
import {
  attr,
  descendantsNS,
  firstChildNS,
  getChildrenNS,
  textContent,
  wAttr,
  wVal,
} from "./xml-utils.ts"

export interface LoadedDoc {
  reader: DocxReader
  resolver: StyleResolver
  paragraphs: ParsedParagraph[]
  elements: DocumentElement[]
  neighborItems: NeighborItem[]
  sections: SectionInfo[]
  summary: FingerprintSummary[]
  metadata: {
    fileName: string
    fileSize: number
    author?: string
    title?: string
  }
  numberingDoc: Document | null
  themeDoc: Document | null
  stylesDoc: Document | null
  documentDoc: Document
}

export async function loadDocx(filePath: string): Promise<LoadedDoc> {
  const reader = await DocxReader.open(filePath)
  const stylesDoc = await reader.readXml("word/styles.xml")
  const themeDoc = await reader.readXml("word/theme/theme1.xml")
  const numberingDoc = await reader.readXml("word/numbering.xml")
  const documentDoc = await reader.readXml("word/document.xml")
  if (!documentDoc) throw new Error("word/document.xml not found")

  // load relationships
  const relsDoc = await reader.readXml("word/_rels/document.xml.rels")
  const rels = new Map<string, { type: string; target: string }>()
  if (relsDoc) {
    const root = relsDoc.documentElement
    if (root) {
      for (const r of getChildrenNS(
        root,
        "http://schemas.openxmlformats.org/package/2006/relationships",
        "Relationship",
      )) {
        const id = r.getAttribute("Id") || ""
        const type = r.getAttribute("Type") || ""
        const target = r.getAttribute("Target") || ""
        if (id) rels.set(id, { type, target })
      }
    }
  }

  // load header/footer XML files
  const headerDocs = new Map<string, Document>()
  const footerDocs = new Map<string, Document>()
  for (const [, rel] of rels) {
    if (rel.type.endsWith("/header")) {
      const path = `word/${rel.target}`
      const d = await reader.readXml(path)
      if (d) headerDocs.set(rel.target, d)
    } else if (rel.type.endsWith("/footer")) {
      const path = `word/${rel.target}`
      const d = await reader.readXml(path)
      if (d) footerDocs.set(rel.target, d)
    }
  }

  // load core properties
  const coreDoc = await reader.readXml("docProps/core.xml")
  let author: string | undefined
  let title: string | undefined
  if (coreDoc && coreDoc.documentElement) {
    const dcCreator = coreDoc.documentElement.getElementsByTagNameNS(NS.dc, "creator")[0]
    const dcTitle = coreDoc.documentElement.getElementsByTagNameNS(NS.dc, "title")[0]
    if (dcCreator) author = textContent(dcCreator) || undefined
    if (dcTitle) title = textContent(dcTitle) || undefined
  }

  const resolver = new StyleResolver(stylesDoc, themeDoc)
  const parser = new DocumentParser(documentDoc, resolver, numberingDoc, {
    headerDocs,
    footerDocs,
    rels,
  })
  const parsed = parser.parse()
  const fp = new Fingerprinter()
  const fpResult = fp.assign(parsed.paragraphs)

  const stat = statSync(filePath)

  return {
    reader,
    resolver,
    paragraphs: parsed.paragraphs,
    elements: parsed.elements,
    neighborItems: parsed.neighborItems,
    sections: parsed.sections,
    summary: fpResult.summary,
    metadata: {
      fileName: basename(filePath),
      fileSize: stat.size,
      author,
      title,
    },
    numberingDoc,
    themeDoc,
    stylesDoc,
    documentDoc,
  }
}

/* ---- numbering helpers ---- */

export interface NumberingDef {
  numId: string
  abstractNumId: string
  levels: NumberingLvlDef[]
}

export interface NumberingLvlDef {
  level: number
  format: string
  text: string
  pStyle?: string
  start: number
}

export function parseNumbering(numberingDoc: Document | null): NumberingDef[] {
  if (!numberingDoc) return []
  const root = numberingDoc.documentElement
  if (!root) return []

  const abstractMap = new Map<string, NumberingLvlDef[]>()
  for (const a of getChildrenNS(root, NS.w, "abstractNum")) {
    const id = wAttr(a, "abstractNumId") || ""
    const lvls: NumberingLvlDef[] = []
    for (const lvl of getChildrenNS(a, NS.w, "lvl")) {
      const ilvl = parseInt(wAttr(lvl, "ilvl") || "0", 10)
      const numFmt = firstChildNS(lvl, NS.w, "numFmt")
      const lvlText = firstChildNS(lvl, NS.w, "lvlText")
      const start = firstChildNS(lvl, NS.w, "start")
      const pStyle = firstChildNS(lvl, NS.w, "pStyle")
      lvls.push({
        level: ilvl,
        format: (numFmt && wVal(numFmt)) || "decimal",
        text: (lvlText && wVal(lvlText)) || "",
        pStyle: pStyle ? wVal(pStyle) || undefined : undefined,
        start: start ? parseInt(wVal(start) || "1", 10) : 1,
      })
    }
    abstractMap.set(id, lvls)
  }
  const out: NumberingDef[] = []
  for (const n of getChildrenNS(root, NS.w, "num")) {
    const id = wAttr(n, "numId") || ""
    const absRef = firstChildNS(n, NS.w, "abstractNumId")
    const absId = (absRef && wVal(absRef)) || ""
    const lvls = abstractMap.get(absId) || []
    out.push({
      numId: id,
      abstractNumId: absId,
      levels: lvls,
    })
  }
  return out
}
