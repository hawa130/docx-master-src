export interface ComputedRunStyle {
  fontAscii?: string
  fontEastAsia?: string
  fontHAnsi?: string
  size?: number          // half-points
  bold?: boolean
  italic?: boolean
  underline?: string
  color?: string         // hex without # or "auto"
  highlight?: string
  strike?: boolean
  caps?: boolean
}

export interface ComputedParaStyle {
  pStyle?: string
  alignment?: string
  spaceBefore?: number   // twips
  spaceAfter?: number    // twips
  lineSpacing?: number
  lineRule?: string
  indentLeft?: number
  indentRight?: number
  firstLineIndent?: number       // twips (1/20 pt)
  firstLineIndentChars?: number  // hundredths of a character; auto-scales with font size
  hangingIndent?: number         // twips
  hangingIndentChars?: number    // hundredths of a character
  outlineLevel?: number
  numId?: string
  numLevel?: number
}

export type TableClassification = "layout" | "data" | "form"

export interface ElementInfo {
  type:
    | "paragraph"
    | "table"
    | "image"
    | "equation"
    | "sectionBreak"
    | "pageBreak"
    | "empty"
  detail?: string
}

export interface ParsedParagraph {
  index: number
  text: string
  rPr: ComputedRunStyle
  pPr: ComputedParaStyle
  styleId: string
  styleName: string
  fingerprint: string
  context: {
    predecessor: ElementInfo | null
    successor: ElementInfo | null
    insideTable: TableClassification | null
    sectionIndex: number
    /**
     * Nearest image / table within ~3 non-paragraph elements before this
     * paragraph (or N intervening paragraphs as configured). Used by the
     * agent to distinguish "table caption above table" (nearestTableAfter
     * is set) from "figure caption below image" (nearestImageBefore is
     * set) without inspecting raw XML.
     */
    nearestImageBefore: { distance: number; widthCm: number; heightCm: number } | null
    nearestImageAfter: { distance: number; widthCm: number; heightCm: number } | null
    nearestTableBefore: { distance: number; rows: number; cols: number } | null
    nearestTableAfter: { distance: number; rows: number; cols: number } | null
  }
}

export interface StyleDefinition {
  id: string
  name: string
  type: string
  basedOn: string | null
  rPr: ComputedRunStyle
  pPr: ComputedParaStyle
  isDefault: boolean
  usageCount: number
}

export interface NumberingLevel {
  level: number
  format: string
  text: string
  pStyle?: string
  start: number
}

export interface NumberingDefinition {
  numId: string
  abstractNumId: string
  levels: NumberingLevel[]
}

export interface SectionInfo {
  index: number
  paraRange: [number, number]
  pageSize: { width: number; height: number }
  margins: { top: number; bottom: number; left: number; right: number }
  orientation: "portrait" | "landscape"
  header: string | null
  footer: string | null
  headerHasImage: boolean
  footerPageNumFormat?: string
  headerFontInfo?: string
  footerFontInfo?: string
}

export type DocumentElement =
  | { kind: "paragraph"; paragraph: ParsedParagraph }
  | { kind: "image"; widthCm: number; heightCm: number; sectionIndex: number }
  | {
      kind: "table"
      classification: TableClassification
      rows: number
      cols: number
      headers: string[]
      sectionIndex: number
      paragraphs: ParsedParagraph[] // only populated for layout tables
    }
  | { kind: "equation"; sectionIndex: number }
  | { kind: "pageBreak"; sectionIndex: number }
  | { kind: "sectionBreak"; sectionIndex: number }
  | { kind: "emptyRun"; count: number; firstIndex: number; sectionIndex: number }

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  dc: "http://purl.org/dc/elements/1.1/",
} as const
