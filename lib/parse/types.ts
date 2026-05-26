export interface ComputedRunStyle {
  fontAscii?: string
  fontEastAsia?: string
  fontHAnsi?: string
  size?: number // half-points
  bold?: boolean
  italic?: boolean
  underline?: string
  color?: string // hex without # or "auto"
  highlight?: string
  strike?: boolean
  caps?: boolean
  /** OOXML `<w:vertAlign>` baseline shift. "baseline" is the explicit
   * reset — distinct from undefined (= inherit / unspecified). */
  vertAlign?: "superscript" | "subscript" | "baseline"
}

export interface ComputedParaStyle {
  pStyle?: string
  alignment?: string
  spaceBefore?: number // twips
  spaceAfter?: number // twips
  lineSpacing?: number
  lineRule?: string
  indentLeft?: number
  indentRight?: number
  firstLineIndent?: number // twips (1/20 pt)
  firstLineIndentChars?: number // hundredths of a character; auto-scales with font size
  hangingIndent?: number // twips
  hangingIndentChars?: number // hundredths of a character
  outlineLevel?: number
  numId?: string
  numLevel?: number
}

export type TableClassification = "layout" | "data"

export interface ParsedParagraph {
  index: number
  text: string
  rPr: ComputedRunStyle
  pPr: ComputedParaStyle
  /** Direct run-level rPr captured at parse time — what's literally on the
   * paragraph's runs (dominant) and/or pMark rPr, BEFORE the cascade is
   * merged. Used by the vs-target-direct dry-run check to classify each
   * agent-declared style field as override / redundant / new against what
   * the paragraph carries directly. Empty when no direct rPr anywhere. */
  directRPr: ComputedRunStyle
  /** Direct pPr from the paragraph element itself (excluding pMark rPr).
   * Same purpose as directRPr, for paragraph-level direct properties. */
  directPPr: ComputedParaStyle
  styleId: string
  styleName: string
  fingerprint: string
  context: {
    insideTable: TableClassification | null
    sectionIndex: number
  }
}

/**
 * Flat ordered list of every element in the document body — paragraphs
 * (compressed empty runs are NOT collapsed here; each empty paragraph is
 * its own entry), images, tables, equations, breaks. Exposed to
 * inspect_neighbors so neighbor lookup is on-demand instead of an
 * always-on annotation on every paragraph. Layout-table inner paragraphs
 * are inlined; data table inner paragraphs are NOT.
 */
export type NeighborItem =
  | { kind: "paragraph"; paraIndex: number; isEmpty: boolean; sectionIndex: number }
  | { kind: "image"; widthCm: number; heightCm: number; sectionIndex: number }
  | {
      kind: "table"
      classification: TableClassification
      rows: number
      cols: number
      sectionIndex: number
    }
  | { kind: "equation"; sectionIndex: number }
  | { kind: "pageBreak"; sectionIndex: number }
  | { kind: "sectionBreak"; sectionIndex: number }

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

export interface SectionInfo {
  /** 0-based internal section number (document order). Display layers add 1
   * to match agent-facing "Section N" labels; do NOT surface this raw value
   * through any tool output or locator field. */
  index: number
  /** 1-based [from, to], inclusive — matches `#NNN` in overview. */
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
      firstRowLooksLikeHeader: boolean
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
