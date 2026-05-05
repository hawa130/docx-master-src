/**
 * Parses free-form Chinese typographic requirements into StyleConfigEntry
 * field values. Designed to consume what users actually write in school
 * thesis specs / formatting briefs:
 *
 *   "小四宋体首行缩进2字符1.5倍行距"
 *   "二号黑体加粗居中段前12磅段后6磅"
 *   "Times New Roman 10.5pt italic 行距固定值15磅"
 *
 * Handles tokens in any order; longest match wins for ambiguous size names
 * (小四 before 四号, 小初 before 初号). Returns a partial style spec the
 * caller merges into the user's structured config — requirements override
 * fromParagraph extraction (they're authored intent), but don't override
 * explicit `overrides` field (that's a deliberate per-style escape).
 */

export interface ParsedRequirement {
  font?: string
  fontEastAsia?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: string
  color?: string
  alignment?: "left" | "center" | "right" | "both"
  lineSpacing?: number
  lineRule?: "auto" | "exact" | "atLeast"
  spaceBefore?: number
  spaceAfter?: number
  firstLineIndent?: string | number
  hangingIndent?: string | number
  outlineLevel?: number
  /** Tokens that didn't match any known rule. Surfaced in validation. */
  unparsed: string[]
}

const SIZE_NAMES: ReadonlyArray<readonly [string, number]> = [
  // longest first to avoid prefix collision (小初 must beat 小)
  ["小初", 36],
  ["小一", 24],
  ["小二", 18],
  ["小三", 15],
  ["小四", 12],
  ["小五", 9],
  ["小六", 6.5],
  ["小七", 5.5],
  ["初号", 42],
  ["一号", 26],
  ["二号", 22],
  ["三号", 16],
  ["四号", 14],
  ["五号", 10.5],
  ["六号", 7.5],
  ["七号", 5.5],
  ["八号", 5],
]

const FONT_KEYWORDS = [
  "宋体", "黑体", "楷体", "仿宋", "等线", "微软雅黑",
  "华文宋体", "华文楷体", "华文仿宋", "华文黑体", "华文中宋", "华文细黑",
  "华文琥珀", "华文行楷", "华文新魏", "华文隶书",
  "方正小标宋", "方正姚体", "方正书宋",
  "DengXian", "SimSun", "SimHei", "KaiTi", "FangSong", "Microsoft YaHei",
  "Times New Roman", "Times", "Arial", "Calibri", "Cambria", "Helvetica",
  "Courier New", "Courier", "Verdana", "Georgia", "Tahoma",
] as const

const COLOR_NAMES: Record<string, string> = {
  红色: "FF0000",
  红:   "FF0000",
  绿色: "00B050",
  绿:   "00B050",
  蓝色: "0070C0",
  蓝:   "0070C0",
  黑色: "000000",
  黑:   "000000",
  白色: "FFFFFF",
  白:   "FFFFFF",
  灰色: "808080",
  灰:   "808080",
  黄色: "FFC000",
  黄:   "FFC000",
  橙色: "ED7D31",
  橙:   "ED7D31",
  紫色: "7030A0",
  紫:   "7030A0",
  深蓝: "002060",
  浅蓝: "5B9BD5",
}

export function parseRequirement(text: string): ParsedRequirement {
  const result: ParsedRequirement = { unparsed: [] }
  if (!text) return result
  let s = text.trim()

  // Font names (longest first; case-insensitive for Latin names)
  const fontPatterns = [...FONT_KEYWORDS].sort((a, b) => b.length - a.length)
  for (const name of fontPatterns) {
    const re = new RegExp(escapeRegex(name), "i")
    const m = s.match(re)
    if (m) {
      // East-Asian fonts populate both font (latin slot) and fontEastAsia
      // when used alone; the agent can override via explicit overrides if
      // they want different latin/CJK fonts.
      const isCJK = /[一-鿿]/.test(name)
      if (isCJK) {
        result.fontEastAsia ??= name
        result.font ??= name
      } else {
        result.font ??= name
      }
      s = s.replace(re, " ")
    }
  }

  // Size: 字号 keywords first (longest match), then "Npt" / "N磅" / "N号"
  for (const [name, pt] of SIZE_NAMES) {
    if (s.includes(name)) {
      result.size = pt
      s = s.split(name).join(" ")
      break
    }
  }
  if (result.size === undefined) {
    const m = s.match(/(\d+(?:\.\d+)?)\s*(?:pt|磅)/i)
    if (m) {
      result.size = parseFloat(m[1]!)
      s = s.replace(m[0], " ")
    }
  }

  // Weight / italic / underline
  if (/(?:加粗|粗体|\bbold\b)/i.test(s)) {
    result.bold = true
    s = s.replace(/(?:加粗|粗体|\bbold\b)/gi, " ")
  }
  if (/(?:倾斜|斜体|\bitalic\b)/i.test(s)) {
    result.italic = true
    s = s.replace(/(?:倾斜|斜体|\bitalic\b)/gi, " ")
  }
  const underlineMatch = s.match(/(?:下划线|\bunderline\b)/i)
  if (underlineMatch) {
    result.underline = "single"
    s = s.replace(/(?:下划线|\bunderline\b)/gi, " ")
  }

  // Line spacing — try most specific first
  let m: RegExpMatchArray | null
  if ((m = s.match(/(?:固定值|固定行距|行距固定值|行距固定)\s*(\d+(?:\.\d+)?)\s*磅/))) {
    result.lineSpacing = parseFloat(m[1]!)
    result.lineRule = "exact"
    s = s.replace(m[0], " ")
  } else if ((m = s.match(/(?:最小值|至少|atLeast)\s*(\d+(?:\.\d+)?)\s*磅/))) {
    result.lineSpacing = parseFloat(m[1]!)
    result.lineRule = "atLeast"
    s = s.replace(m[0], " ")
  } else if ((m = s.match(/(?:行距)?\s*(\d+(?:\.\d+)?)\s*倍\s*(?:行距)?/))) {
    result.lineSpacing = parseFloat(m[1]!)
    result.lineRule = "auto"
    s = s.replace(m[0], " ")
  } else if (/单倍行距|单倍/.test(s)) {
    result.lineSpacing = 1
    result.lineRule = "auto"
    s = s.replace(/单倍行距|单倍/g, " ")
  } else if (/1\.5\s*倍行距|一点五倍/.test(s)) {
    result.lineSpacing = 1.5
    result.lineRule = "auto"
    s = s.replace(/1\.5\s*倍行距|一点五倍/g, " ")
  } else if (/双倍行距|两倍行距/.test(s)) {
    result.lineSpacing = 2
    result.lineRule = "auto"
    s = s.replace(/双倍行距|两倍行距/g, " ")
  } else if ((m = s.match(/行距\s*(\d+(?:\.\d+)?)\s*磅/))) {
    // bare "行距20磅" → exact line height in pt
    result.lineSpacing = parseFloat(m[1]!)
    result.lineRule = "exact"
    s = s.replace(m[0], " ")
  }

  // Paragraph spacing
  if ((m = s.match(/段前\s*(\d+(?:\.\d+)?)\s*磅/))) {
    result.spaceBefore = parseFloat(m[1]!)
    s = s.replace(m[0], " ")
  }
  if ((m = s.match(/段后\s*(\d+(?:\.\d+)?)\s*磅/))) {
    result.spaceAfter = parseFloat(m[1]!)
    s = s.replace(m[0], " ")
  }
  if ((m = s.match(/段前段后(?:各)?\s*(\d+(?:\.\d+)?)\s*磅/))) {
    result.spaceBefore = parseFloat(m[1]!)
    result.spaceAfter = parseFloat(m[1]!)
    s = s.replace(m[0], " ")
  }

  // Indents — char takes priority over pt (Chinese semantics)
  if ((m = s.match(/首行缩进\s*(\d+(?:\.\d+)?)\s*字符/))) {
    result.firstLineIndent = `${m[1]}char`
    s = s.replace(m[0], " ")
  } else if ((m = s.match(/首行缩进\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)/i))) {
    result.firstLineIndent = `${m[1]}pt`
    s = s.replace(m[0], " ")
  } else if (/无首行缩进|不缩进|首行不缩进/.test(s)) {
    result.firstLineIndent = 0
    s = s.replace(/无首行缩进|不缩进|首行不缩进/g, " ")
  }
  if ((m = s.match(/悬挂缩进\s*(\d+(?:\.\d+)?)\s*字符/))) {
    result.hangingIndent = `${m[1]}char`
    s = s.replace(m[0], " ")
  } else if ((m = s.match(/悬挂缩进\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)/i))) {
    result.hangingIndent = `${m[1]}pt`
    s = s.replace(m[0], " ")
  }

  // Alignment
  if (/居中|水平居中|\bcenter\b/i.test(s)) {
    result.alignment = "center"
    s = s.replace(/居中|水平居中|\bcenter\b/gi, " ")
  } else if (/(?:左对齐|居左|\bleft\b)/i.test(s)) {
    result.alignment = "left"
    s = s.replace(/(?:左对齐|居左|\bleft\b)/gi, " ")
  } else if (/(?:右对齐|居右|\bright\b)/i.test(s)) {
    result.alignment = "right"
    s = s.replace(/(?:右对齐|居右|\bright\b)/gi, " ")
  } else if (/两端对齐|分散对齐|\bjustify\b/i.test(s)) {
    result.alignment = "both"
    s = s.replace(/两端对齐|分散对齐|\bjustify\b/gi, " ")
  }

  // Outline level — supports "大纲级别N" / "outlineLevel N"
  if ((m = s.match(/(?:大纲级别|outlineLevel)\s*(\d)/i))) {
    result.outlineLevel = parseInt(m[1]!, 10)
    s = s.replace(m[0], " ")
  }

  // Color: hex first, then named colors (longest first)
  if ((m = s.match(/#([0-9A-Fa-f]{6})/))) {
    result.color = m[1]!.toUpperCase()
    s = s.replace(m[0], " ")
  } else {
    const colorNames = Object.keys(COLOR_NAMES).sort((a, b) => b.length - a.length)
    for (const name of colorNames) {
      if (s.includes(name)) {
        result.color = COLOR_NAMES[name]!
        s = s.split(name).join(" ")
        break
      }
    }
  }

  // Anything left after stripping known tokens — record so the validator
  // can warn the agent that some user-stated requirement wasn't honored.
  const leftover = s.replace(/[\s，,;；、。·\-/]+/g, " ").trim()
  if (leftover.length > 0) {
    result.unparsed = leftover.split(/\s+/).filter((t) => t.length > 0)
  }

  return result
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}
