import {
  NS,
  type ComputedParaStyle,
  type ComputedRunStyle,
  type StyleDefinition,
} from "./types.ts"
import {
  attr,
  firstChildNS,
  getChildrenNS,
  parseToggle,
  textContent,
  wAttr,
  wVal,
} from "./xml-utils.ts"

interface ThemeFonts {
  majorLatin?: string
  majorEastAsia?: string
  minorLatin?: string
  minorEastAsia?: string
}

/** Map of theme color slot (dk1, lt1, accent1, ...) → hex. */
interface ThemeColors {
  [slot: string]: string
}

/** Style entry as parsed from styles.xml — basedOn chain not yet resolved. */
interface RawStyle {
  id: string
  name: string
  type: string
  basedOn: string | null
  rPr: ComputedRunStyle
  pPr: ComputedParaStyle
  isDefault: boolean
}

interface ResolvedStyle {
  rPr: ComputedRunStyle
  pPr: ComputedParaStyle
  styleName: string
  chain: string[]
}

export class StyleResolver {
  private docDefaultsRPr: ComputedRunStyle = {}
  private docDefaultsPPr: ComputedParaStyle = {}
  private styles = new Map<string, StyleDefinition>()
  private themeFonts: ThemeFonts = {}
  private themeColors: ThemeColors = {}
  private rawStyles = new Map<string, RawStyle>()
  private resolvedCache = new Map<string, ResolvedStyle>()

  constructor(stylesDoc: Document | null, themeDoc: Document | null) {
    if (themeDoc) this.parseTheme(themeDoc)
    if (stylesDoc) this.parseStyles(stylesDoc)
  }

  private parseTheme(theme: Document): void {
    const root = theme.documentElement
    if (!root) return
    const themeElements = root.getElementsByTagNameNS(NS.a, "themeElements")[0]
    if (!themeElements) return
    const fontScheme = themeElements.getElementsByTagNameNS(NS.a, "fontScheme")[0]
    if (fontScheme) {
      const major = fontScheme.getElementsByTagNameNS(NS.a, "majorFont")[0]
      const minor = fontScheme.getElementsByTagNameNS(NS.a, "minorFont")[0]
      if (major) {
        const latin = major.getElementsByTagNameNS(NS.a, "latin")[0]
        const ea = major.getElementsByTagNameNS(NS.a, "ea")[0]
        if (latin) this.themeFonts.majorLatin = attr(latin, "", "typeface") || undefined
        if (ea) this.themeFonts.majorEastAsia = attr(ea, "", "typeface") || undefined
      }
      if (minor) {
        const latin = minor.getElementsByTagNameNS(NS.a, "latin")[0]
        const ea = minor.getElementsByTagNameNS(NS.a, "ea")[0]
        if (latin) this.themeFonts.minorLatin = attr(latin, "", "typeface") || undefined
        if (ea) this.themeFonts.minorEastAsia = attr(ea, "", "typeface") || undefined
      }
    }
    const clrScheme = themeElements.getElementsByTagNameNS(NS.a, "clrScheme")[0]
    if (clrScheme) {
      const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
      for (const slot of slots) {
        const el = clrScheme.getElementsByTagNameNS(NS.a, slot)[0]
        if (!el) continue
        const srgb = el.getElementsByTagNameNS(NS.a, "srgbClr")[0]
        const sys = el.getElementsByTagNameNS(NS.a, "sysClr")[0]
        if (srgb) {
          const v = attr(srgb, "", "val")
          if (v) this.themeColors[slot] = v.toUpperCase()
        } else if (sys) {
          const v = attr(sys, "", "lastClr") || attr(sys, "", "val")
          if (v) this.themeColors[slot] = v.toUpperCase()
        }
      }
    }
  }

  private parseStyles(stylesDoc: Document): void {
    const root = stylesDoc.documentElement
    if (!root) return
    const docDefaults = firstChildNS(root, NS.w, "docDefaults")
    if (docDefaults) {
      const rPrDefault = firstChildNS(docDefaults, NS.w, "rPrDefault")
      if (rPrDefault) {
        const rPr = firstChildNS(rPrDefault, NS.w, "rPr")
        this.docDefaultsRPr = this.parseRPr(rPr)
      }
      const pPrDefault = firstChildNS(docDefaults, NS.w, "pPrDefault")
      if (pPrDefault) {
        const pPr = firstChildNS(pPrDefault, NS.w, "pPr")
        this.docDefaultsPPr = this.parsePPr(pPr)
      }
    }

    for (const styleEl of getChildrenNS(root, NS.w, "style")) {
      const id = wAttr(styleEl, "styleId") || ""
      if (!id) continue
      const type = wAttr(styleEl, "type") || "paragraph"
      const isDefault = wAttr(styleEl, "default") === "1"
      const nameEl = firstChildNS(styleEl, NS.w, "name")
      const name = wVal(nameEl) || id
      const basedOnEl = firstChildNS(styleEl, NS.w, "basedOn")
      const basedOn = basedOnEl ? wVal(basedOnEl) : null
      const rPr = this.parseRPr(firstChildNS(styleEl, NS.w, "rPr"))
      const pPr = this.parsePPr(firstChildNS(styleEl, NS.w, "pPr"))

      this.rawStyles.set(id, {
        id,
        name,
        type,
        basedOn,
        rPr,
        pPr,
        isDefault,
      })
    }

    // build StyleDefinitions (raw) — usageCount filled later by parser
    for (const [id, raw] of this.rawStyles) {
      this.styles.set(id, {
        id,
        name: raw.name,
        type: raw.type,
        basedOn: raw.basedOn,
        rPr: raw.rPr,
        pPr: raw.pPr,
        isDefault: raw.isDefault,
        usageCount: 0,
      })
    }
  }

  parseRPr(rPrEl: Element | null): ComputedRunStyle {
    const out: ComputedRunStyle = {}
    if (!rPrEl) return out
    const rFonts = firstChildNS(rPrEl, NS.w, "rFonts")
    if (rFonts) {
      const ascii = wAttr(rFonts, "ascii")
      const hAnsi = wAttr(rFonts, "hAnsi")
      const eastAsia = wAttr(rFonts, "eastAsia")
      const asciiTheme = wAttr(rFonts, "asciiTheme")
      const hAnsiTheme = wAttr(rFonts, "hAnsiTheme")
      const eastAsiaTheme = wAttr(rFonts, "eastAsiaTheme")
      if (ascii) out.fontAscii = ascii
      if (hAnsi) out.fontHAnsi = hAnsi
      if (eastAsia) out.fontEastAsia = eastAsia
      if (!ascii && asciiTheme) out.fontAscii = this.resolveThemeFont(asciiTheme, false)
      if (!hAnsi && hAnsiTheme) out.fontHAnsi = this.resolveThemeFont(hAnsiTheme, false)
      if (!eastAsia && eastAsiaTheme)
        out.fontEastAsia = this.resolveThemeFont(eastAsiaTheme, true)
    }
    const sz = firstChildNS(rPrEl, NS.w, "sz")
    if (sz) {
      const v = wVal(sz)
      if (v) out.size = parseInt(v, 10)
    }
    const szCs = firstChildNS(rPrEl, NS.w, "szCs")
    if (szCs && out.size === undefined) {
      const v = wVal(szCs)
      if (v) out.size = parseInt(v, 10)
    }
    const b = firstChildNS(rPrEl, NS.w, "b")
    const bCs = firstChildNS(rPrEl, NS.w, "bCs")
    const bToggle = parseToggle(b) ?? parseToggle(bCs)
    if (bToggle !== undefined) out.bold = bToggle
    const i = firstChildNS(rPrEl, NS.w, "i")
    const iCs = firstChildNS(rPrEl, NS.w, "iCs")
    const iToggle = parseToggle(i) ?? parseToggle(iCs)
    if (iToggle !== undefined) out.italic = iToggle
    const u = firstChildNS(rPrEl, NS.w, "u")
    if (u) {
      out.underline = wVal(u) || "single"
    }
    const color = firstChildNS(rPrEl, NS.w, "color")
    if (color) {
      const themeColor = wAttr(color, "themeColor")
      if (themeColor) {
        out.color = this.resolveThemeColor(themeColor) ?? wVal(color) ?? "auto"
      } else {
        const v = wVal(color)
        if (v) out.color = v
      }
    }
    const highlight = firstChildNS(rPrEl, NS.w, "highlight")
    if (highlight) out.highlight = wVal(highlight) || undefined
    const strike = firstChildNS(rPrEl, NS.w, "strike")
    const strikeToggle = parseToggle(strike)
    if (strikeToggle !== undefined) out.strike = strikeToggle
    const caps = firstChildNS(rPrEl, NS.w, "caps")
    const capsToggle = parseToggle(caps)
    if (capsToggle !== undefined) out.caps = capsToggle
    return out
  }

  parsePPr(pPrEl: Element | null): ComputedParaStyle {
    const out: ComputedParaStyle = {}
    if (!pPrEl) return out
    const pStyle = firstChildNS(pPrEl, NS.w, "pStyle")
    if (pStyle) {
      const v = wVal(pStyle)
      if (v) out.pStyle = v
    }
    const jc = firstChildNS(pPrEl, NS.w, "jc")
    if (jc) {
      const v = wVal(jc)
      if (v) out.alignment = v
    }
    const spacing = firstChildNS(pPrEl, NS.w, "spacing")
    if (spacing) {
      const before = wAttr(spacing, "before")
      const after = wAttr(spacing, "after")
      const line = wAttr(spacing, "line")
      const lineRule = wAttr(spacing, "lineRule")
      if (before) out.spaceBefore = parseInt(before, 10)
      if (after) out.spaceAfter = parseInt(after, 10)
      if (line) out.lineSpacing = parseInt(line, 10)
      if (lineRule) out.lineRule = lineRule
    }
    const ind = firstChildNS(pPrEl, NS.w, "ind")
    if (ind) {
      const left = wAttr(ind, "left") ?? wAttr(ind, "start")
      const right = wAttr(ind, "right") ?? wAttr(ind, "end")
      const firstLine = wAttr(ind, "firstLine")
      const hanging = wAttr(ind, "hanging")
      const firstLineChars = wAttr(ind, "firstLineChars")
      const hangingChars = wAttr(ind, "hangingChars")
      if (left) out.indentLeft = parseInt(left, 10)
      if (right) out.indentRight = parseInt(right, 10)
      // Preserve both representations independently. Word's `firstLineChars`
      // (hundredths of a character) auto-scales with the run font size at
      // render time; the legacy fixed-twip `firstLine` does not. Folding one
      // into the other (as the previous code did with a hard-coded 240
      // twips/char assumption) silently destroyed the character semantics for
      // any font size other than 12pt.
      if (firstLine) out.firstLineIndent = parseInt(firstLine, 10)
      if (firstLineChars) out.firstLineIndentChars = parseInt(firstLineChars, 10)
      if (hanging) out.hangingIndent = parseInt(hanging, 10)
      if (hangingChars) out.hangingIndentChars = parseInt(hangingChars, 10)
    }
    const outlineLvl = firstChildNS(pPrEl, NS.w, "outlineLvl")
    if (outlineLvl) {
      const v = wVal(outlineLvl)
      if (v !== null) out.outlineLevel = parseInt(v, 10)
    }
    const numPr = firstChildNS(pPrEl, NS.w, "numPr")
    if (numPr) {
      const numId = firstChildNS(numPr, NS.w, "numId")
      const ilvl = firstChildNS(numPr, NS.w, "ilvl")
      if (numId) {
        const v = wVal(numId)
        if (v) out.numId = v
      }
      if (ilvl) {
        const v = wVal(ilvl)
        if (v !== null) out.numLevel = parseInt(v, 10)
      }
    }
    return out
  }

  private resolveThemeFont(themeRef: string, isEastAsia: boolean): string | undefined {
    // OOXML theme font tokens are a closed set (ECMA-376 §17.18.96).
    // Match the exact token rather than fuzzy substring checks — guards
    // against accidental matches in unrelated values.
    switch (themeRef) {
      case "majorAscii":
      case "majorHAnsi":
      case "majorBidi":
        return this.themeFonts.majorLatin
      case "minorAscii":
      case "minorHAnsi":
      case "minorBidi":
        return this.themeFonts.minorLatin
      case "majorEastAsia":
        return this.themeFonts.majorEastAsia
      case "minorEastAsia":
        return this.themeFonts.minorEastAsia
    }
    // Unknown token: fall back to the slot the caller asked for. Major vs
    // minor is no longer derivable, so default to minor (the more common
    // body-text slot).
    return isEastAsia ? this.themeFonts.minorEastAsia : this.themeFonts.minorLatin
  }

  private resolveThemeColor(themeRef: string): string | undefined {
    // Map dark1/light1 → dk1/lt1
    const map: Record<string, string> = {
      dark1: "dk1",
      light1: "lt1",
      dark2: "dk2",
      light2: "lt2",
      hyperlink: "hlink",
      followedHyperlink: "folHlink",
      background1: "lt1",
      background2: "lt2",
      text1: "dk1",
      text2: "dk2",
    }
    const slot = map[themeRef] || themeRef
    return this.themeColors[slot]
  }

  /** Compute final style for a paragraph + optional inline run rPr. */
  computeRunStyle(
    paraStyleId: string,
    runRPrElement: Element | null,
  ): {
    rPr: ComputedRunStyle
    pPr: ComputedParaStyle
    styleName: string
    styleChain: string[]
  } {
    const styleResolved = this.resolveStyleChain(paraStyleId)
    let rPr = { ...styleResolved.rPr }
    let pPr = { ...styleResolved.pPr }
    if (runRPrElement) {
      // direct rPr: parse and overlay
      const direct = this.parseRPr(runRPrElement)
      // rStyle inside direct: a character style overlay
      const rStyle = firstChildNS(runRPrElement, NS.w, "rStyle")
      if (rStyle) {
        const csId = wVal(rStyle)
        if (csId) {
          const cs = this.resolveStyleChain(csId)
          rPr = mergeRPr(rPr, cs.rPr)
        }
      }
      rPr = mergeRPr(rPr, direct)
    }
    return {
      rPr,
      pPr,
      styleName: styleResolved.styleName,
      styleChain: styleResolved.chain,
    }
  }

  /** Resolve a style by id, walking basedOn chain from root downwards. */
  resolveStyleChain(styleId: string): ResolvedStyle {
    const cached = this.resolvedCache.get(styleId)
    if (cached) return cached
    const chain = this.buildChain(styleId)
    let rPr: ComputedRunStyle = { ...this.docDefaultsRPr }
    let pPr: ComputedParaStyle = { ...this.docDefaultsPPr }
    let styleName = styleId
    for (const id of chain) {
      const raw = this.rawStyles.get(id)
      if (!raw) continue
      pPr = mergePPr(pPr, raw.pPr)
      rPr = mergeRPr(rPr, raw.rPr)
      styleName = raw.name
    }
    if (chain.length === 0) {
      // unknown style; still return defaults
      styleName = styleId || "Normal"
    }
    const result = { rPr, pPr, styleName, chain }
    this.resolvedCache.set(styleId, result)
    return result
  }

  private buildChain(styleId: string): string[] {
    // Walk basedOn upwards, then reverse to get root → target
    const visited = new Set<string>()
    const path: string[] = []
    let cur: string | null = styleId
    while (cur && !visited.has(cur)) {
      visited.add(cur)
      const raw = this.rawStyles.get(cur)
      if (!raw) break
      path.push(cur)
      cur = raw.basedOn
    }
    return path.reverse()
  }

  getStyleDefinition(styleId: string): StyleDefinition | null {
    return this.styles.get(styleId) || null
  }

  getAllStyles(): StyleDefinition[] {
    return Array.from(this.styles.values())
  }

  getDocDefaults(): { rPr: ComputedRunStyle; pPr: ComputedParaStyle } {
    return { rPr: this.docDefaultsRPr, pPr: this.docDefaultsPPr }
  }

  getThemeFonts(): ThemeFonts {
    return this.themeFonts
  }

  getThemeColors(): ThemeColors {
    return this.themeColors
  }

  incrementUsage(styleId: string): void {
    const def = this.styles.get(styleId)
    if (def) def.usageCount++
  }
}

/**
 * Overlay only the defined fields of `overlay` onto a copy of `base`. Used
 * for merging both rPr and pPr inheritance frames; the field semantics are
 * the same (later wins iff non-undefined) so one helper covers both.
 */
function mergeStyle<T extends object>(base: T, overlay: T): T {
  const out = { ...base }
  for (const k of Object.keys(overlay) as (keyof T)[]) {
    const v = overlay[k]
    if (v !== undefined) (out as any)[k] = v
  }
  return out
}

function mergeRPr(base: ComputedRunStyle, overlay: ComputedRunStyle): ComputedRunStyle {
  return mergeStyle(base, overlay)
}

function mergePPr(base: ComputedParaStyle, overlay: ComputedParaStyle): ComputedParaStyle {
  return mergeStyle(base, overlay)
}
