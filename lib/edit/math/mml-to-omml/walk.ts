/**
 * Main MathML → OMML walker.
 *
 * `emitChildren` is the recursion point: it iterates over an mrow-like
 * scope, applies n-ary fusion lookahead, then dispatches each remaining
 * element through `emitElement`. Per-element emitters call back into
 * `emitChildren` for their own sub-scopes.
 *
 * Element mappings ported from TEI Stylesheets' mml2omml.xsl (BSD-2 /
 * CC-BY-SA, © 2011–2020 TEI Consortium). N-ary fusion and the
 * transparent passthrough for mpadded/mstyle are net-new — neither TEI
 * nor mml2omml gets those right.
 */

import {
  elementChildren,
  isMmlElement,
  mmlText,
  mmlTextLiteral,
  mEl,
  setMVal,
  attr,
} from "./dom.ts"
import { ACCENT_CHARS, BAR_OVER_CHARS, GROUP_CHR_MAP } from "./constants.ts"
import { buildRun } from "./run.ts"
import { detectNary, type NaryMatch } from "./nary.ts"
import type { LeafKind } from "./style.ts"

/** Emit the OMML children of `mmlParent` into `ommlParent`, applying
 *  n-ary fusion at mrow scope. Single-child transparent wrappers
 *  (mrow / mstyle / mpadded / menclose / maction with one element
 *  child) are flattened first — temml emits a wrapping <mrow> around
 *  every operator-with-limits, which would otherwise hide the n-ary
 *  head from its operand siblings and reproduce the empty-<m:e/> bug. */
export function emitChildren(mmlParent: Element, ommlParent: Element, doc: Document): void {
  const kids = flattenSingleChildWrappers(elementChildren(mmlParent))
  for (let i = 0; i < kids.length; i++) {
    const nary = detectNary(kids, i)
    if (nary !== null) {
      emitNary(nary, ommlParent, doc)
      i += nary.consumed - 1
      continue
    }
    emitElement(kids[i]!, ommlParent, doc)
  }
}

const TRANSPARENT_WRAPPERS: ReadonlySet<string> = new Set([
  "mrow",
  "mstyle",
  "mpadded",
  "menclose",
  "maction",
])

function flattenSingleChildWrappers(kids: Element[]): Element[] {
  const out: Element[] = []
  for (const k of kids) {
    if (TRANSPARENT_WRAPPERS.has(k.localName) && elementChildren(k).length === 1) {
      out.push(...flattenSingleChildWrappers(elementChildren(k)))
    } else {
      out.push(k)
    }
  }
  return out
}

/** Emit a list of MathML elements into an OMML container, with n-ary
 *  fusion + single-child wrapper flattening applied to the list (same
 *  as emitChildren but the kids come from a passed array rather than
 *  an Element's children — used for n-ary operand/limit recursion). */
function emitInto(host: Element, items: Element[], doc: Document): void {
  const kids = flattenSingleChildWrappers(items)
  for (let i = 0; i < kids.length; i++) {
    const nary = detectNary(kids, i)
    if (nary !== null) {
      emitNary(nary, host, doc)
      i += nary.consumed - 1
      continue
    }
    emitElement(kids[i]!, host, doc)
  }
}

function emitNary(m: NaryMatch, parent: Element, doc: Document): void {
  const nary = mEl(doc, "nary")
  const naryPr = mEl(doc, "naryPr")

  const chr = mEl(doc, "chr")
  setMVal(chr, m.chr)
  naryPr.appendChild(chr)

  const limLoc = mEl(doc, "limLoc")
  setMVal(limLoc, m.limitsAboveBelow ? "undOvr" : "subSup")
  naryPr.appendChild(limLoc)

  if (m.sub === null) {
    const subHide = mEl(doc, "subHide")
    setMVal(subHide, "1")
    naryPr.appendChild(subHide)
  }
  if (m.sup === null) {
    const supHide = mEl(doc, "supHide")
    setMVal(supHide, "1")
    naryPr.appendChild(supHide)
  }
  nary.appendChild(naryPr)

  const sub = mEl(doc, "sub")
  if (m.sub !== null) emitInto(sub, m.sub, doc)
  nary.appendChild(sub)

  const sup = mEl(doc, "sup")
  if (m.sup !== null) emitInto(sup, m.sup, doc)
  nary.appendChild(sup)

  const e = mEl(doc, "e")
  emitInto(e, m.operand, doc)
  nary.appendChild(e)

  parent.appendChild(nary)
}

/** Dispatch one MathML element to its OMML emitter and append into
 *  `parent`. */
export function emitElement(el: Element, parent: Element, doc: Document): void {
  const ln = el.localName
  switch (ln) {
    case "mi":
    case "mn":
    case "mo":
    case "mtext":
    case "ms":
      parent.appendChild(emitLeaf(el, ln as LeafKind, doc))
      return
    case "mspace":
      // mspace → an empty run with a single space. Word collapses
      // empty <m:r> blocks, so emit nothing for zero-width spaces and
      // a thin space otherwise. Simpler heuristic: always one space;
      // refine if a fixture needs more.
      parent.appendChild(buildRun(doc, " ", "mspace", undefined))
      return
    case "mrow":
      // Stretchy-fence detection: temml emits \left(…\right) as an
      // mrow with `<mo fence form=prefix stretchy>` first and matching
      // postfix mo last. Without recognizing this we'd render the
      // parens as ordinary runs — Word doesn't scale them with the
      // body. Fuse the whole thing into <m:d>.
      if (isFenceMrow(el)) {
        parent.appendChild(emitStretchyFence(el, doc))
        return
      }
      // Plain mrow — emit children directly. Fusion (n-ary, nested
      // fence) runs against the inner list.
      emitChildren(el, parent, doc)
      return
    case "mpadded":
    case "mstyle":
    case "menclose":
    case "maction":
      // Transparent passthrough — these only affect rendering nuances
      // (spacing, color, click handling) that have no clean OMML
      // counterpart. Dropping the wrapper keeps the operand visible,
      // which is the right tradeoff vs mml2omml's silent-drop and
      // TEI's invisible-phantom paths.
      emitChildren(el, parent, doc)
      return
    case "mphantom": {
      const phant = mEl(doc, "phant")
      const phantE = mEl(doc, "e")
      emitChildren(el, phantE, doc)
      phant.appendChild(phantE)
      parent.appendChild(phant)
      return
    }
    case "msub":
      parent.appendChild(emitSubSup(el, doc, "sub"))
      return
    case "msup":
      parent.appendChild(emitSubSup(el, doc, "sup"))
      return
    case "msubsup":
      parent.appendChild(emitSubSup(el, doc, "both"))
      return
    case "munder":
      parent.appendChild(emitUnderOver(el, doc, "under"))
      return
    case "mover":
      parent.appendChild(emitUnderOver(el, doc, "over"))
      return
    case "munderover":
      parent.appendChild(emitUnderOver(el, doc, "both"))
      return
    case "mfrac":
      parent.appendChild(emitFrac(el, doc))
      return
    case "msqrt":
      parent.appendChild(emitSqrt(el, doc))
      return
    case "mroot":
      parent.appendChild(emitRoot(el, doc))
      return
    case "mfenced":
      parent.appendChild(emitFenced(el, doc))
      return
    case "mtable":
      parent.appendChild(emitTable(el, doc))
      return
    case "merror":
      throw new Error(
        `MathML <merror> in input — temml hit an error rendering this LaTeX. ` +
          `Inspect the operand: ${(el.textContent ?? "").slice(0, 120)}`,
      )
    case "semantics":
      // <semantics> wraps presentation MathML + annotation(s). Take the
      // first child (presentation), discard annotations.
      {
        const first = elementChildren(el)[0]
        if (first) emitElement(first, parent, doc)
      }
      return
    case "annotation":
    case "annotation-xml":
      // Annotations carry the original LaTeX (temml puts it there).
      // Drop on the OMML side.
      return
    default:
      throw new Error(
        `Unsupported MathML element <${el.localName}> in MathML→OMML conversion. ` +
          `Switch this equation to the omml escape hatch on the EquationBlock; ` +
          `see references/equations.md "Known fragile LaTeX tokens".`,
      )
  }
}

function emitLeaf(el: Element, kind: LeafKind, doc: Document): Element {
  // mtext / ms preserve whitespace (MathML 3 §3.2.6) — leading/trailing
  // spaces in `\text{ ... }` carry the only typographic gap to the
  // adjacent math runs. mi/mn/mo are stylized identifiers/numbers/ops
  // where whitespace is incidental.
  const text = kind === "mtext" || kind === "ms" ? mmlTextLiteral(el) : mmlText(el)
  return buildRun(doc, text, kind, attr(el, "mathvariant"))
}

/** msub / msup / msubsup → <m:sSub> / <m:sSup> / <m:sSubSup>. */
function emitSubSup(el: Element, doc: Document, which: "sub" | "sup" | "both"): Element {
  const kids = elementChildren(el)
  const tag = which === "sub" ? "sSub" : which === "sup" ? "sSup" : "sSubSup"
  const out = mEl(doc, tag)
  const e = mEl(doc, "e")
  emitElement(kids[0]!, e, doc)
  out.appendChild(e)
  if (which === "sub" || which === "both") {
    const sub = mEl(doc, "sub")
    emitElement(kids[1]!, sub, doc)
    out.appendChild(sub)
  }
  if (which === "sup") {
    const sup = mEl(doc, "sup")
    emitElement(kids[1]!, sup, doc)
    out.appendChild(sup)
  }
  if (which === "both") {
    const sup = mEl(doc, "sup")
    emitElement(kids[2]!, sup, doc)
    out.appendChild(sup)
  }
  return out
}

/** munder / mover / munderover → <m:limLow> / <m:limUpp> / <m:limLowUpp>
 *  when the over/under is an accent character → <m:acc>;
 *  when over is bar-class → <m:bar>;
 *  otherwise the limit/limit-low-upp shapes. */
function emitUnderOver(el: Element, doc: Document, which: "under" | "over" | "both"): Element {
  const kids = elementChildren(el)
  // \overbrace / \underbrace / paren / square-bracket grouping →
  // <m:groupChr>. Detect on both over and under positions.
  if (which === "over" || which === "under") {
    const decorator = kids[1]!
    const decText = mmlText(decorator)
    const pos = GROUP_CHR_MAP.get(decText)
    if (pos !== undefined && (which === "over" ? pos === "top" : pos === "bot")) {
      return emitGroupChr(kids[0]!, decText, pos, doc)
    }
  }
  // accent="true" or accent character → m:acc (mover only)
  if (which === "over") {
    const over = kids[1]!
    const overText = mmlText(over)
    if (BAR_OVER_CHARS.has(overText)) {
      const bar = mEl(doc, "bar")
      const barPr = mEl(doc, "barPr")
      const pos = mEl(doc, "pos")
      setMVal(pos, "top")
      barPr.appendChild(pos)
      bar.appendChild(barPr)
      const e = mEl(doc, "e")
      emitElement(kids[0]!, e, doc)
      bar.appendChild(e)
      return bar
    }
    // LaTeX accents (\hat, \vec, \tilde, \dot, ...) all render as
    // <mover> with a single-character <mo> over the base. Treat any
    // single-char operator in over position as an accent — matches
    // the LaTeX-side intent and avoids enumerating every accent
    // codepoint. accent="true" attribute (set by some authoring
    // tools) is also honored.
    const explicit = attr(el, "accent") === "true"
    const isSingleCharMo = isMmlElement(over, "mo") && [...overText].length === 1
    if (explicit || ACCENT_CHARS.has(overText) || isSingleCharMo) {
      const acc = mEl(doc, "acc")
      const accPr = mEl(doc, "accPr")
      const chr = mEl(doc, "chr")
      setMVal(chr, overText)
      accPr.appendChild(chr)
      acc.appendChild(accPr)
      const e = mEl(doc, "e")
      emitElement(kids[0]!, e, doc)
      acc.appendChild(e)
      return acc
    }
  }
  if (which === "under") {
    const under = kids[1]!
    if (BAR_OVER_CHARS.has(mmlText(under))) {
      const bar = mEl(doc, "bar")
      const barPr = mEl(doc, "barPr")
      const pos = mEl(doc, "pos")
      setMVal(pos, "bot")
      barPr.appendChild(pos)
      bar.appendChild(barPr)
      const e = mEl(doc, "e")
      emitElement(kids[0]!, e, doc)
      bar.appendChild(e)
      return bar
    }
  }
  // Fall through to limit shapes.
  const tag = which === "under" ? "limLow" : which === "over" ? "limUpp" : "limLowUpp"
  const out = mEl(doc, tag)
  const e = mEl(doc, "e")
  emitElement(kids[0]!, e, doc)
  out.appendChild(e)
  if (which === "under" || which === "both") {
    const lim = mEl(doc, "lim")
    emitElement(kids[1]!, lim, doc)
    out.appendChild(lim)
  }
  if (which === "over") {
    const lim = mEl(doc, "lim")
    emitElement(kids[1]!, lim, doc)
    out.appendChild(lim)
  }
  if (which === "both") {
    const lim = mEl(doc, "lim")
    emitElement(kids[2]!, lim, doc)
    out.appendChild(lim)
  }
  return out
}

function emitGroupChr(base: Element, chr: string, pos: "top" | "bot", doc: Document): Element {
  const g = mEl(doc, "groupChr")
  const gPr = mEl(doc, "groupChrPr")
  const chrEl = mEl(doc, "chr")
  setMVal(chrEl, chr)
  gPr.appendChild(chrEl)
  const posEl = mEl(doc, "pos")
  setMVal(posEl, pos)
  gPr.appendChild(posEl)
  // m:vertJc controls which end of the base the chr attaches to.
  // "bot" → bracket below base (underbrace); "top" → above (overbrace).
  const vertJc = mEl(doc, "vertJc")
  setMVal(vertJc, pos)
  gPr.appendChild(vertJc)
  g.appendChild(gPr)
  const e = mEl(doc, "e")
  emitElement(base, e, doc)
  g.appendChild(e)
  return g
}

function emitFrac(el: Element, doc: Document): Element {
  const kids = elementChildren(el)
  const f = mEl(doc, "f")
  const fPr = mEl(doc, "fPr")
  const linethickness = attr(el, "linethickness")
  if (linethickness === "0" || linethickness === "0pt") {
    const type = mEl(doc, "type")
    setMVal(type, "noBar")
    fPr.appendChild(type)
    f.appendChild(fPr)
  }
  const num = mEl(doc, "num")
  emitElement(kids[0]!, num, doc)
  f.appendChild(num)
  const den = mEl(doc, "den")
  emitElement(kids[1]!, den, doc)
  f.appendChild(den)
  return f
}

function emitSqrt(el: Element, doc: Document): Element {
  const rad = mEl(doc, "rad")
  const radPr = mEl(doc, "radPr")
  const degHide = mEl(doc, "degHide")
  setMVal(degHide, "1")
  radPr.appendChild(degHide)
  rad.appendChild(radPr)
  rad.appendChild(mEl(doc, "deg"))
  const e = mEl(doc, "e")
  emitChildren(el, e, doc)
  rad.appendChild(e)
  return rad
}

function emitRoot(el: Element, doc: Document): Element {
  const kids = elementChildren(el)
  const rad = mEl(doc, "rad")
  const deg = mEl(doc, "deg")
  emitElement(kids[1]!, deg, doc)
  rad.appendChild(deg)
  const e = mEl(doc, "e")
  emitElement(kids[0]!, e, doc)
  rad.appendChild(e)
  return rad
}

function emitFenced(el: Element, doc: Document): Element {
  const open = attr(el, "open") ?? "("
  const close = attr(el, "close") ?? ")"
  const separators = attr(el, "separators") ?? ","
  const d = mEl(doc, "d")
  const dPr = mEl(doc, "dPr")
  const begChr = mEl(doc, "begChr")
  setMVal(begChr, open)
  dPr.appendChild(begChr)
  const endChr = mEl(doc, "endChr")
  setMVal(endChr, close)
  dPr.appendChild(endChr)
  if (separators !== ",") {
    const sepChr = mEl(doc, "sepChr")
    setMVal(sepChr, separators.charAt(0))
    dPr.appendChild(sepChr)
  }
  d.appendChild(dPr)
  for (const child of elementChildren(el)) {
    const e = mEl(doc, "e")
    emitElement(child, e, doc)
    d.appendChild(e)
  }
  return d
}

function isFenceMrow(mrow: Element): boolean {
  // Detects both stretchy `\left…\right` and non-stretchy `\binom`-style
  // pairs. The fence attribute alone is enough — temml sets it on every
  // matched delimiter. Stretchy vs not is signaled by `m:grow` in the
  // OMML output (we always emit grow; renderers ignore it when the body
  // is short).
  const kids = elementChildren(mrow)
  if (kids.length < 2) return false
  const first = kids[0]!
  const last = kids[kids.length - 1]!
  return (
    isMmlElement(first, "mo") &&
    attr(first, "fence") === "true" &&
    isMmlElement(last, "mo") &&
    attr(last, "fence") === "true"
  )
}

function emitStretchyFence(mrow: Element, doc: Document): Element {
  const kids = elementChildren(mrow)
  const open = kids[0]!
  const close = kids[kids.length - 1]!
  const body = kids.slice(1, -1)

  const d = mEl(doc, "d")
  const dPr = mEl(doc, "dPr")
  // For `\left.` (invisible delimiter) temml emits an empty <mo></mo>.
  // The empty string must be emitted *explicitly* as `m:val=""` — omitting
  // begChr/endChr falls back to the OMML default `(` / `)`, which renders
  // wrongly for `\begin{cases}` and other half-open shapes (Word/LO would
  // close `{ ... }` with `)` because endChr defaulted).
  const begChr = mEl(doc, "begChr")
  setMVal(begChr, mmlText(open))
  dPr.appendChild(begChr)
  const endChr = mEl(doc, "endChr")
  setMVal(endChr, mmlText(close))
  dPr.appendChild(endChr)
  // Stretchy → m:grow tells Word to scale the delimiter to the body
  // height. Without it, some renderers (LibreOffice) keep the delimiter
  // at base size even when the body is tall.
  dPr.appendChild(mEl(doc, "grow"))
  d.appendChild(dPr)

  // Single <m:e> containing the body. Splitting on `<mo>,</mo>`
  // separators (the OMML m:sepChr pattern) is an optimization for
  // (a,b,c) tuple display — Word renders either form, the single-e
  // form preserves the original spacing more faithfully.
  const e = mEl(doc, "e")
  for (const child of body) emitElement(child, e, doc)
  d.appendChild(e)

  return d
}

function emitTable(el: Element, doc: Document): Element {
  const m = mEl(doc, "m")
  // m:mPr is optional; skip for now (Word will use defaults).
  for (const tr of elementChildren(el)) {
    if (!isMmlElement(tr, "mtr") && !isMmlElement(tr, "mlabeledtr")) continue
    const mr = mEl(doc, "mr")
    for (const td of elementChildren(tr)) {
      if (!isMmlElement(td, "mtd")) continue
      const e = mEl(doc, "e")
      emitChildren(td, e, doc)
      mr.appendChild(e)
    }
    m.appendChild(mr)
  }
  return m
}
