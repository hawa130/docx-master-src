/**
 * N-ary operator detection and fusion.
 *
 * Why this exists: temml emits a sum/integral/product as
 *
 *   <mrow>
 *     <munderover><mo>∑</mo><mrow>i=1</mrow><mi>n</mi></munderover>
 *     <msup><mi>i</mi><mn>2</mn></msup>     <-- the operand
 *   </mrow>
 *
 * The operator-with-limits and the operand are *sibling* children of an
 * mrow. Node-by-node converters (mml2omml LGPL, TEI XSLT) translate each
 * sibling independently and emit `<m:nary>` with an empty `<m:e/>` and
 * the operand as a sibling outside the nary — Word renders that as a
 * dashed empty box before the actual operand.
 *
 * Fix: at every mrow scope, scan for an n-ary-bearing element (`mo`,
 * `munder`/`mover`/`munderover` or `msub`/`msup`/`msubsup` whose base is
 * `mo` in the n-ary chr set) and fuse it with the following siblings as
 * operand, up to an infix-binary stop or end of mrow.
 *
 * Returns either null (the element at index `i` is not n-ary) or a
 * NaryMatch describing how many siblings to consume and what their
 * roles are. The walker then emits one `<m:nary>` and skips the consumed
 * indices.
 */

import { isMmlElement, mmlText, elementChildren } from "./dom.ts"
import { NARY_OPERATORS, NARY_OPERAND_TERMINATORS } from "./constants.ts"

export interface NaryMatch {
  /** OMML `<m:chr m:val="…">` value — the n-ary operator symbol. */
  chr: string
  /** Elements forming the lower limit (sub), or null. */
  sub: Element[] | null
  /** Elements forming the upper limit (sup), or null. */
  sup: Element[] | null
  /** Elements forming the operand body. May be empty for `\sum` written
   *  without any operand to its right. */
  operand: Element[]
  /** Number of sibling positions consumed starting at `startIdx`
   *  (always ≥ 1; the n-ary head element itself counts as one). */
  consumed: number
  /** Whether limits render above/below (munderover, default ∑/∏) or
   *  in subscript/superscript position (msubsup, default ∫). Caller
   *  decides the OMML `<m:limLoc>` setting from this. */
  limitsAboveBelow: boolean
}

/** Examine sibling list `siblings` starting at `i`. If `siblings[i]` is
 *  an n-ary head, fuse with following siblings into a NaryMatch.
 *  Returns null when no n-ary at that position. */
export function detectNary(siblings: Element[], i: number): NaryMatch | null {
  const head = siblings[i]
  if (!head) return null

  // Three head shapes we recognize:
  //   - bare mo (no limits)
  //   - munder / mover / munderover with mo base
  //   - msub / msup / msubsup with mo base
  let chr: string | undefined
  let sub: Element[] | null = null
  let sup: Element[] | null = null
  let limitsAboveBelow = false

  if (isMmlElement(head, "mo")) {
    chr = NARY_OPERATORS.get(mmlText(head))
    // Bare n-ary defaults to undOvr placement per Word.
    limitsAboveBelow = true
  } else if (
    isMmlElement(head, "munder") ||
    isMmlElement(head, "mover") ||
    isMmlElement(head, "munderover")
  ) {
    const kids = elementChildren(head)
    const base = kids[0]
    if (base && isMmlElement(base, "mo")) {
      chr = NARY_OPERATORS.get(mmlText(base))
      // m{under,over,underover} → limits render above/below by MathML
      // convention; matches Word's default for ∑/∏. For ∫ this gives
      // above-below too, which is the LaTeX `\displaystyle` shape.
      limitsAboveBelow = true
      if (isMmlElement(head, "munder")) sub = [kids[1]!]
      else if (isMmlElement(head, "mover")) sup = [kids[1]!]
      else {
        sub = [kids[1]!]
        sup = [kids[2]!]
      }
    }
  } else if (
    isMmlElement(head, "msub") ||
    isMmlElement(head, "msup") ||
    isMmlElement(head, "msubsup")
  ) {
    const kids = elementChildren(head)
    const base = kids[0]
    if (base && isMmlElement(base, "mo")) {
      chr = NARY_OPERATORS.get(mmlText(base))
      // ms{ub,up,ubsup} → limits in subscript/superscript position;
      // matches Word's default for ∫. For ∑ used as msubsup (rare in
      // temml output) this gives the inline-script shape.
      limitsAboveBelow = false
      if (isMmlElement(head, "msub")) sub = [kids[1]!]
      else if (isMmlElement(head, "msup")) sup = [kids[1]!]
      else {
        sub = [kids[1]!]
        sup = [kids[2]!]
      }
    }
  }

  if (chr === undefined) return null

  // Walk forward to collect operand siblings, stopping at an infix
  // binary operator (terminator), at another n-ary head, or end of
  // list.
  const operand: Element[] = []
  let j = i + 1
  while (j < siblings.length) {
    const next = siblings[j]!
    if (isMmlElement(next, "mo") && NARY_OPERAND_TERMINATORS.has(mmlText(next))) break
    if (detectNary(siblings, j) !== null) break
    operand.push(next)
    j++
  }

  return {
    chr,
    sub,
    sup,
    operand,
    consumed: j - i,
    limitsAboveBelow,
  }
}
