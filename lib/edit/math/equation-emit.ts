/**
 * EquationBlock + InlineEquation → OOXML.
 *
 * Block form emits a `<w:p>` carrying `<m:oMathPara>` (centered display
 * math). Inline form returns the `<m:oMath>` element to splice into a run
 * sequence inside an existing paragraph.
 *
 * The OMML produced by `latexToOmml` already declares `xmlns:m` (and the
 * unused-but-harmless `xmlns:w`) on the `<m:oMath>` root — xmldom preserves
 * these on import, so the result serializes correctly without needing to
 * touch the document root's namespace declarations.
 */

import { NS } from "@lib/parse/types.ts"
import { parseXml } from "@lib/xml/reader.ts"
import type { Block } from "@lib/config/edit-types.ts"
import type { EmitContext } from "@lib/edit/fragment-emit.ts"
import { buildPPrChildren } from "@lib/edit/fragment-emit.ts"
import { getOmmlSync } from "@lib/edit/math/latex-to-omml.ts"

const w = NS.w
const m = NS.m

/** Import the cached `<m:oMath>` element into `ownerDoc`. The latex was
 * resolved up-front by the engine's prepareLatex pre-walk; this just parses
 * and adopts the string. */
function buildOMath(latex: string, displayMode: boolean, ownerDoc: Document): Element {
  const ommlString = getOmmlSync(latex, displayMode)
  const parsed = parseXml(ommlString)
  const root = parsed.documentElement
  if (!root)
    throw new Error(`LaTeX-to-OMML produced no root element. latex=${JSON.stringify(latex)}`)
  return ownerDoc.importNode(root, true) as Element
}

export function emitEquationBlock(
  block: Extract<Block, { type: "equation" }>,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  const p = ownerDoc.createElementNS(w, "w:p")
  if (block.styleId !== undefined || block.paraFormat !== undefined) {
    const pPr = ownerDoc.createElementNS(w, "w:pPr")
    if (block.styleId) {
      const ps = ownerDoc.createElementNS(w, "w:pStyle")
      ps.setAttributeNS(w, "w:val", block.styleId)
      pPr.appendChild(ps)
    }
    if (block.paraFormat) {
      for (const c of buildPPrChildren(block.paraFormat, ownerDoc)) pPr.appendChild(c)
    }
    p.appendChild(pPr)
  }
  const oMath = buildOMath(block.latex, true, ownerDoc)
  const oMathPara = ownerDoc.createElementNS(m, "m:oMathPara")
  oMathPara.appendChild(oMath)
  p.appendChild(oMathPara)
  if (block.anchor) {
    if (!ctx.adoptAnchor) {
      throw new Error(
        "EquationBlock.anchor encountered but ctx.adoptAnchor was not provided by the engine",
      )
    }
    ctx.adoptAnchor(block.anchor, p)
  }
  return p
}

/** Build the `<m:oMath>` element for an inline equation. Caller splices it
 * between `<w:r>` siblings inside the paragraph. */
export function emitInlineEquation(latex: string, ownerDoc: Document): Element {
  return buildOMath(latex, false, ownerDoc)
}
