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
import type { Block } from "@lib/config/edit-types.ts"
import type { EmitContext } from "@lib/edit/fragment-emit.ts"
import { buildPPrChildren } from "@lib/edit/fragment-emit.ts"
import { buildOMath } from "@lib/edit/math/omml-build.ts"
import { emitNumberedEquation } from "@lib/edit/caption-emit.ts"

const w = NS.w
const m = NS.m

export function emitEquationBlock(
  block: Extract<Block, { type: "equation" }>,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  // When captionId is set, route to the caption pipeline — emits a 3-col
  // borderless table with the equation in the middle cell and the
  // SEQ-based caption in the right cell.
  if (block.captionId !== undefined) {
    return emitNumberedEquationDispatch(block, block.captionId, ownerDoc, ctx)
  }
  return emitUnnumberedEquationLegacy(block, ownerDoc, ctx)
}

function emitNumberedEquationDispatch(
  block: Extract<Block, { type: "equation" }>,
  captionId: string,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  if (!ctx.captions) {
    throw new Error(
      "EquationBlock.captionId: ctx.captions callbacks not provided by the engine. " +
        "Numbered equations require the captions table to be declared in the apply config.",
    )
  }
  const config = ctx.captions.resolve(captionId)
  if (!config) {
    throw new Error(`EquationBlock: captionId "${captionId}" is not declared in captions table.`)
  }
  if (block.subGroup !== undefined && config.subCounter === undefined) {
    throw new Error(
      `EquationBlock: subGroup="${block.subGroup}" requires captions["${captionId}"].subCounter to be declared.`,
    )
  }
  const mathSource: { latex: string } | { omml: string } =
    block.latex !== undefined ? { latex: block.latex } : { omml: block.omml! }

  const bookmark =
    block.anchor !== undefined ? ctx.captions.allocateBookmark(block.anchor) : undefined
  const { table, fill } = emitNumberedEquation(ownerDoc, {
    mathSource,
    equationStyleId: block.styleId ?? "Equation",
    captionConfig: config,
    subGroup: block.subGroup,
    bookmark,
    usableWidthTwips: ctx.usableWidthTwips,
  })
  if (block.anchor !== undefined) {
    ctx.captions.bindBookmark(block.anchor, fill.paragraph)
  }
  ctx.captions.registerFill(fill)
  return table
}

function emitUnnumberedEquationLegacy(
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
  // Schema requires exactly one of `latex` / `omml` on every equation block;
  // the throw is an engine invariant guard (validation should have rejected
  // a malformed block before reaching emit).
  const mathSource =
    block.latex !== undefined
      ? { latex: block.latex }
      : block.omml !== undefined
        ? { omml: block.omml }
        : (() => {
            throw new Error(
              "EquationBlock: no math source — schema validation should have caught this. " +
                "This is an engine invariant violation.",
            )
          })()
  const oMath = buildOMath(mathSource, ownerDoc, true)
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
  return buildOMath({ latex }, ownerDoc, false)
}
