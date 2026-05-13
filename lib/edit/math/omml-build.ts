/**
 * Shared `<m:oMath>` builder.
 *
 * Resolves a math source (LaTeX via temml, or raw OMML escape hatch)
 * into an `<m:oMath>` Element imported into `ownerDoc`. The OMML
 * produced by `latexToOmml` already declares `xmlns:m` (and the
 * unused-but-harmless `xmlns:w`) on its root — xmldom preserves these
 * on import, so the result serializes correctly without touching the
 * document root's namespace declarations.
 *
 * `displayMode` is forwarded to temml; OMML escape-hatch sources ignore
 * it (the agent-supplied XML already encodes display vs inline).
 *
 * Both caption-emit (numbered equation cell content) and equation-emit
 * (numbered + unnumbered top-level equations) share this resolver —
 * splitting it out avoids a circular import (equation-emit already
 * imports `emitNumberedEquation` from caption-emit).
 */

import { parseXml } from "@lib/xml/reader.ts"
import { getOmmlSync } from "@lib/edit/math/latex-to-omml.ts"

export type MathSource = { latex: string } | { omml: string }

export function buildOMath(source: MathSource, ownerDoc: Document, displayMode: boolean): Element {
  const xml = "latex" in source ? getOmmlSync(source.latex, displayMode) : source.omml
  const parsed = parseXml(xml)
  const root = parsed.documentElement
  if (!root) {
    throw new Error(`Math source produced no root element: ${JSON.stringify(source)}`)
  }
  return ownerDoc.importNode(root, true) as Element
}
