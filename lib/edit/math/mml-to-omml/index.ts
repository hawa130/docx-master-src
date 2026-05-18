/**
 * MathML → OMML converter — in-tree replacement for the LGPL
 * `mathml2omml` package. Public surface is one function.
 *
 *   import { convertMathMLToOMML } from "@lib/edit/math/mml-to-omml/index.ts"
 *
 * The output is a string suitable for embedding as `<m:oMath>` inside
 * a Word `<w:p>` (display) or inline run sequence. The string carries
 * `xmlns:m` on its root, so it survives importNode into any document
 * without further namespace bookkeeping.
 *
 * Module layout (all under this directory):
 *   constants.ts — namespaces, n-ary chr table, accent + bar chars,
 *                  mathvariant → m:sty mapping
 *   style.ts     — leaf-kind → m:sty resolution (italic-by-length
 *                  default for mi)
 *   run.ts       — <m:r> with <w:rPr> + optional <m:rPr> + <m:t>
 *   dom.ts       — DOM walking helpers + element factories
 *   nary.ts      — n-ary fusion detection (this is the structural fix
 *                  vs mml2omml's empty-<m:e/> bug)
 *   walk.ts      — element dispatch table; recurses through emitChildren
 *
 * Element coverage:
 *   leaves      mi mn mo mtext ms mspace
 *   structural  mrow msub msup msubsup munder mover munderover
 *               mfrac msqrt mroot mfenced mtable mtr mtd
 *               mphantom (→ m:phant)
 *   transparent mpadded mstyle menclose maction (→ children only)
 *   discarded   semantics, annotation, annotation-xml
 *   accents     mover with accent="true" or accent-class char → m:acc
 *               mover/munder with bar-class char → m:bar
 *
 * Any other element triggers a throw with a pointer to the omml escape
 * hatch in references/equations.md. The error contract matches the rest
 * of the math pipeline (latex-to-omml.ts).
 */

import { parseXml } from "@lib/xml/reader.ts"
import { XMLSerializer } from "@xmldom/xmldom"
import { MML_NS, M_NS } from "./constants.ts"
import { emitChildren } from "./walk.ts"

const serializer = new XMLSerializer()

export function convertMathMLToOMML(mathmlXml: string): string {
  const parsed = parseXml(mathmlXml)
  const root = parsed.documentElement
  if (!root) {
    throw new Error(`MathML → OMML: input has no root element`)
  }
  if (root.namespaceURI !== MML_NS || root.localName !== "math") {
    throw new Error(
      `MathML → OMML: expected <math xmlns="${MML_NS}"> root, got ` +
        `<${root.localName} xmlns="${root.namespaceURI ?? ""}">`,
    )
  }
  // Build the OMML output document by parsing a stub <m:oMath> with the
  // m: prefix declared — avoids using DOMImplementation.createDocument
  // (whose @xmldom/xmldom types don't line up with the lib.dom Document
  // we use everywhere else).
  const outDoc = parseXml(`<m:oMath xmlns:m="${M_NS}"/>`)
  const oMath = outDoc.documentElement!
  emitChildren(root, oMath, outDoc)
  return serializer.serializeToString(
    outDoc as unknown as Parameters<XMLSerializer["serializeToString"]>[0],
  )
}
