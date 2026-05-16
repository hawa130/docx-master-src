/**
 * Pair detection for `\big`-class fixed-size delimiters.
 *
 * temml emits `\big\| x \big\|` as
 *
 *   <mrow>
 *     <mo minsize="1.2em" maxsize="1.2em">‖</mo>
 *     ...body...
 *     <mo minsize="1.2em" maxsize="1.2em">‖</mo>
 *   </mrow>
 *
 * and `\big\| x \big\|_2^2` as
 *
 *   <mrow>
 *     <mo minsize="1.2em" maxsize="1.2em">‖</mo>
 *     ...body...
 *     <msubsup>
 *       <mrow><mo minsize="1.2em" maxsize="1.2em">‖</mo></mrow>
 *       <mn>2</mn>
 *       <mn>2</mn>
 *     </msubsup>
 *   </mrow>
 *
 * Both `fence="false"` (different from `\left…\right` which is `fence="true"`),
 * so the regular fence detector doesn't catch them. Without fusion they
 * render as plain runs that ignore Word's `w:sz` hint — the delimiters
 * end up visibly smaller than auto-grown `<m:d>` around tall content.
 *
 * Fix: detect the pair, wrap the body in `<m:d>` with `<m:grow/>`. If
 * the closing `<mo>` was inside an `m{sub,sup,subsup}`, lift the
 * script wrappers around the `<m:d>` as `<m:sSub>` / `<m:sSup>` /
 * `<m:sSubSup>`.
 */

import { isMmlElement, mmlText, elementChildren, attr } from "./dom.ts"

export interface SizedFenceMatch {
  begChr: string
  endChr: string
  body: Element[]
  /** Script wrapper around the close delimiter, if any. The sub/sup
   *  children belong to the wrapper, not the delimiter. */
  script: "none" | "sub" | "sup" | "subsup"
  scriptSub: Element | null
  scriptSup: Element | null
  consumed: number
}

/** Returns true if the element is a sized `<mo>` (`\big`, `\Big`, `\bigg`,
 *  `\Bigg`) — `<mo>` with `minsize` attribute matching `Nem` form. */
function isSizedMo(el: Element | undefined): boolean {
  if (!el || !isMmlElement(el, "mo")) return false
  const minsize = attr(el, "minsize")
  return minsize !== undefined && /^[\d.]+\s*em$/.test(minsize)
}

/** Returns the sized `<mo>` inside a script-element's base, or null.
 *  Looks through a single layer of `<mrow>` wrapping. */
function scriptBaseSizedMo(scriptEl: Element): Element | null {
  const kids = elementChildren(scriptEl)
  const base = kids[0]
  if (!base) return null
  if (isSizedMo(base)) return base
  // temml's actual shape: msubsup > mrow > mo
  if (isMmlElement(base, "mrow")) {
    const baseKids = elementChildren(base)
    if (baseKids.length >= 1 && isSizedMo(baseKids[0])) return baseKids[0]!
  }
  return null
}

export function detectSizedFence(siblings: Element[], i: number): SizedFenceMatch | null {
  const open = siblings[i]
  if (!isSizedMo(open)) return null
  const begChr = mmlText(open!)

  // Scan forward for the closing mo — either a direct sibling sized mo,
  // or a script-element whose base contains a sized mo.
  for (let j = i + 1; j < siblings.length; j++) {
    const next = siblings[j]!
    if (isSizedMo(next)) {
      return {
        begChr,
        endChr: mmlText(next),
        body: siblings.slice(i + 1, j),
        script: "none",
        scriptSub: null,
        scriptSup: null,
        consumed: j - i + 1,
      }
    }
    const scriptKind = scriptElementKind(next)
    if (scriptKind !== null) {
      const closeMo = scriptBaseSizedMo(next)
      if (closeMo !== null) {
        const scriptKids = elementChildren(next)
        return {
          begChr,
          endChr: mmlText(closeMo),
          body: siblings.slice(i + 1, j),
          script: scriptKind,
          scriptSub: scriptKind === "sub" || scriptKind === "subsup" ? scriptKids[1]! : null,
          scriptSup:
            scriptKind === "sup" ? scriptKids[1]! : scriptKind === "subsup" ? scriptKids[2]! : null,
          consumed: j - i + 1,
        }
      }
    }
  }
  return null
}

function scriptElementKind(el: Element): "sub" | "sup" | "subsup" | null {
  if (isMmlElement(el, "msub")) return "sub"
  if (isMmlElement(el, "msup")) return "sup"
  if (isMmlElement(el, "msubsup")) return "subsup"
  return null
}
