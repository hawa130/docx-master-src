/**
 * Schema-correct insertion helpers for OOXML mutation paths.
 *
 * The engine builds elements with a per-call convenience order (e.g.
 * `pPrAdditions` pushes spacing → ind → jc → outlineLvl in upsertStyle, the
 * order they're constructed in code). CT_PPr's actual schema order is
 * different (spacing → ind → jc → outlineLvl per ECMA-376). Inserting
 * blindly with `appendChild` produces schema-violating XML that Word
 * rejects with "needs repair".
 *
 * `insertChildInOrder` places a new child at its schema-correct position
 * relative to existing children, given an explicit ordered list of localNames.
 *
 * Kept separate from docx-validate.ts so the engine can use it without
 * importing the validator.
 */

import { NS } from "@lib/parse/types.ts"
import { getChildren } from "@lib/xml/xml-utils.ts"

/** EG_RPrBase child order, per ECMA-376 17.7.9.1 (children of CT_RPr inside
 * style / run / paragraph mark). Truncated to the children this engine
 * actually emits. Notably: `b` / `bCs` come BEFORE `sz` / `szCs` — Word's
 * loader rejects mis-ordered run properties with "file needs repair" even
 * though XSD validators (xmllint) often pass them via relaxed xs:all. */
export const RPR_CHILD_ORDER = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
] as const

/** CT_PPr child order, per ECMA-376 17.3.1.26. Truncated to children we
 * encounter when building or mutating paragraph properties. */
export const PPR_CHILD_ORDER = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
] as const

/**
 * Insert `newChild` at its schema-correct position per `expectedOrder`.
 * Children whose localName is later in the order go after; earlier ones go
 * before. Falls back to append if `newChild`'s localName isn't in the list.
 */
export function insertChildInOrder(
  parent: Element,
  newChild: Element,
  expectedOrder: ReadonlyArray<string>,
  ns = NS.w,
): void {
  const newName = newChild.localName!
  const newIdx = expectedOrder.indexOf(newName)
  if (newIdx < 0) {
    parent.appendChild(newChild)
    return
  }
  for (const sibling of getChildren(parent)) {
    if (sibling.namespaceURI !== ns) continue
    const sibIdx = expectedOrder.indexOf(sibling.localName!)
    if (sibIdx > newIdx) {
      parent.insertBefore(newChild, sibling)
      return
    }
  }
  parent.appendChild(newChild)
}
