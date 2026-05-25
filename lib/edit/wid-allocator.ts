/**
 * Shared `w:id` allocator for the document-wide ID space.
 *
 * OOXML elements that carry a `w:id` attribute share a single document-scoped
 * ID space: `<w:bookmarkStart>` / `<w:bookmarkEnd>`, `<w:ins>` / `<w:del>` /
 * `<w:moveFrom>` / `<w:moveTo>` and their range markers, `<w:commentRangeStart>` /
 * end / reference, `<w:rPrChange>` / `<w:pPrChange>` / `<w:tblPrChange>` /
 * `<w:tcPrChange>` / `<w:trPrChange>` / `<w:sectPrChange>` / `<w:numPrChange>`.
 *
 * Allocating per-subsystem from 0 collides both with source IDs already present
 * in the document and across subsystems (bookmarks vs revisions allocating the
 * same number). Route every allocation through one instance per apply so the
 * counter is monotonic and starts above every pre-existing ID.
 *
 * Scan strategy: walk every element and read the `w:id` attribute (the
 * attribute IS in the `w` namespace — `getAttributeNS(w, "id")`). Scanning
 * by attribute rather than enumerating tag names survives future OOXML
 * extensions that introduce new `w:id`-bearing elements without code changes.
 */

import { NS } from "@lib/parse/types.ts"

const w = NS.w

/**
 * Two-tier `w:id` lookup. `getAttributeNS` alone is unreliable: @xmldom/xmldom
 * does not consistently namespace-resolve prefixed attributes on parsed XML,
 * so `w:id="N"` may arrive with `namespaceURI === null` and qualified name
 * `"w:id"`. The codebase-wide `wAttr` helper adds a third tier — bare
 * `getAttribute("id")` — which is correct for most `w:*` attributes but
 * **wrong here**: DrawingML wraps inside `w:drawing` carry unrelated `id`
 * attributes (`<wp:docPr id="..."/>`, `<pic:cNvPr id="..."/>`) that live in
 * a separate ID space from OOXML's `w:id`. Picking those up pollutes the
 * counter with millions and breaks deterministic emission.
 */
function readWId(el: Element): string | null {
  const ns = el.getAttributeNS(w, "id")
  if (ns !== null && ns !== "") return ns
  const raw = el.getAttribute("w:id")
  return raw && raw !== "" ? raw : null
}

export class WIdAllocator {
  private nextId: number

  constructor(documentDoc: Document) {
    let max = -1
    const root = documentDoc.documentElement
    if (root) {
      const all = root.getElementsByTagName("*")
      for (let i = 0; i < all.length; i++) {
        const idAttr = readWId(all[i]!)
        if (!idAttr) continue
        const n = parseInt(idAttr, 10)
        if (Number.isFinite(n) && n > max) max = n
      }
    }
    this.nextId = max + 1
  }

  next(): number {
    return this.nextId++
  }
}
