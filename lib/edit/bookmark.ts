/**
 * Bookmark allocator for cross-reference targets.
 *
 * OOXML bookmarks (`<w:bookmarkStart w:id="N" w:name="..."/>` and matching
 * `<w:bookmarkEnd w:id="N"/>`) pair a numeric id with a string name. The id
 * is unique within the document; the name is the handle REF fields use to
 * resolve targets at render time. Word auto-emits these with names like
 * `_Ref` + 8-digit decimal whenever the user inserts a cross-reference.
 *
 * This allocator scans the source `document.xml` for max existing `w:id`
 * and used names on construction, then hands out fresh `(id, name)` pairs
 * via `getOrAllocate(paragraphElement)`. Re-calling with the same Element
 * returns the existing assignment — so an InlineRef emitted twice against
 * one paragraph wraps that paragraph exactly once.
 *
 * Lazy by design: only paragraphs actually referenced by an `InlineRef`
 * get wrapped. Walking the doc to eager-wrap every numbered paragraph is
 * cheap but adds bookmark elements the user never asked for; lazy keeps
 * the OOXML diff scoped to what the edits actually use.
 *
 * Call `commit(documentDoc)` after all edits finish to perform the actual
 * `<w:bookmarkStart>` / `<w:bookmarkEnd>` wrapping — deferring the DOM
 * mutation until edits complete avoids re-resolving Element refs that the
 * mid-edit insertion logic relies on.
 */

import { NS } from "@lib/parse/types.ts"
import { getChildren, wAttr } from "@lib/xml/xml-utils.ts"

export interface BookmarkAssignment {
  id: number
  name: string
}

export class BookmarkAllocator {
  private nextId: number
  private usedNames: Set<string>
  private byElement = new Map<Element, BookmarkAssignment>()

  constructor(documentDoc: Document) {
    this.nextId = 0
    this.usedNames = new Set()
    const root = documentDoc.documentElement
    if (root) {
      const starts = root.getElementsByTagNameNS(NS.w, "bookmarkStart")
      for (let i = 0; i < starts.length; i++) {
        const el = starts[i]!
        const idAttr = wAttr(el, "id")
        if (idAttr) {
          const n = parseInt(idAttr, 10)
          if (Number.isFinite(n) && n >= this.nextId) this.nextId = n + 1
        }
        const name = wAttr(el, "name")
        if (name) this.usedNames.add(name)
      }
    }
  }

  /** Return the assignment for this paragraph element. Allocates lazily on
   * first call per element. Stable within an allocator's lifetime so
   * multiple refs to the same paragraph share one bookmark. */
  getOrAllocate(pEl: Element): BookmarkAssignment {
    const cached = this.byElement.get(pEl)
    if (cached) return cached
    const id = this.nextId++
    const name = this.allocName()
    const assignment: BookmarkAssignment = { id, name }
    this.byElement.set(pEl, assignment)
    return assignment
  }

  /** True iff at least one bookmark was allocated. Caller uses this to
   * decide whether to touch settings.xml's updateFields flag. */
  hasAllocations(): boolean {
    return this.byElement.size > 0
  }

  /** Wrap each allocated paragraph with `<w:bookmarkStart>` / `<w:bookmarkEnd>`.
   * Start goes immediately after `<w:pPr>` (or at the front when no pPr);
   * end goes at the paragraph's tail. Idempotent: subsequent calls do nothing
   * because the internal queue is consumed. */
  commit(documentDoc: Document): void {
    const w = NS.w
    for (const [pEl, { id, name }] of this.byElement) {
      const start = documentDoc.createElementNS(w, "w:bookmarkStart")
      start.setAttributeNS(w, "w:id", String(id))
      start.setAttributeNS(w, "w:name", name)
      const end = documentDoc.createElementNS(w, "w:bookmarkEnd")
      end.setAttributeNS(w, "w:id", String(id))

      // Insert bookmarkStart right after <w:pPr> if present, otherwise at
      // the head. bookmarkEnd appended to the paragraph's tail. Per
      // ECMA-376 the bookmark element pair is permitted at <w:p>'s child
      // level and wrapping the entire paragraph content is what Word
      // produces when you insert a cross-reference target.
      const children = getChildren(pEl)
      const pPr = children.find((c) => c.namespaceURI === w && c.localName === "pPr")
      if (pPr && pPr.nextSibling) {
        pEl.insertBefore(start, pPr.nextSibling)
      } else if (pPr) {
        pEl.appendChild(start)
      } else if (pEl.firstChild) {
        pEl.insertBefore(start, pEl.firstChild)
      } else {
        pEl.appendChild(start)
      }
      pEl.appendChild(end)
    }
    this.byElement.clear()
  }

  /** Word convention: `_Ref` + 8 decimal digits, allocator-local counter.
   * Re-rolls on collision with a pre-existing name. The numeric search
   * space is 10^8, so collision is statistically impossible in normal
   * use; the retry is defensive only. */
  private allocName(): string {
    let attempt = 0
    while (attempt < 1000) {
      const n = Math.floor(Math.random() * 100_000_000)
      const name = `_Ref${n.toString().padStart(8, "0")}`
      if (!this.usedNames.has(name)) {
        this.usedNames.add(name)
        return name
      }
      attempt++
    }
    throw new Error("BookmarkAllocator: could not find a unique bookmark name after 1000 retries")
  }
}
