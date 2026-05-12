/**
 * Bookmark allocator for cross-reference targets.
 *
 * OOXML bookmarks (`<w:bookmarkStart w:id="N" w:name="..."/>` and matching
 * `<w:bookmarkEnd w:id="N"/>`) pair a numeric id with a string name. The id
 * is unique within the document; the name is the handle REF fields use to
 * resolve targets at render time. Word auto-emits these with names like
 * `_Ref` + 8-digit decimal whenever the user inserts a cross-reference.
 *
 * Two registration paths:
 *
 *   1. `getOrAllocate(pEl)` — anonymous bookmark, auto-named `_Ref<8hex>`.
 *      Used for InlineRefs that target a paragraph by index; the agent
 *      doesn't care what the bookmark is called.
 *
 *   2. `adoptName(name, pEl)` — caller-supplied stable name (validated
 *      against Word's bookmark name rules upstream). Used by
 *      ParagraphBlock.anchor so later refs can address the paragraph by
 *      name across edits[] (or across apply runs).
 *
 * Element-keyed: re-registering the same element returns the existing
 * assignment, so an InlineRef + an anchor on the same paragraph share
 * one bookmark.
 *
 * Name lookup via `resolveByName(name)` covers two sources:
 *   - adopted anchors (current apply run)
 *   - source bookmarks (pre-existing in document.xml), captured at
 *     construction; only those whose `<w:bookmarkStart>` lives directly
 *     inside a `<w:p>` are exposed — bookmarks spanning multiple
 *     paragraphs or sitting at body level have no paragraph-level target
 *     we can render.
 *
 * Lazy by design: only paragraphs actually referenced (or explicitly
 * named via anchor) get wrapped at `commit` time.
 */

import { NS } from "@lib/parse/types.ts"
import { getChildren, wAttr } from "@lib/xml/xml-utils.ts"

export interface BookmarkAssignment {
  id: number
  name: string
}

interface NameRecord {
  element: Element
  /** "source" means the bookmark already exists in document.xml — commit
   * skips it. "allocated" / "adopted" means we'll wrap on commit. */
  origin: "source" | "allocated" | "adopted"
}

export class BookmarkAllocator {
  private nextId: number
  private usedNames: Set<string>
  private byElement = new Map<Element, BookmarkAssignment>()
  /** Reverse index: name → record. Covers source bookmarks (read-only),
   * our adopted anchors, and auto-allocated names. */
  private nameIndex = new Map<string, NameRecord>()

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
        if (!name) continue
        this.usedNames.add(name)
        // Map source bookmarks → enclosing paragraph (when applicable).
        // Bookmarks whose parent isn't a <w:p> (e.g. spanning multiple
        // paragraphs at body level) are intentionally NOT indexed — we
        // have no paragraph-level target to surface for REF rendering.
        const parent = el.parentNode
        if (
          parent &&
          (parent as Element).namespaceURI === NS.w &&
          (parent as Element).localName === "p"
        ) {
          this.nameIndex.set(name, { element: parent as Element, origin: "source" })
        }
      }
    }
  }

  /** Return the assignment for this paragraph element. Allocates lazily on
   * first call per element with an auto-generated `_Ref<8hex>` name.
   * Stable within an allocator's lifetime so multiple refs to the same
   * paragraph share one bookmark. */
  getOrAllocate(pEl: Element): BookmarkAssignment {
    const cached = this.byElement.get(pEl)
    if (cached) return cached
    const id = this.nextId++
    const name = this.allocName()
    const assignment: BookmarkAssignment = { id, name }
    this.byElement.set(pEl, assignment)
    this.nameIndex.set(name, { element: pEl, origin: "allocated" })
    return assignment
  }

  /** Adopt a caller-supplied name for this paragraph (the `anchor` field
   * on a ParagraphBlock). Two failure modes:
   *
   *   - `name` collides with an existing source bookmark or with a name
   *     already adopted in this run → throws (the agent picked a name
   *     that's not unique).
   *   - The same `pEl` is being re-bound to a different name → throws (a
   *     paragraph carries at most one anchor; if you need a second handle,
   *     refTo it by paragraph index instead).
   *
   * Idempotent for the same `(name, pEl)` pair. */
  adoptName(name: string, pEl: Element): BookmarkAssignment {
    const existing = this.byElement.get(pEl)
    if (existing) {
      if (existing.name === name) return existing
      throw new Error(
        `BookmarkAllocator.adoptName: paragraph already bound to bookmark "${existing.name}" — cannot also adopt "${name}". Each paragraph carries at most one named anchor.`,
      )
    }
    if (this.usedNames.has(name)) {
      const existingRec = this.nameIndex.get(name)
      const where =
        existingRec?.origin === "source"
          ? "exists in the source document"
          : "was already adopted in this apply"
      throw new Error(
        `BookmarkAllocator.adoptName: bookmark name "${name}" ${where}. Pick a different anchor name.`,
      )
    }
    const id = this.nextId++
    this.usedNames.add(name)
    const assignment: BookmarkAssignment = { id, name }
    this.byElement.set(pEl, assignment)
    this.nameIndex.set(name, { element: pEl, origin: "adopted" })
    return assignment
  }

  /** Resolve a bookmark name to its paragraph element. Searches adopted
   * anchors (this run's ParagraphBlock.anchor) and source bookmarks. Returns
   * undefined when the name isn't found in either; caller throws with
   * agent-readable context. */
  resolveByName(name: string): NameRecord | undefined {
    return this.nameIndex.get(name)
  }

  /** True iff at least one bookmark will be wrapped at commit. Source
   * bookmarks don't count — they're already in the XML. */
  hasAllocations(): boolean {
    return this.byElement.size > 0
  }

  /** Wrap each allocated / adopted paragraph with `<w:bookmarkStart>` /
   * `<w:bookmarkEnd>`. Start goes immediately after `<w:pPr>` (or at the
   * front when no pPr); end goes at the paragraph's tail. Idempotent:
   * subsequent calls do nothing because the internal queue is consumed.
   * Source bookmarks are skipped — they already exist in document.xml. */
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
