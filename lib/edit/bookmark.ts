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

export interface NameRecord {
  /** Element is null for `"reserved"` records: the pre-scan reserved the
   * name before the corresponding ParagraphBlock emitted its <w:p>. The
   * element is filled in when `adoptName` upgrades the record to
   * `"adopted"`. */
  element: Element | null
  /** "source" means the bookmark already exists in document.xml — commit
   * skips it. "allocated" / "adopted" means we'll wrap on commit.
   * "reserved" is a pre-scan placeholder: the engine has seen the name
   * declared on a forthcoming ParagraphBlock.anchor but hasn't emitted
   * the paragraph yet. Reserved records don't carry an element and are
   * not surfaced via `resolveByName` until upgraded by `adoptName`. */
  origin: "source" | "allocated" | "adopted" | "reserved"
  /** Numbering hint for `"reserved"` records — captured from the
   * forthcoming Block's `styleId` and direct `numbering` field. Lets a
   * forward InlineRef answer "target is auto-numbered?" before the
   * target element exists. Unset for non-reserved records (the element
   * carries the answer directly). */
  predictedNumbering?: { styleId?: string; directlyNumbered?: boolean }
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

  /** Pre-scan reservation: register a name that an upcoming Block will
   * adopt, so a forward InlineRef emitted earlier in the pipeline can
   * still recognize the name. The element isn't known yet — `adoptName`
   * fills it in once the block emits. Optional `ctx` captures the
   * Block's `styleId` and whether it declares `numbering` directly, so
   * a forward ref can answer "target is auto-numbered?" via
   * `targetIsAutoNumbered` before the target element exists.
   *
   * Throws on collision with a source bookmark, an existing reservation,
   * or an already-adopted name — the engine runs this once per declared
   * anchor before emit starts, so duplicates surface before any mutation. */
  reserveName(name: string, ctx?: { styleId?: string; directlyNumbered?: boolean }): void {
    if (this.usedNames.has(name)) {
      const existingRec = this.nameIndex.get(name)
      const where =
        existingRec?.origin === "source"
          ? "exists in the source document"
          : existingRec?.origin === "reserved"
            ? "is already reserved by an earlier ParagraphBlock.anchor in this apply"
            : "was already adopted in this apply"
      throw new Error(
        `BookmarkAllocator.reserveName: bookmark name "${name}" ${where}. Pick a different anchor name.`,
      )
    }
    this.usedNames.add(name)
    this.nameIndex.set(name, {
      element: null,
      origin: "reserved",
      predictedNumbering:
        ctx?.styleId !== undefined || ctx?.directlyNumbered !== undefined
          ? { styleId: ctx.styleId, directlyNumbered: ctx.directlyNumbered }
          : undefined,
    })
  }

  /** True iff `name` has been pre-scan reserved but not yet bound to an
   * element via `adoptName`. Lets the emitter branch into forward-ref
   * handling (style-cascade numbering check, late-bound backfill). */
  isReserved(name: string): boolean {
    return this.nameIndex.get(name)?.origin === "reserved"
  }

  /** Numbering hint captured at reservation time. Returns undefined when
   * the name isn't reserved or no hint was supplied. */
  predictedNumberingFor(
    name: string,
  ): { styleId?: string; directlyNumbered?: boolean } | undefined {
    const rec = this.nameIndex.get(name)
    return rec?.origin === "reserved" ? rec.predictedNumbering : undefined
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
   * If the name was already `"reserved"` by the pre-scan, the record is
   * upgraded to `"adopted"` with `element` filled in — not a duplicate.
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
    const existingRec = this.nameIndex.get(name)
    if (existingRec && existingRec.origin !== "reserved") {
      const where =
        existingRec.origin === "source"
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
    // Upgrade the reserved record (or insert a fresh adopted one).
    this.nameIndex.set(name, { element: pEl, origin: "adopted" })
    return assignment
  }

  /** Resolve a bookmark name to its paragraph element. Searches adopted
   * anchors (this run's ParagraphBlock.anchor), source bookmarks, and
   * auto-allocated names. Returns undefined when the name isn't found at
   * all, OR when the only matching record is a pre-scan `"reserved"`
   * placeholder (no element bound yet — caller branches via `isReserved`).
   * Caller throws with agent-readable context. */
  resolveByName(name: string): NameRecord | undefined {
    const rec = this.nameIndex.get(name)
    if (!rec) return undefined
    if (rec.element === null) return undefined
    return rec
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
