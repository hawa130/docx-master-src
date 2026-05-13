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
 * Two name tables:
 *
 *   - `nameIndex` — names bound to a paragraph element (origin `source` /
 *     `allocated` / `adopted`). Resolves via `resolveByName`. Source
 *     bookmarks whose `<w:bookmarkStart>` doesn't live directly inside a
 *     `<w:p>` are intentionally excluded — no paragraph-level target to
 *     surface for REF.
 *   - `reservations` — pre-scan placeholders for names a forthcoming
 *     ParagraphBlock.anchor will adopt. Carry a numbering hint so a
 *     forward InlineRef can answer "target is auto-numbered?" before the
 *     target element exists. `adoptName` consumes the reservation and
 *     binds the name into `nameIndex`.
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

/** A name bound to a paragraph element. `origin` is informational only —
 * affects nothing at the API surface; `commit` decides what to wrap based
 * on the `byElement` map. */
export interface BoundRecord {
  element: Element
  origin: "source" | "allocated" | "adopted"
}

/** Pre-scan reservation carrying the numbering hint declared on the
 * forthcoming Block (`styleId` from the Block, `directlyNumbered` true if
 * the Block declares `numbering: { numId, level }`). Lets a forward
 * InlineRef classify the target before its element exists. */
export interface Reservation {
  styleId?: string
  directlyNumbered?: boolean
}

export class BookmarkAllocator {
  private nextId: number
  /** Tracks every name in use across all three sources: source bookmarks
   * (including non-paragraph-level ones not surfaced in `nameIndex`),
   * bound records, and pending reservations. Used for collision detection
   * in `reserveName` / `adoptName` and to seed `allocName`'s retry loop. */
  private usedNames: Set<string>
  private byElement = new Map<Element, BookmarkAssignment>()
  private nameIndex = new Map<string, BoundRecord>()
  private reservations = new Map<string, Reservation>()

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
   * still recognize the name. Optional `ctx` captures the Block's
   * `styleId` and whether it declares `numbering` directly, so a forward
   * ref can answer "target is auto-numbered?" via `targetIsAutoNumbered`
   * before the target element exists.
   *
   * Throws on collision with a source bookmark, an existing reservation,
   * or an already-adopted name — the engine runs this once per declared
   * anchor before emit starts, so duplicates surface before any mutation. */
  reserveName(name: string, ctx?: Reservation): void {
    if (this.usedNames.has(name)) {
      throw new Error(
        `anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`,
      )
    }
    this.usedNames.add(name)
    this.reservations.set(name, {
      styleId: ctx?.styleId,
      directlyNumbered: ctx?.directlyNumbered,
    })
  }

  /** True iff `name` has been pre-scan reserved but not yet bound to an
   * element via `adoptName`. Lets the emitter branch into forward-ref
   * handling (style-cascade numbering check, late-bound backfill). */
  isReserved(name: string): boolean {
    return this.reservations.has(name)
  }

  /** Numbering hint captured at reservation time. Returns undefined when
   * the name isn't reserved. */
  predictedNumberingFor(name: string): Reservation | undefined {
    return this.reservations.get(name)
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
   * If the name was pre-scan reserved, the reservation is consumed and
   * the name is bound to `pEl` — not a duplicate.
   *
   * Idempotent for the same `(name, pEl)` pair. */
  adoptName(name: string, pEl: Element): BookmarkAssignment {
    const existing = this.byElement.get(pEl)
    if (existing) {
      if (existing.name === name) return existing
      throw new Error(
        `paragraph already declared anchor "${existing.name}" — cannot also declare anchor "${name}". Each paragraph carries at most one anchor.`,
      )
    }
    // Bound elsewhere (source bookmark / earlier adoption) blocks; a
    // pending reservation is the expected upgrade path.
    if (this.nameIndex.has(name)) {
      throw new Error(
        `anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`,
      )
    }
    this.reservations.delete(name)
    const id = this.nextId++
    this.usedNames.add(name)
    const assignment: BookmarkAssignment = { id, name }
    this.byElement.set(pEl, assignment)
    this.nameIndex.set(name, { element: pEl, origin: "adopted" })
    return assignment
  }

  /** Resolve a bookmark name to its paragraph element. Returns undefined
   * when the name isn't bound (either unknown or only pre-scan reserved —
   * caller branches via `isReserved` for the latter). */
  resolveByName(name: string): BoundRecord | undefined {
    return this.nameIndex.get(name)
  }

  /** Reserve a name + id for an inline (caption-emitter) bookmark
   * without binding to any element yet. Caption-emit needs the id+name
   * upfront to emit the bookmarkStart/End XML, but the paragraph
   * doesn't exist yet — it's about to be built around the bookmark.
   *
   * Same collision semantics as adoptName: source / prior-adopted name
   * → throw; reserved name → consume the reservation. The returned
   * assignment is NOT yet in nameIndex; call `bindRangeBookmark` after
   * the caption paragraph is constructed to record the binding so
   * `resolveByName` works for REF cross-references. */
  allocateRangeBookmark(name: string): BookmarkAssignment {
    // Reservations are the expected upgrade path (forward-ref pre-scan
    // reserves the name before caption emit allocates it). All other
    // usedNames hits are collisions: bound paragraphs in `nameIndex`,
    // or body-level source bookmarks in `usedNames` but not
    // `nameIndex` (spans multiple paragraphs, no surface for REF
    // resolution but still occupies the name).
    if (this.nameIndex.has(name)) {
      throw new Error(
        `anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`,
      )
    }
    if (this.usedNames.has(name) && !this.reservations.has(name)) {
      throw new Error(
        `anchor "${name}" ${this.describeCollision(name)}. Pick a unique anchor name.`,
      )
    }
    this.reservations.delete(name)
    const id = this.nextId++
    this.usedNames.add(name)
    return { id, name }
  }

  /** Post-allocation binding for `allocateRangeBookmark`. Records the
   * name → paragraph mapping in `nameIndex` so REF backfill can resolve
   * the target. Does NOT touch `byElement` — caption-emit already
   * emitted bookmarkStart/End inline, so commit() must NOT wrap the
   * paragraph again. */
  bindRangeBookmark(name: string, pEl: Element): void {
    this.nameIndex.set(name, { element: pEl, origin: "adopted" })
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

  /** Describe why `name` collides for the error message. Caller guarantees
   * the name is in `usedNames`; this just routes to the right phrase. */
  private describeCollision(name: string): string {
    const bound = this.nameIndex.get(name)
    if (bound) {
      return bound.origin === "source"
        ? "exists in the source document"
        : "was already adopted in this apply"
    }
    if (this.reservations.has(name)) {
      return "is already reserved by an earlier ParagraphBlock.anchor in this apply"
    }
    // Non-paragraph-level source bookmark — in usedNames but not nameIndex.
    return "exists in the source document"
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
    throw new Error(
      "BookmarkAllocator.allocName: could not find a unique bookmark name after 1000 retries. " +
        "This is an engine invariant violation; the 10^8 search space should make collision statistically impossible.",
    )
  }
}
