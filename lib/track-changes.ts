/**
 * Track-changes wrappers — produce <w:ins> / <w:del> / <w:rPrChange> /
 * <w:pPrChange> / <w:numPrChange> elements from edit ops.
 *
 * The engine wires `TrackContext.enabled`; when false, every helper is a
 * no-op (deletion still removes, insertion still inserts plain). When
 * true, deletions get wrapped, insertions get wrapped, format ops attach
 * a snapshot of the previous run / paragraph properties.
 *
 * Author and date defaults come from TrackContext (author empty, date
 * fixed for the run). The schema (ECMA-376 §17.13) requires both
 * attributes; we emit empty author since identity is not in scope.
 */

import { NS } from "@lib/types.ts"
import { firstChildNS, getChildren, getChildrenNS } from "@lib/xml-utils.ts"
import type { TrackContext } from "./edit-types.ts"

const w = NS.w

/* ------------- text-level wrappers ------------- */

/**
 * Wrap a list of paragraph children (typically <w:r> elements) inside a
 * single <w:ins>. The caller passes the already-built run nodes; this just
 * groups them. Returns the <w:ins> element ready to insert.
 */
export function wrapInsertion(runs: Element[], ownerDoc: Document, ctx: TrackContext): Element {
  const ins = ownerDoc.createElementNS(w, "w:ins")
  ins.setAttributeNS(w, "w:id", String(ctx.nextId()))
  ins.setAttributeNS(w, "w:author", ctx.author)
  ins.setAttributeNS(w, "w:date", ctx.date)
  for (const r of runs) ins.appendChild(r)
  return ins
}

/**
 * Wrap a list of <w:r> elements inside a <w:del> element, converting each
 * descendant <w:t> to <w:delText> per ECMA-376 §17.4.16. The runs are
 * detached from any current parent before being re-appended; callers are
 * responsible for splicing the returned element back into the tree.
 */
export function wrapDeletion(runs: Element[], ownerDoc: Document, ctx: TrackContext): Element {
  const del = ownerDoc.createElementNS(w, "w:del")
  del.setAttributeNS(w, "w:id", String(ctx.nextId()))
  del.setAttributeNS(w, "w:author", ctx.author)
  del.setAttributeNS(w, "w:date", ctx.date)
  for (const r of runs) {
    convertTtoDelText(r, ownerDoc)
    del.appendChild(r)
  }
  return del
}

function convertTtoDelText(node: Element, ownerDoc: Document): void {
  // Recurse: <w:t> → <w:delText>; <w:instrText> → <w:delInstrText>. Other
  // namespaces / elements left alone.
  for (const child of Array.from(getChildren(node))) {
    if (child.namespaceURI === w) {
      if (child.localName === "t") replaceLocalName(child, "w:delText", ownerDoc)
      else if (child.localName === "instrText") replaceLocalName(child, "w:delInstrText", ownerDoc)
      else convertTtoDelText(child, ownerDoc)
    } else {
      convertTtoDelText(child, ownerDoc)
    }
  }
}

function replaceLocalName(el: Element, qname: string, ownerDoc: Document): void {
  const rep = ownerDoc.createElementNS(w, qname)
  // Carry attributes over
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes[i]!
    if (attr.namespaceURI) rep.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
    else rep.setAttribute(attr.name, attr.value)
  }
  // Carry children over
  while (el.firstChild) rep.appendChild(el.firstChild)
  el.parentNode!.replaceChild(rep, el)
}

/**
 * Mark the paragraph mark (<w:p>'s pilcrow) as deleted, so when the user
 * accepts changes the paragraph merges with the next. Used when an entire
 * paragraph is deleted under track-changes; without this, accept leaves an
 * empty paragraph behind.
 *
 * Implementation: ensure <w:pPr><w:rPr><w:del/></w:rPr> exists.
 */
export function markParagraphMarkDeleted(p: Element, ownerDoc: Document, ctx: TrackContext): void {
  let pPr = firstChildNS(p, w, "pPr")
  if (!pPr) {
    pPr = ownerDoc.createElementNS(w, "w:pPr")
    p.insertBefore(pPr, p.firstChild)
  }
  let rPr = firstChildNS(pPr, w, "rPr")
  if (!rPr) {
    rPr = ownerDoc.createElementNS(w, "w:rPr")
    // rPr is the LAST child of pPr per the schema.
    pPr.appendChild(rPr)
  }
  // No-op if a <w:del/> is already there (idempotent — repeated edits don't
  // double-mark).
  for (const c of getChildren(rPr)) {
    if (c.namespaceURI === w && c.localName === "del") return
  }
  const del = ownerDoc.createElementNS(w, "w:del")
  del.setAttributeNS(w, "w:id", String(ctx.nextId()))
  del.setAttributeNS(w, "w:author", ctx.author)
  del.setAttributeNS(w, "w:date", ctx.date)
  rPr.appendChild(del)
}

/**
 * Wrap every <w:r> child of paragraph `p` inside a single <w:del>, marking
 * the run content as deleted. The paragraph mark itself is NOT marked here —
 * call `markParagraphMarkDeleted` separately if the whole paragraph is going.
 */
export function wrapParagraphContentInDel(p: Element, ownerDoc: Document, ctx: TrackContext): void {
  const runs = getChildrenNS(p, w, "r")
  if (runs.length === 0) return
  // Build the wrapper using the first run's parent as anchor; remove runs
  // from the paragraph in iteration order, group inside <w:del>, then
  // re-insert <w:del> at the original position.
  const firstRun = runs[0]!
  const before = firstRun.previousSibling
  const detached: Element[] = []
  for (const r of runs) {
    p.removeChild(r)
    detached.push(r)
  }
  const del = wrapDeletion(detached, ownerDoc, ctx)
  if (before && before.nextSibling) p.insertBefore(del, before.nextSibling)
  else if (before) p.appendChild(del)
  else p.insertBefore(del, p.firstChild)
}

/**
 * Wrap every <w:r> child of paragraph `p` inside a single <w:ins>, and mark
 * the paragraph mark as inserted (`<w:pPr><w:rPr><w:ins/>`). Used when a
 * fresh paragraph is being added under track-changes — without the mark
 * insertion, accept-all leaves the paragraph break as if it were original.
 */
export function markParagraphAsInserted(p: Element, ownerDoc: Document, ctx: TrackContext): void {
  const runs = getChildrenNS(p, w, "r")
  if (runs.length > 0) {
    const firstRun = runs[0]!
    const before = firstRun.previousSibling
    const detached: Element[] = []
    for (const r of runs) {
      p.removeChild(r)
      detached.push(r)
    }
    const ins = wrapInsertion(detached, ownerDoc, ctx)
    if (before && before.nextSibling) p.insertBefore(ins, before.nextSibling)
    else if (before) p.appendChild(ins)
    else p.insertBefore(ins, p.firstChild)
  }
  // Mark the paragraph mark inserted regardless of whether content runs
  // existed — empty paragraphs (page break, horizontal rule) still need the
  // mark itself flagged so accept-all keeps them.
  let pPr = firstChildNS(p, w, "pPr")
  if (!pPr) {
    pPr = ownerDoc.createElementNS(w, "w:pPr")
    p.insertBefore(pPr, p.firstChild)
  }
  let rPr = firstChildNS(pPr, w, "rPr")
  if (!rPr) {
    rPr = ownerDoc.createElementNS(w, "w:rPr")
    pPr.appendChild(rPr)
  }
  for (const c of getChildren(rPr)) {
    if (c.namespaceURI === w && c.localName === "ins") return
  }
  const insMark = ownerDoc.createElementNS(w, "w:ins")
  insMark.setAttributeNS(w, "w:id", String(ctx.nextId()))
  insMark.setAttributeNS(w, "w:author", ctx.author)
  insMark.setAttributeNS(w, "w:date", ctx.date)
  rPr.appendChild(insMark)
}

/* ------------- format-level wrappers ------------- */

/**
 * Append a <w:rPrChange> snapshot to a freshly-mutated <w:rPr>. Snapshot is
 * the *previous* state of the rPr (before mutation); pass null when there
 * was no rPr at all (in which case <w:rPrChange> contains an empty <w:rPr/>
 * to indicate "previous state was nothing").
 */
export function attachRPrChange(
  rPr: Element,
  previousRPrSnapshot: Element | null,
  ownerDoc: Document,
  ctx: TrackContext,
): void {
  if (!ctx.enabled) return
  const change = ownerDoc.createElementNS(w, "w:rPrChange")
  change.setAttributeNS(w, "w:id", String(ctx.nextId()))
  change.setAttributeNS(w, "w:author", ctx.author)
  change.setAttributeNS(w, "w:date", ctx.date)
  // The schema requires the previous <w:rPr> as a child — clone it (or
  // create empty if the original had none). Cloning here means callers
  // don't need to worry about whether their snapshot is still attached
  // somewhere — we copy.
  const inner = previousRPrSnapshot
    ? (previousRPrSnapshot.cloneNode(true) as Element)
    : ownerDoc.createElementNS(w, "w:rPr")
  change.appendChild(inner)
  rPr.appendChild(change)
}

/**
 * Same idea for paragraph properties. Snapshot of the previous <w:pPr> goes
 * inside <w:pPrChange>. Schema: <w:pPrChange> is the LAST child of <w:pPr>,
 * after even <w:rPr>; the wrapped snapshot may include its own <w:rPr> as
 * the previous paragraph-mark formatting.
 */
export function attachPPrChange(
  pPr: Element,
  previousPPrSnapshot: Element | null,
  ownerDoc: Document,
  ctx: TrackContext,
): void {
  if (!ctx.enabled) return
  const change = ownerDoc.createElementNS(w, "w:pPrChange")
  change.setAttributeNS(w, "w:id", String(ctx.nextId()))
  change.setAttributeNS(w, "w:author", ctx.author)
  change.setAttributeNS(w, "w:date", ctx.date)
  const inner = previousPPrSnapshot
    ? (previousPPrSnapshot.cloneNode(true) as Element)
    : ownerDoc.createElementNS(w, "w:pPr")
  change.appendChild(inner)
  pPr.appendChild(change)
}
