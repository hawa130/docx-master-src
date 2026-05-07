/**
 * Edit-engine orchestrator.
 *
 * Pipeline:
 *   1. parseEditConfig (zod)             — shape + locally-derivable invariants
 *   2. Open docx, parse paragraphs        — needed for indexed-paragraph walk
 *   3. Build resolver context, blockers   — pre-mutation snapshot
 *   4. Resolve all locators upfront       — Element refs survive later mutations
 *   5. Validate against blockers          — refuse before touching the docx
 *   6. Apply ops in array order           — track stale Elements between ops
 *   7. Validate output XML, write fresh   — original file never touched
 *
 * Stale-element rule: when op A removes / replaces paragraphs and op B's
 * resolved target overlaps A's removed paragraphs, B fails with a clear
 * message. This is rarer than it sounds — agents typically construct
 * non-overlapping edits — but cheaper to catch than to debug.
 */

import { copyFileSync, existsSync, mkdirSync, unlinkSync } from "node:fs"
import { dirname } from "node:path"
import { DocxReader, serializeXml } from "@lib/reader.ts"
import { DocumentParser } from "@lib/document-parser.ts"
import { StyleResolver } from "@lib/style-resolver.ts"
import { NS } from "@lib/types.ts"
import { firstChildNS, getChildren, getChildrenNS } from "@lib/xml-utils.ts"
import { validateOutput } from "./docx-plumbing.ts"
import {
  assertNever,
  makeTrackContext,
  type EditConfig,
  type EditOp,
  type ResolvedEdit,
  type ResolvedTarget,
  type TrackContext,
} from "./edit-types.ts"
import { parseEditConfig } from "./edit-config-schema.ts"
import {
  buildResolverContext,
  resolveLocator,
  trailingBodySectPr,
  type ResolverContext,
} from "./locator.ts"
import {
  detectBlockers,
  explainBlockerReason,
  summarizeBlockers,
  type BlockerScan,
} from "./blockers.ts"
import {
  buildPPrChildren,
  buildRPrChildren,
  emitFragment,
  PPR_MANAGED_LOCAL_NAMES,
  RPR_MANAGED_LOCAL_NAMES,
  type EmitContext,
} from "./fragment-emit.ts"
import {
  attachPPrChange,
  attachRPrChange,
  markParagraphAsInserted,
  markParagraphMarkDeleted,
  wrapParagraphContentInDel,
} from "./track-changes.ts"
import { ImageAssetRegistry } from "./image-asset.ts"

const w = NS.w

/* ------------- public entry ------------- */

export interface ApplyEditsReport {
  applied: number
  trackChanges: boolean
  blockerCounts: Record<"tracked-change" | "field" | "sdt", number>
  perOp: Array<{ index: number; op: EditOp["op"]; touched: number }>
}

export async function applyEdits(
  source: string,
  output: string,
  rawConfig: unknown,
): Promise<ApplyEditsReport> {
  const config = parseEditConfig(rawConfig)

  // 1. Stage output (copy original; we mutate the copy so a failure leaves
  // the source untouched).
  mkdirSync(dirname(output), { recursive: true })
  copyFileSync(source, output)

  try {
    return await runOnCopy(output, config)
  } catch (err) {
    if (existsSync(output)) {
      try {
        unlinkSync(output)
      } catch {}
    }
    throw err
  }
}

async function runOnCopy(outputPath: string, config: EditConfig): Promise<ApplyEditsReport> {
  // 2. Open + parse for index/style awareness.
  const reader = await DocxReader.open(outputPath)
  const documentDoc = await reader.readXml("word/document.xml")
  if (!documentDoc) throw new Error("word/document.xml not found")
  const stylesDoc = await reader.readXml("word/styles.xml")
  const themeDoc = await reader.readXml("word/theme/theme1.xml")
  const numberingDoc = await reader.readXml("word/numbering.xml")

  const resolver = new StyleResolver(stylesDoc, themeDoc)
  if (stylesDoc) resolver.expandThemedFontsInStyles(stylesDoc)
  const parser = new DocumentParser(documentDoc, resolver, numberingDoc)
  const parsed = parser.parse()

  // 3. Resolver + blockers. indexByElement is built inside resolverCtx; reuse
  // it for blocker scanning so the blocker map can echo paragraph indices.
  const resolverCtx = buildResolverContext(documentDoc, parsed.paragraphs)
  const blockers = detectBlockers(documentDoc, resolverCtx.indexByElement)

  // 4. Pre-resolve every op's locator. Resolution failures throw with the
  // locator-specific message from locator.ts.
  const resolved: ResolvedEdit[] = []
  for (const [i, op] of config.edits.entries()) {
    let target: ResolvedTarget
    try {
      target = resolveLocator(op.at, resolverCtx)
    } catch (err) {
      throw new Error(`edits[${i}] (${op.op}): ${(err as Error).message}`, { cause: err })
    }
    resolved.push({ op, target })
  }

  // 5. Blocker validation. A target paragraph that's blocked rejects the op
  // before any mutation. Field / SDT / existing-tracked-change zones each
  // surface their reason.
  validateAgainstBlockers(resolved, blockers, resolverCtx.indexByElement)

  // 6. Apply ops. Track stale Elements so a later op that targets a removed
  // paragraph fails loudly. Image asset registry is opened lazily — most
  // edit passes don't touch images, and lazy init keeps the rels / content-
  // types XML untouched in the output zip when no images are added.
  const trackContext = makeTrackContext(config.trackChanges ?? false)
  const stale = new Set<Element>()
  const imageRegistry = await ImageAssetRegistry.open(reader)
  const emitCtx: EmitContext = {
    emitImage: (src, widthPt, heightPt, alt, ownerDoc) => {
      const { rId } = imageRegistry.registerImage(src)
      return imageRegistry.buildDrawing(rId, widthPt, heightPt, alt, ownerDoc)
    },
  }
  const perOp: ApplyEditsReport["perOp"] = []
  for (const [i, edit] of resolved.entries()) {
    for (const p of edit.target.paragraphs) {
      if (stale.has(p)) {
        throw new Error(
          `edits[${i}] (${edit.op.op}): targets a paragraph that was removed by an earlier edit. ` +
            `Reorder edits, or split into separate apply_edits runs so each pass sees a fresh document.`,
        )
      }
    }
    const touched = applyOne(edit, documentDoc, trackContext, emitCtx, stale, resolverCtx)
    perOp.push({ index: i, op: edit.op.op, touched })
  }

  // 7. Serialize, validate, write. Image registry pushes any new media /
  // rels / content-type updates into the replacements map; flushTo is a
  // no-op when no images were added in this pass.
  const replacements = new Map<string, string | Uint8Array>()
  replacements.set("word/document.xml", serializeXml(documentDoc))
  imageRegistry.flushTo(replacements)
  await reader.copyAndModify(outputPath, replacements)
  // Only validate XML/rels entries — image binaries will trip the parser.
  const xmlKeys = Array.from(replacements.keys()).filter(
    (k) => k.endsWith(".xml") || k.endsWith(".rels"),
  )
  const validation = await validateOutput(outputPath, xmlKeys)
  if (!validation.ok) {
    throw new Error(`output failed XML validation: ${validation.error}`)
  }

  return {
    applied: resolved.length,
    trackChanges: trackContext.enabled,
    blockerCounts: summarizeBlockers(blockers),
    perOp,
  }
}

/* ------------- blocker enforcement ------------- */

function validateAgainstBlockers(
  resolved: ResolvedEdit[],
  blockers: BlockerScan,
  indexByElement: Map<Element, number>,
): void {
  const failures: string[] = []
  for (const [i, edit] of resolved.entries()) {
    for (const p of edit.target.paragraphs) {
      const reason = blockers.byElement.get(p)
      if (reason) {
        const idx = indexByElement.get(p)
        const where = idx !== undefined ? `paragraph #${idx}` : `cell paragraph`
        failures.push(
          `edits[${i}] (${edit.op.op}): ${where} is blocked — ${explainBlockerReason(reason)}`,
        )
      }
    }
  }
  if (failures.length > 0) {
    throw new Error(
      `${failures.length} edit(s) targeted blocked paragraphs:\n  ${failures.join("\n  ")}`,
    )
  }
}

/* ------------- per-op apply ------------- */

function applyOne(
  edit: ResolvedEdit,
  documentDoc: Document,
  trackContext: TrackContext,
  emitCtx: EmitContext,
  stale: Set<Element>,
  resolverCtx: ResolverContext,
): number {
  switch (edit.op.op) {
    case "replace":
      return applyReplace(edit.target, edit.op.with, documentDoc, trackContext, emitCtx, stale)
    case "insert-before":
      return applyInsertBefore(
        edit.target,
        edit.op.content,
        documentDoc,
        trackContext,
        emitCtx,
        resolverCtx,
      )
    case "insert-after":
      return applyInsertAfter(
        edit.target,
        edit.op.content,
        documentDoc,
        trackContext,
        emitCtx,
        resolverCtx,
      )
    case "delete":
      return applyDelete(edit.target, documentDoc, trackContext, stale)
    case "format":
      return applyFormat(edit.target, edit.op, documentDoc, trackContext)
    default:
      return assertNever(edit.op)
  }
}

/* ------------- format inheritance (Match Destination Formatting) -------------
 *
 * When `replace` / `insert-before` / `insert-after` produce a fresh paragraph,
 * the slot's expected formatting (pStyle + pPr direct overrides) is what
 * the user wants to keep — same semantics as Word's "Match Destination
 * Formatting" paste mode. Block-level `styleId` / `format` are explicit
 * overrides; otherwise we inherit from the anchor.
 *
 * Anchor by op:
 *   - replace        → first replaced paragraph
 *   - insert-before  → first target paragraph
 *   - insert-after   → last target paragraph
 *
 * Inheritance is additive at pPr-child granularity: we only copy children
 * whose localName isn't already on the new paragraph, so an explicit
 * Block.format spacing/jc/etc. always wins. pPrChange is skipped (it's a
 * tracked-changes artifact, not content).
 *
 * Image / page-break / horizontal-rule blocks are NOT inherited into —
 * they have their own structural meaning (a horizontal rule's pBdr would
 * conflict with the anchor's borders, an image paragraph's centering is
 * its own decision). Only `paragraph` blocks adopt the destination format.
 */

function isParagraphElement(el: Element): boolean {
  return el.namespaceURI === w && el.localName === "p"
}

function isImageOnlyParagraph(p: Element): boolean {
  // Image / page-break / horizontal-rule blocks all emit a <w:p>, but the
  // ones we want to skip inheritance on have specific shapes:
  //   - image: <w:r><w:drawing>… inside, no <w:t>
  //   - page-break: <w:r><w:br type="page"/>
  //   - horizontal-rule: empty <w:r>s + <w:pBdr> in pPr
  // Detect: any <w:drawing> descendant, or a <w:br type="page">, or an
  // existing <w:pBdr> child of pPr.
  for (const r of getChildrenNS(p, w, "r")) {
    for (const c of getChildren(r)) {
      if (c.namespaceURI !== w) continue
      if (c.localName === "drawing") return true
      if (c.localName === "br") {
        const t = c.getAttributeNS(w, "type") || c.getAttribute("w:type")
        if (t === "page") return true
      }
    }
  }
  const pPr = firstChildNS(p, w, "pPr")
  if (pPr && firstChildNS(pPr, w, "pBdr")) return true
  return false
}

function inheritPPrFromAnchor(newP: Element, anchor: Element, ownerDoc: Document): void {
  const anchorPPr = firstChildNS(anchor, w, "pPr")
  if (!anchorPPr) return
  let newPPr = firstChildNS(newP, w, "pPr")
  const existingLocalNames = new Set<string>()
  if (newPPr) {
    for (const c of getChildren(newPPr)) {
      if (c.namespaceURI === w) existingLocalNames.add(c.localName!)
    }
  }
  const toInherit: Element[] = []
  for (const c of getChildren(anchorPPr)) {
    if (c.namespaceURI !== w) continue
    // Skip the tracked-changes pPr snapshot — that's history, not content.
    if (c.localName === "pPrChange") continue
    // Skip if the new paragraph already declared this property explicitly.
    if (existingLocalNames.has(c.localName!)) continue
    toInherit.push(c.cloneNode(true) as Element)
  }
  if (toInherit.length === 0) return
  if (!newPPr) {
    newPPr = ownerDoc.createElementNS(w, "w:pPr")
    newP.insertBefore(newPPr, newP.firstChild)
  }
  // Append rather than re-sort — Word reads mis-ordered pPr children
  // tolerantly, same convention as style-mutation.ts elsewhere in this repo.
  for (const c of toInherit) newPPr.appendChild(c)
}

function inheritFormatForNewParagraphs(
  newEls: Element[],
  anchor: Element | null,
  ownerDoc: Document,
): void {
  if (!anchor) return
  for (const el of newEls) {
    if (!isParagraphElement(el)) continue
    if (isImageOnlyParagraph(el)) continue
    inheritPPrFromAnchor(el, anchor, ownerDoc)
  }
}

/* ------------- delete ------------- */

function applyDelete(
  target: ResolvedTarget,
  documentDoc: Document,
  trackContext: TrackContext,
  stale: Set<Element>,
): number {
  if (target.paragraphs.length === 0) return 0
  for (const p of target.paragraphs) {
    if (trackContext.enabled) {
      // Strikethrough-mode: keep the paragraph in the tree, mark its content
      // and paragraph-mark deleted so Word's review UI shows it as a deletion
      // for the user to accept or reject.
      wrapParagraphContentInDel(p, documentDoc, trackContext)
      markParagraphMarkDeleted(p, documentDoc, trackContext)
    } else {
      const parent = p.parentNode
      if (parent) parent.removeChild(p)
      stale.add(p)
    }
  }
  return target.paragraphs.length
}

/* ------------- insert ------------- */

function applyInsertBefore(
  target: ResolvedTarget,
  fragment: import("./edit-types.ts").Fragment,
  documentDoc: Document,
  trackContext: TrackContext,
  emitCtx: EmitContext,
  resolverCtx: ResolverContext,
): number {
  const newEls = emitFragment(fragment, documentDoc, emitCtx)
  if (newEls.length === 0) return 0
  // Match destination formatting: inherit pPr from the first target paragraph
  // before any track-changes wrapping (so <w:ins> on the para mark sits on
  // top of the inherited rPr, not under it).
  const anchor = target.paragraphs[0] ?? null
  inheritFormatForNewParagraphs(newEls, anchor, documentDoc)
  if (trackContext.enabled) {
    for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext)
  }
  if (anchor) {
    const parent = anchor.parentNode!
    for (const el of newEls) parent.insertBefore(el, anchor)
  } else {
    insertAtContainerEnd(target.container, newEls, resolverCtx)
  }
  return newEls.length
}

function applyInsertAfter(
  target: ResolvedTarget,
  fragment: import("./edit-types.ts").Fragment,
  documentDoc: Document,
  trackContext: TrackContext,
  emitCtx: EmitContext,
  resolverCtx: ResolverContext,
): number {
  const newEls = emitFragment(fragment, documentDoc, emitCtx)
  if (newEls.length === 0) return 0
  const anchor = target.paragraphs[target.paragraphs.length - 1] ?? null
  inheritFormatForNewParagraphs(newEls, anchor, documentDoc)
  if (trackContext.enabled) {
    for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext)
  }
  if (anchor) {
    const parent = anchor.parentNode!
    const next = anchor.nextSibling
    if (next) for (const el of newEls) parent.insertBefore(el, next)
    else for (const el of newEls) parent.appendChild(el)
  } else {
    insertAtContainerEnd(target.container, newEls, resolverCtx)
  }
  return newEls.length
}

function insertAtContainerEnd(
  container: Element,
  newEls: Element[],
  resolverCtx: ResolverContext,
): void {
  // Body-level insertion respects the trailing <w:sectPr> sibling: append
  // before it so the section descriptor stays last. Cell containers (<w:tc>)
  // have no equivalent — append plainly.
  if (container === resolverCtx.body) {
    const sectPr = trailingBodySectPr(container)
    if (sectPr) {
      for (const el of newEls) container.insertBefore(el, sectPr)
      return
    }
  }
  for (const el of newEls) container.appendChild(el)
}

/* ------------- replace ------------- */

function applyReplace(
  target: ResolvedTarget,
  fragment: import("./edit-types.ts").Fragment,
  documentDoc: Document,
  trackContext: TrackContext,
  emitCtx: EmitContext,
  stale: Set<Element>,
): number {
  if (target.paragraphs.length === 0) {
    throw new Error("replace op resolved to zero paragraphs — locator must select at least one")
  }
  const newEls = emitFragment(fragment, documentDoc, emitCtx)
  const anchor = target.paragraphs[0]!
  const parent = anchor.parentNode!

  // Match destination formatting before any track-changes wrapping. Inherit
  // happens against the *original* anchor (still in the tree at this point);
  // when track-changes is on, the anchor's pPr is later mutated by
  // markParagraphMarkDeleted to add <w:del/>, but inheritance has already
  // captured a clone, so the new paragraphs see the unmutated formatting.
  inheritFormatForNewParagraphs(newEls, anchor, documentDoc)

  if (trackContext.enabled) {
    // 1. Mark old paragraphs deleted (keep in tree).
    for (const p of target.paragraphs) {
      wrapParagraphContentInDel(p, documentDoc, trackContext)
      markParagraphMarkDeleted(p, documentDoc, trackContext)
    }
    // 2. Tag new paragraphs as insertions.
    for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext)
    // 3. Insert new paragraphs before the (still-present) old anchor so the
    // accept-changes order keeps the inserted content in the right spot.
    for (const el of newEls) parent.insertBefore(el, anchor)
  } else {
    // Plain replace: insert new before anchor, remove old.
    for (const el of newEls) parent.insertBefore(el, anchor)
    for (const p of target.paragraphs) {
      parent.removeChild(p)
      stale.add(p)
    }
  }
  return newEls.length
}

/* ------------- format ------------- */

function applyFormat(
  target: ResolvedTarget,
  op: Extract<EditOp, { op: "format" }>,
  documentDoc: Document,
  trackContext: TrackContext,
): number {
  if (target.paragraphs.length === 0) {
    throw new Error("format op resolved to zero paragraphs — locator must select at least one")
  }
  let touched = 0
  for (const p of target.paragraphs) {
    if (op.runFormat) {
      for (const r of getChildrenNS(p, w, "r")) {
        const oldRPr = firstChildNS(r, w, "rPr")
        const snapshot = oldRPr ? (oldRPr.cloneNode(true) as Element) : null
        let rPr: Element
        if (oldRPr) {
          rPr = oldRPr
          for (const c of Array.from(getChildren(rPr))) {
            if (c.namespaceURI === w && RPR_MANAGED_LOCAL_NAMES.has(c.localName!)) {
              rPr.removeChild(c)
            }
          }
        } else {
          rPr = documentDoc.createElementNS(w, "w:rPr")
          r.insertBefore(rPr, r.firstChild)
        }
        for (const c of buildRPrChildren(op.runFormat, documentDoc)) rPr.appendChild(c)
        attachRPrChange(rPr, snapshot, documentDoc, trackContext)
      }
    }
    if (op.paraFormat || op.styleId) {
      const oldPPr = firstChildNS(p, w, "pPr")
      const snapshot = oldPPr ? (oldPPr.cloneNode(true) as Element) : null
      let pPr: Element
      if (oldPPr) {
        pPr = oldPPr
      } else {
        pPr = documentDoc.createElementNS(w, "w:pPr")
        p.insertBefore(pPr, p.firstChild)
      }
      if (op.styleId) {
        let pStyle = firstChildNS(pPr, w, "pStyle")
        if (!pStyle) {
          pStyle = documentDoc.createElementNS(w, "w:pStyle")
          pPr.insertBefore(pStyle, pPr.firstChild)
        }
        pStyle.setAttributeNS(w, "w:val", op.styleId)
      }
      if (op.paraFormat) {
        for (const c of Array.from(getChildren(pPr))) {
          if (c.namespaceURI === w && PPR_MANAGED_LOCAL_NAMES.has(c.localName!)) {
            pPr.removeChild(c)
          }
        }
        for (const c of buildPPrChildren(op.paraFormat, documentDoc)) pPr.appendChild(c)
      }
      attachPPrChange(pPr, snapshot, documentDoc, trackContext)
    }
    touched++
  }
  return touched
}
