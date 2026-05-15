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

import { DocxReader } from "@lib/xml/reader.ts"
import { NS, type ParsedParagraph, type SectionInfo } from "@lib/parse/types.ts"
import { sectionForParagraph, sectionUsableWidthTwips } from "@lib/parse/section-metrics.ts"
import { firstChildNS, getChildren, getChildrenNS } from "@lib/xml/xml-utils.ts"
import { PPR_CHILD_ORDER, insertChildInOrder } from "@lib/xml/xml-order.ts"
import {
  assertNever,
  makeTrackContext,
  type EditOp,
  type Fragment,
  type ResolvedEdit,
  type ResolvedTarget,
  type TrackContext,
} from "@lib/config/edit-types.ts"
import {
  buildResolverContext,
  resolveLocator,
  resolveRunLocator,
  trailingBodySectPr,
  type ResolverContext,
} from "@lib/edit/locator.ts"
import {
  detectBlockers,
  explainBlockerReason,
  summarizeBlockers,
  type BlockerScan,
} from "@lib/edit/blockers.ts"
import {
  buildPPrChildren,
  buildRPrChildren,
  emitFragment,
  PPR_MANAGED_LOCAL_NAMES,
  RPR_MANAGED_LOCAL_NAMES,
  type EmitContext,
} from "@lib/edit/fragment-emit.ts"
import { normalizeTableSequencing } from "@lib/edit/table-emit.ts"
import { prepareLatex } from "@lib/edit/math/latex-to-omml.ts"
import {
  attachPPrChange,
  attachRPrChange,
  markParagraphAsInserted,
  markParagraphMarkDeleted,
  wrapParagraphContentInDel,
} from "@lib/edit/track-changes.ts"
import { DocxAssetRegistry } from "@lib/edit/asset-registry.ts"
import { emitHyperlinkNode, ensureHyperlinkCharStyle } from "@lib/edit/hyperlink.ts"
import { BookmarkAllocator } from "@lib/edit/bookmark.ts"
import {
  emitRefField,
  switchesForDisplay,
  type PendingRefBackfill,
} from "@lib/edit/fields/ref-field.ts"
import type {
  PendingCaptionFill,
  PendingCaptionReset,
  ResolvedCaptionConfig,
} from "@lib/edit/caption-counter.ts"
import { applyEditCaption, resolveEditCaptionTarget } from "@lib/edit/edit-caption-op.ts"

const w = NS.w

/* ------------- public entry ------------- */

export interface ApplyEditsReport {
  applied: number
  trackChanges: boolean
  blockerCounts: Record<"tracked-change" | "field" | "sdt", number>
  perOp: Array<{ index: number; op: EditOp["op"]; touched: number }>
}

/** Per-op summary returned by the dry-run preview path: locator resolution
 * only, no mutation. The agent reads this to verify edits will land where
 * intended (and which paragraphs they'll touch — so implicit-keep counts
 * can subtract them and stop reading as false-positive coverage gaps). */
export interface EditsPreviewEntry {
  index: number
  op: EditOp["op"]
  /** Resolved target paragraphs. For replace / delete / format these are
   * the paragraphs that will be mutated or removed. For insert-* this is
   * the anchor (no mutation; new paragraphs land before/after). */
  targetParaIndices: number[]
  /** Paragraphs that will be replaced or deleted (subset of targetParaIndices
   * for replace/delete; empty for insert/format). */
  willReplaceOrDeleteIndices: number[]
  /** Number of new paragraphs the op will insert. Approximate for image /
   * page-break / hr blocks (each emits one paragraph). */
  willInsertCount: number
  /** Container kind — body or table cell. */
  container: "body" | "cell"
}

export interface PreviewEditsInput {
  documentDoc: Document
  parsedParagraphs: ParsedParagraph[]
  edits: EditOp[]
}

export interface PreviewEditsOutput {
  entries: EditsPreviewEntry[]
  /** Paragraph indices that edits will replace/delete — used by apply-styles
   * dry-run path to subtract from implicit-keep so the count reflects
   * post-edits state. */
  replacedOrDeletedIndices: Set<number>
}

/**
 * Resolve all edit locators against the (pre-edits) document and report what
 * each op will touch — without mutating. Used by `apply-styles --dry-run` so
 * the change report can show predicted edit effect alongside style /
 * pattern_rule effects.
 *
 * Throws on locator-resolution errors and on blocker conflicts (same checks
 * the real apply path does pre-mutation).
 */
export function previewEditOps(input: PreviewEditsInput): PreviewEditsOutput {
  const { documentDoc, parsedParagraphs, edits } = input
  const resolverCtx = buildResolverContext(documentDoc, parsedParagraphs)
  const blockers = detectBlockers(documentDoc, resolverCtx.indexByElement)

  const resolved: ResolvedEdit[] = []
  for (const [i, op] of edits.entries()) {
    try {
      resolved.push(resolveOneEdit(op, resolverCtx, i))
    } catch (err) {
      throw new Error(`edits[${i}] (${op.op}): ${(err as Error).message}`, { cause: err })
    }
  }
  validateAgainstBlockers(resolved, blockers, resolverCtx.indexByElement)

  const replacedOrDeletedIndices = new Set<number>()
  const entries: EditsPreviewEntry[] = []
  for (const [i, edit] of resolved.entries()) {
    const targetParaIndices = edit.target.paragraphs.map(
      (p) => resolverCtx.indexByElement.get(p) ?? -1,
    )
    const op = edit.op.op
    let willReplaceOrDeleteIndices: number[] = []
    let willInsertCount = 0
    if (op === "replace") {
      willReplaceOrDeleteIndices = [...targetParaIndices]
      willInsertCount = edit.op.with.length
      for (const idx of targetParaIndices) if (idx >= 0) replacedOrDeletedIndices.add(idx)
    } else if (op === "delete") {
      willReplaceOrDeleteIndices = [...targetParaIndices]
      for (const idx of targetParaIndices) if (idx >= 0) replacedOrDeletedIndices.add(idx)
    } else if (op === "insert-before" || op === "insert-after") {
      willInsertCount = edit.op.content.length
    }
    const container: "body" | "cell" =
      edit.target.container.namespaceURI === w && edit.target.container.localName === "tc"
        ? "cell"
        : "body"
    entries.push({
      index: i,
      op,
      targetParaIndices,
      willReplaceOrDeleteIndices,
      willInsertCount,
      container,
    })
  }
  return { entries, replacedOrDeletedIndices }
}

/* ------------- exported in-memory entry (for apply-styles integration) ------------- */

export interface RunEditOpsInput {
  documentDoc: Document
  parsedParagraphs: ParsedParagraph[]
  reader: DocxReader
  edits: EditOp[]
  trackChanges: boolean
  /** Live styles.xml with this apply's numbering bindings applied —
   * used by InlineRef to verify the target paragraph's style cascade
   * resolves to a numId. Optional: when absent, the check falls back
   * to the parsed paragraph's pre-apply pPr.numId, which still
   * catches source-bound targets but may reject newly-bound targets.  */
  stylesDoc?: Document | null
  /** Parsed section list — used to look up the host section per op and
   * derive its usable content width (LaTeX `\textwidth`). Drives autofit
   * gridCol seeds for inserted tables so they don't overflow narrow page
   * setups. Optional: when absent, table emit falls back to a conservative
   * A4-tight constant. */
  sections?: SectionInfo[]
  /** Resolved captions config from apply config. Engine wires this into
   * EmitContext.resolveCaption so caption-bearing blocks (EquationBlock
   * with captionId, CaptionBlock, CaptionCounterReset) can resolve their
   * identifier at emit time. Absent → caption blocks throw at emit. */
  captions?: Map<string, ResolvedCaptionConfig>
}

export interface RunEditOpsOutput {
  imageRegistry: DocxAssetRegistry
  report: ApplyEditsReport
  /** Cross-reference state produced by `InlineRef` emissions during this
   * edit pass. apply-styles consumes these post-edits to (a) backfill REF
   * placeholder text from the numbering counter simulator, (b) commit
   * bookmark wrapping on target paragraphs, and (c) flip settings.xml's
   * updateFields flag when at least one ref was emitted. Empty when no
   * InlineRef appeared in the edits[]. */
  crossRefs: {
    bookmarkAllocator: BookmarkAllocator
    pendingBackfills: PendingRefBackfill[]
    pendingCaptionFills: PendingCaptionFill[]
    pendingCaptionResets: PendingCaptionReset[]
  }
}

/**
 * In-memory edit-ops application — for callers that already have the docx
 * open and parsed (e.g. `apply` integrating edits into its single
 * pipeline). The caller owns file I/O, validation, and writing.
 *
 * Mutates `documentDoc` in place. Returns the image registry (caller
 * flushes its staged binaries / rels / content-type updates into its
 * own replacement map) plus the applied-ops report.
 */
export async function runEditOps(input: RunEditOpsInput): Promise<RunEditOpsOutput> {
  const { documentDoc, parsedParagraphs, reader, edits, trackChanges, stylesDoc } = input

  const resolverCtx = buildResolverContext(documentDoc, parsedParagraphs)
  const blockers = detectBlockers(documentDoc, resolverCtx.indexByElement)

  const resolved: ResolvedEdit[] = []
  for (const [i, op] of edits.entries()) {
    try {
      resolved.push(resolveOneEdit(op, resolverCtx, i))
    } catch (err) {
      throw new Error(`edits[${i}] (${op.op}): ${(err as Error).message}`, { cause: err })
    }
  }

  validateAgainstBlockers(resolved, blockers, resolverCtx.indexByElement)

  // TrackChanges + new TableBlock / equation content: not supported.
  // OOXML's tracked-changes model wraps run / paragraph mutation; there's no
  // equivalent container for "table inserted" or "equation inserted" — at
  // best each cell paragraph or math element would appear as a separate /
  // partial insertion, which agents won't read correctly. Throw at the
  // boundary so it's a clear contract, not a half-rendered output.
  if (trackChanges) {
    for (const [i, op] of edits.entries()) {
      const frag =
        op.op === "replace"
          ? op.with
          : op.op === "insert-before" || op.op === "insert-after"
            ? op.content
            : null
      if (!frag) continue
      if (fragmentContainsTable(frag)) {
        throw new Error(
          `edits[${i}] (${op.op}): inserting a TableBlock under trackChanges=true is not supported. ` +
            `Run table insertion in a separate apply without trackChanges, then use trackChanges for subsequent cell edits.`,
        )
      }
      if (fragmentContainsEquation(frag)) {
        throw new Error(
          `edits[${i}] (${op.op}): inserting an equation (block or inline) under trackChanges=true is not supported. ` +
            `Run equation insertion in a separate apply without trackChanges, then use trackChanges for surrounding edits.`,
        )
      }
    }
  }

  // Pre-resolve every LaTeX expression to OMML before the synchronous emit
  // chain runs. Per-edit dispatch keeps the rethrown error pointing at the
  // offending `edits[N]` index — without this the temml / mml2omml stack
  // bubbles up with the latex source but no caller context.
  for (const [i, op] of edits.entries()) {
    const frag = fragmentOf(op)
    if (!frag) continue
    const items: LatexItem[] = []
    for (const b of frag) items.push(...collectLatexFromBlock(b))
    if (items.length === 0) continue
    try {
      await prepareLatex(items)
    } catch (err) {
      throw new Error(`edits[${i}] (${op.op}): ${(err as Error).message}`, { cause: err })
    }
  }

  const trackContext = makeTrackContext(trackChanges)
  const stale = new Set<Element>()
  const imageRegistry = await DocxAssetRegistry.open(reader)
  const bookmarkAllocator = new BookmarkAllocator(documentDoc)

  // Pre-scan declared anchors so forward refs (a ref citing an anchor in a
  // later edit, or in a later Block of the same op) resolve correctly.
  // Reserves the name + captures a numbering hint (styleId + whether the
  // block declares `numbering` directly) so a forward ref can answer
  // "target is auto-numbered?" before the target element exists. Run
  // before the main applyOne loop so duplicate / colliding names surface
  // ahead of any mutation.
  for (const op of edits) {
    const fragment = fragmentOf(op)
    if (!fragment) continue
    walkBlocksForAnchors(fragment, (hint) => {
      bookmarkAllocator.reserveName(hint.anchor, {
        styleId: hint.styleId,
        directlyNumbered: hint.directlyNumbered,
        isCaption: hint.isCaption,
      })
    })
  }

  const pendingBackfills: PendingRefBackfill[] = []
  const pendingCaptionFills: PendingCaptionFill[] = []
  const pendingCaptionResets: PendingCaptionReset[] = []
  const captionsMap = input.captions
  const emitCtx: EmitContext = {
    emitImage: (src, width, height, alt, ownerDoc) => {
      const { rId } = imageRegistry.registerImage(src)
      return imageRegistry.buildDrawing(rId, width, height, alt, ownerDoc)
    },
    emitHyperlink: (link, text, format, ownerDoc) => {
      // Inject the Hyperlink character style on first hyperlink emit only —
      // sparse-by-design (no style added when no hyperlink declared). The
      // helper is idempotent; the side-effect on stylesDoc is what we gate.
      // stylesDoc may be absent on edits-only callers without an apply
      // context — in that case the run still emits, just without the
      // character-style guarantee.
      if (stylesDoc) ensureHyperlinkCharStyle(stylesDoc)
      return emitHyperlinkNode(ownerDoc, link, text, format, imageRegistry)
    },
    captions: captionsMap
      ? {
          resolve: (identifier: string) => captionsMap.get(identifier),
          allocateBookmark: (name) => bookmarkAllocator.allocateRangeBookmark(name),
          bindBookmark: (name, pEl) => bookmarkAllocator.bindRangeBookmark(name, pEl),
          registerFill: (fill) => pendingCaptionFills.push(fill),
          registerReset: (reset) => pendingCaptionResets.push(reset),
        }
      : undefined,
    emitRef: (ref, ownerDoc, defaultFormat) => {
      // Resolve target → (paragraph element OR forward-ref name, bookmark
      // name). Two locator forms:
      //   - paragraph index: pre-edit, validates against the indexed
      //     paragraph map. Always element-bound at emit time.
      //   - anchor name: looks up adopted anchors / source bookmarks via
      //     the allocator. A forward ref (anchor declared later in edits[]
      //     or later in this op's Block list) finds the name in the
      //     pre-scan reservation table — `resolveByName` returns null, but
      //     `isReserved` is true, and the numbering check runs against the
      //     reserved record's predicted styleId instead of an element.
      let targetEl: Element | null
      let bookmarkName: string
      if (ref.refTo.type === "paragraph") {
        const target = resolverCtx.indexed[ref.refTo.index - 1]
        if (!target || target.index !== ref.refTo.index) {
          throw new Error(
            `InlineRef: refTo.paragraph=${ref.refTo.index} is out of range. ` +
              `Document has ${resolverCtx.indexed.length} indexed paragraphs.`,
          )
        }
        targetEl = target.element
        bookmarkName = bookmarkAllocator.getOrAllocate(targetEl).name
      } else {
        const found = bookmarkAllocator.resolveByName(ref.refTo.name)
        const reserved = !found && bookmarkAllocator.isReserved(ref.refTo.name)
        if (!found && !reserved) {
          throw new Error(
            `InlineRef: refTo.anchor="${ref.refTo.name}" was not found. ` +
              `Declare a ParagraphBlock.anchor (or EquationBlock.anchor) with this name somewhere in edits[], ` +
              `or reference an existing bookmark on a paragraph in the source document.`,
          )
        }
        targetEl = found?.element ?? null
        bookmarkName = ref.refTo.name
      }
      const display = ref.display ?? "label"
      // Target must be bound to a numbering scheme for label / number
      // (REF \n and \r switches render the lvlText / counter). The "full"
      // display resolves to the bookmark's text content, which any
      // non-empty paragraph supports — so we relax the check there.
      //
      // Caption-class anchors are SEQ-numbered (not numPr) and pass
      // unconditionally: forward refs match via reserveName's
      // `directlyNumbered: !!b.captionId` hint; backward refs match via
      // the `captionAnchorNames` set populated at caption emit time.
      // Backward refs: caption emit already called allocateRangeBookmark
      // → bookmarkAllocator.isRangeBookmark catches it.
      // Forward refs: bookmark not yet allocated; pre-scan reservation
      // carries `isCaption` from the block's captionId presence.
      const isCaptionAnchor =
        ref.refTo.type === "anchor" &&
        (bookmarkAllocator.isRangeBookmark(ref.refTo.name) ||
          bookmarkAllocator.predictedNumberingFor(ref.refTo.name)?.isCaption === true)
      // Caption-class targets: display:"full" would need a paragraph-wide
      // secondary bookmark for REF \h to return body text. The pipeline
      // only emits the primary bookmark (number + decoration), so
      // display:"full" diverges between pre-F9 placeholder and Word's
      // post-F9 render. Throw rather than ship divergent output. Agents
      // citing a caption use display:"label" (the rendered "(2.1)" or
      // "图 2.1" is the canonical citation form anyway).
      if (display === "full" && isCaptionAnchor) {
        throw new Error(
          `InlineRef: display="full" is not supported on caption-class anchors ` +
            `(anchor "${(ref.refTo as { name: string }).name}"). Caption refs use the SEQ-rendered ` +
            `text (prefix + chapter + counter + suffix); switch to display: "label".`,
        )
      }
      if (display !== "full" && !isCaptionAnchor) {
        const predicted =
          ref.refTo.type === "anchor"
            ? bookmarkAllocator.predictedNumberingFor(ref.refTo.name)
            : undefined
        const ok = targetIsAutoNumbered(
          {
            element: targetEl,
            styleId: predicted?.styleId,
            directlyNumbered: predicted?.directlyNumbered,
          },
          stylesDoc ?? null,
        )
        if (!ok) {
          const where =
            ref.refTo.type === "paragraph"
              ? `target paragraph #${ref.refTo.index}`
              : `target of anchor "${ref.refTo.name}"`
          throw new Error(
            `InlineRef: ${where} is not bound to a numbering scheme. ` +
              `display="${display}" requires an auto-numbered target (Word's \\n / \\r switches render from the numbering binding). ` +
              `Either bind the target's pStyle to a numbering[] level, or set display: "full" to use the paragraph's body text instead.`,
          )
        }
      }
      // Placeholder text is empty here; backfilled post-edit once the
      // numbering counter simulator yields rendered values. settings.xml's
      // updateFields=true also ensures Word resolves on open, so users who
      // skip the backfill (e.g. parser called outside the apply pipeline)
      // still see correct text after their first F9.
      // Caption-class anchors are SEQ-numbered, not numPr-bound — `\n`
      // would fail to resolve (no numbering binding). Use `\h` only so
      // REF reads the bookmark contents directly (which is the number
      // + decoration range that caption emit wrapped).
      const switches = isCaptionAnchor ? ["\\h"] : switchesForDisplay(display)
      const { runs, resultTextEl } = emitRefField(ownerDoc, {
        bookmarkName,
        switches,
        placeholder: "",
        format: ref.format ?? defaultFormat,
      })
      pendingBackfills.push({
        placeholderTextEl: resultTextEl,
        targetName: bookmarkName,
        display,
      })
      return runs
    },
    adoptAnchor: (name, pEl) => {
      bookmarkAllocator.adoptName(name, pEl)
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
    const perOpCtx: EmitContext = {
      ...emitCtx,
      usableWidthTwips: usableWidthForTarget(edit.target, input.sections, resolverCtx),
    }
    // Wrap with edits[N] context so emit-side throws (caption-not-declared,
    // adoptAnchor invariants, image emitter wiring, equation conversion)
    // surface with the locator the agent wrote. Without this the message
    // bubbles up bare ("captionId X is not declared") and the agent has no
    // way to locate the offending op in a multi-edit config.
    let touched: number
    try {
      touched = applyOne(edit, documentDoc, trackContext, perOpCtx, stale, resolverCtx, {
        bookmarkAllocator,
        captionsMap: input.captions,
      })
    } catch (err) {
      const msg = (err as Error).message
      // Don't double-wrap when applyOne (or resolveOneEdit upstream) already
      // emitted an `edits[N]` prefix.
      if (msg.startsWith(`edits[${i}]`)) throw err
      throw new Error(`edits[${i}] (${edit.op.op}): ${msg}`, { cause: err })
    }
    perOp.push({ index: i, op: edit.op.op, touched })
  }

  return {
    imageRegistry,
    report: {
      applied: resolved.length,
      trackChanges: trackContext.enabled,
      blockerCounts: summarizeBlockers(blockers),
      perOp,
    },
    crossRefs: {
      bookmarkAllocator,
      pendingCaptionFills,
      pendingCaptionResets,
      pendingBackfills,
    },
  }
}

/** Single auto-numbering check for InlineRef targets. The target is
 * represented either as a live paragraph element (already emitted — direct
 * `<w:p>/<w:pPr>/<w:numPr>` and pStyle cascade both observable) or as a
 * pre-scan hint (forward ref, target not yet emitted — `directlyNumbered`
 * mirrors the Block's direct `numbering` field, `styleId` runs the cascade
 * walk). Both representations now cover the same fact set; forward refs
 * are no longer blind to direct-numbering targets. */
function targetIsAutoNumbered(
  spec: { element: Element | null; styleId?: string; directlyNumbered?: boolean },
  stylesDoc: Document | null,
): boolean {
  if (spec.element) {
    const pPr = firstChildNS(spec.element, w, "pPr")
    if (pPr) {
      if (firstChildNS(pPr, w, "numPr")) return true
      const pStyle = firstChildNS(pPr, w, "pStyle")
      if (pStyle && stylesDoc) {
        const styleId = pStyle.getAttributeNS(w, "val") ?? pStyle.getAttribute("w:val") ?? ""
        if (styleId && styleHasNumPrInCascade(stylesDoc, styleId)) return true
      }
    }
    return false
  }
  if (spec.directlyNumbered) return true
  if (spec.styleId && stylesDoc && styleHasNumPrInCascade(stylesDoc, spec.styleId)) return true
  return false
}

function styleHasNumPrInCascade(stylesDoc: Document, styleId: string): boolean {
  const seen = new Set<string>()
  let current: string | null = styleId
  while (current && !seen.has(current)) {
    seen.add(current)
    const styleEl = findStyleById(stylesDoc, current)
    if (!styleEl) return false
    const sPr = firstChildNS(styleEl, w, "pPr")
    if (sPr && firstChildNS(sPr, w, "numPr")) return true
    const basedOn = firstChildNS(styleEl, w, "basedOn")
    current = basedOn ? (basedOn.getAttributeNS(w, "val") ?? basedOn.getAttribute("w:val")) : null
  }
  return false
}

function findStyleById(stylesDoc: Document, id: string): Element | null {
  const root = stylesDoc.documentElement
  if (!root) return null
  for (const s of getChildrenNS(root, w, "style")) {
    if ((s.getAttributeNS(w, "styleId") ?? s.getAttribute("w:styleId")) === id) return s
  }
  return null
}

/** Resolve one op's locator. Single source of truth used by both
 * `previewEditOps` (dry-run path) and `runEditOps` — keeps set-run's
 * RunLocator special case in one place. */
function resolveOneEdit(op: EditOp, resolverCtx: ResolverContext, _opIndex: number): ResolvedEdit {
  if (op.op === "set-run") {
    const r = resolveRunLocator(op.at, resolverCtx)
    return {
      op,
      target: { paragraphs: [r.paragraph], container: resolverCtx.body },
      runRef: r.run,
    }
  }
  if (op.op === "edit-caption") {
    // Target resolution happens at apply time (needs the live doc to
    // walk SEQ-bearing paragraphs). Synthesize a resolved target with
    // the body element as a placeholder; applyOne re-resolves against
    // the post-emit state.
    return {
      op,
      target: { paragraphs: [], container: resolverCtx.body },
    }
  }
  return { op, target: resolveLocator(op.at, resolverCtx) }
}

/* ------------- per-op host section lookup ------------- */

/** Resolve the usable content width (twips) for the section the op's target
 * lives in. Returns `undefined` for cell-internal targets (we don't read
 * tcW in v1; emitters fall back to a constant) and when no section info
 * is supplied or no section claims the target's paragraph index. */
function usableWidthForTarget(
  target: ResolvedTarget,
  sections: SectionInfo[] | undefined,
  resolverCtx: ResolverContext,
): number | undefined {
  if (!sections || sections.length === 0) return undefined
  // Cell-internal ops: skip section lookup; cell width is the real
  // constraint, not the section's. Future improvement: read tcW.
  if (target.container.namespaceURI === w && target.container.localName === "tc") return undefined
  // Pick the first target paragraph's index; for "whole-body" inserts at
  // the end (no resolved paragraphs), use the last section.
  const firstP = target.paragraphs[0]
  if (firstP === undefined) {
    const last = sections[sections.length - 1]
    if (!last) return undefined
    const width = sectionUsableWidthTwips(last)
    return width > 0 ? width : undefined
  }
  const idx = resolverCtx.indexByElement.get(firstP)
  if (idx === undefined) return undefined
  const sec = sectionForParagraph(sections, idx)
  if (!sec) return undefined
  const width = sectionUsableWidthTwips(sec)
  return width > 0 ? width : undefined
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

/** Detect any TableBlock anywhere in a fragment (handles direct table
 * blocks; doesn't recurse into cell Block[] content because the schema
 * already forbids nested tables in cells). */
function fragmentContainsTable(fragment: Fragment): boolean {
  for (const b of fragment) {
    if ((b as { type?: string }).type === "table") return true
  }
  return false
}

function fragmentContainsEquation(fragment: Fragment): boolean {
  for (const b of fragment) {
    if (collectLatexFromBlock(b).length > 0) return true
  }
  return false
}

interface LatexItem {
  latex: string
  displayMode: boolean
}

function collectLatexFromBlock(b: Fragment[number]): LatexItem[] {
  const out: LatexItem[] = []
  if (b.type === "equation") {
    if (b.latex !== undefined) {
      out.push({ latex: b.latex, displayMode: true })
    }
    // omml escape hatch: no LaTeX to prewarm
  } else if (b.type === "paragraph") {
    if (Array.isArray(b.text)) {
      for (const piece of b.text) {
        if ("math" in piece) out.push({ latex: piece.math, displayMode: false })
      }
    }
  } else if (b.type === "table") {
    for (const row of b.rows) {
      for (const cell of row) {
        let cellContent: unknown
        if (typeof cell === "string" || Array.isArray(cell)) {
          cellContent = cell
        } else if (cell && typeof cell === "object" && "content" in cell) {
          cellContent = cell.content
        }
        if (Array.isArray(cellContent)) {
          for (const piece of cellContent) {
            if (piece && typeof piece === "object") {
              if ("math" in piece && typeof piece.math === "string") {
                out.push({ latex: piece.math, displayMode: false })
              } else if ("type" in piece) {
                out.push(...collectLatexFromBlock(piece as Fragment[number]))
              }
            }
          }
        }
      }
    }
  }
  return out
}

/** Return the inserted/replaced Block[] for an op, or null when the op has
 * no fragment (delete / format / set-run). Centralizes the
 * `with` / `content` discriminant used by the latex collector and the
 * anchor pre-scan. */
function fragmentOf(op: EditOp): Fragment | null {
  if (op.op === "replace") return op.with
  if (op.op === "insert-before" || op.op === "insert-after") return op.content
  return null
}

/** Visit every block in a fragment that carries an `anchor` field, recursing
 * into table cell content. Duck-types on `anchor` presence so future Block
 * variants with an `anchor` field are picked up automatically — no per-type
 * whitelist to keep in sync with the schema.
 *
 * The callback receives the anchor name plus a `numbering` hint extracted
 * from the block's `styleId` and direct `numbering` field (when present).
 * That hint is what the forward-ref numbering check consumes via
 * `targetIsAutoNumbered`. */
interface AnchorHint {
  anchor: string
  styleId: string | undefined
  directlyNumbered: boolean
  isCaption: boolean
}

function walkBlocksForAnchors(fragment: Fragment, visit: (hint: AnchorHint) => void): void {
  for (const block of fragment) walkBlockForAnchors(block, visit)
}

function walkBlockForAnchors(block: Fragment[number], visit: (hint: AnchorHint) => void): void {
  const b = block as Partial<{
    anchor: string
    styleId: string
    captionId: string
    numbering: { numId: string; level: number }
  }>
  if (typeof b.anchor === "string" && b.anchor) {
    const isCaption = !!b.captionId
    visit({
      anchor: b.anchor,
      styleId: b.styleId,
      // Caption blocks (CaptionBlock + EquationBlock with captionId) are
      // SEQ-numbered, not numPr-numbered. They satisfy InlineRef's
      // numbering check without participating in the numPr cascade.
      directlyNumbered: !!b.numbering || isCaption,
      isCaption,
    })
  }
  if (block.type === "table") {
    for (const row of block.rows) {
      for (const cell of row) {
        let cellContent: unknown
        if (typeof cell === "string" || Array.isArray(cell)) {
          cellContent = cell
        } else if (cell && typeof cell === "object" && "content" in cell) {
          cellContent = cell.content
        }
        if (!Array.isArray(cellContent)) continue
        for (const piece of cellContent) {
          if (piece && typeof piece === "object" && "type" in piece) {
            walkBlockForAnchors(piece as Fragment[number], visit)
          }
        }
      }
    }
  }
}

/* ------------- per-op apply ------------- */

interface ApplyDeps {
  bookmarkAllocator: BookmarkAllocator
  captionsMap: Map<string, ResolvedCaptionConfig> | undefined
}

function applyEditCaptionOp(
  op: Extract<EditOp, { op: "edit-caption" }>,
  documentDoc: Document,
  deps: ApplyDeps,
): number {
  if (!deps.captionsMap) {
    throw new Error(
      "edit-caption: captions table not declared in apply config. Add the identifier to `captions`.",
    )
  }
  const para = resolveEditCaptionTarget({
    documentDoc,
    target: op.target,
    text: op.text,
    bookmarkAllocator: deps.bookmarkAllocator,
    captionsConfigs: deps.captionsMap,
  })
  // Find the caption's identifier so we can fetch bodySeparator.
  // The target paragraph's pStyle maps back via the configs index.
  let config: ResolvedCaptionConfig | undefined
  if ("captionId" in op.target) {
    config = deps.captionsMap.get(op.target.captionId)
  } else {
    // Anchor target — walk captions configs and pick the one whose
    // paragraphStyleId matches this paragraph's pStyle.
    const pPr = firstChildNS(para, w, "pPr")
    const pStyle = pPr ? firstChildNS(pPr, w, "pStyle") : null
    const styleId = pStyle ? (pStyle.getAttributeNS(w, "val") ?? null) : null
    if (styleId) {
      for (const c of deps.captionsMap.values()) {
        if (c.paragraphStyleId === styleId) {
          config = c
          break
        }
      }
    }
  }
  if (!config) {
    throw new Error(
      "edit-caption: could not resolve caption config for the target paragraph. " +
        "The paragraph's pStyle doesn't match any captions[<id>].styleId.",
    )
  }
  applyEditCaption(para, op.text, config, documentDoc)
  return 1
}

function applyOne(
  edit: ResolvedEdit,
  documentDoc: Document,
  trackContext: TrackContext,
  emitCtx: EmitContext,
  stale: Set<Element>,
  resolverCtx: ResolverContext,
  deps: ApplyDeps,
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
    case "set-run":
      if (!edit.runRef) {
        throw new Error("set-run: missing resolved run reference (internal)")
      }
      return applySetRun(edit.runRef, edit.op, documentDoc, trackContext)
    case "edit-caption":
      return applyEditCaptionOp(edit.op, documentDoc, deps)
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

// (Previously isImageOnlyParagraph inferred image/page-break/horizontal-rule
// shape from emitted XML — but <w:pBdr> match would false-positive on a
// regular paragraph block whose paraFormat intentionally declared a border.
// MDF skip is now decided from the source Block.type at the inheritance
// call site, where the agent's declared intent is unambiguous.)

function inheritPPrFromAnchor(newP: Element, anchor: Element, ownerDoc: Document): void {
  const anchorPPr = firstChildNS(anchor, w, "pPr")
  if (!anchorPPr) return
  let newPPr = firstChildNS(newP, w, "pPr")
  // Index existing new-pPr children by localName so we can merge per element.
  const existingByName = new Map<string, Element>()
  if (newPPr) {
    for (const c of getChildren(newPPr)) {
      if (c.namespaceURI === w) existingByName.set(c.localName!, c)
    }
  }
  // When the new paragraph has an explicit pStyle, the style cascade governs
  // run-mark formatting (font, size, weight) — the anchor's paragraph-mark
  // rPr is no longer the right inheritance source. Cloning it would shadow
  // the style's intended values: a typical case is anchor.pPr-mark carrying
  // `<w:b/>` (form template's bold label slot), which then bold-pollutes
  // every body paragraph the agent inserts with `styleId: "BodyText"`.
  // Style-level `bold: false` can't override pPr-mark rPr (it's not run rPr,
  // not paragraph rPr — Word reads it as the carriage-return character's
  // formatting). So when styleId is given, skip the pPr-mark rPr inheritance
  // entirely and let the cascade do its job.
  // pStyle lives inside <w:pPr>, not as a direct child of <w:p>; check newPPr.
  const newHasPStyle = !!(newPPr && firstChildNS(newPPr, w, "pStyle"))
  const toClone: Element[] = []
  for (const c of getChildren(anchorPPr)) {
    if (c.namespaceURI !== w) continue
    // Skip the tracked-changes pPr snapshot — that's history, not content.
    if (c.localName === "pPrChange") continue
    if (c.localName === "rPr" && newHasPStyle) continue
    const existing = existingByName.get(c.localName!)
    if (existing) {
      // MDF: agent's paraFormat overrides only the attrs it explicitly sets;
      // unset attrs keep the anchor's values. Skipping by localName instead
      // would let a partial override wipe linked attrs (e.g. `spaceBefore: 6`
      // alone would drop anchor's <w:spacing w:line/lineRule>).
      mergeMissingAttrs(c, existing)
    } else {
      toClone.push(c.cloneNode(true) as Element)
    }
  }
  if (toClone.length === 0) return
  if (!newPPr) {
    newPPr = ownerDoc.createElementNS(w, "w:pPr")
    newP.insertBefore(newPPr, newP.firstChild)
  }
  // Schema-correct insertion: anchor's child can land between elements
  // buildPPrChildren already pushed (e.g. anchor `ind` after our `spacing`
  // and before our `jc`). Plain appendChild would mis-order them.
  for (const c of toClone) insertChildInOrder(newPPr, c, PPR_CHILD_ORDER)
}

/** Copy `source`'s attributes onto `target`, leaving any attribute already
 * set on `target` untouched. Namespace-aware. */
function mergeMissingAttrs(source: Element, target: Element): void {
  const attrs = source.attributes
  if (!attrs) return
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs[i]!
    const ns = a.namespaceURI
    const local = a.localName ?? a.name
    const has = ns ? target.hasAttributeNS(ns, local) : target.hasAttribute(local)
    if (has) continue
    if (ns) target.setAttributeNS(ns, a.name, a.value)
    else target.setAttribute(a.name, a.value)
  }
}

function inheritFormatForNewParagraphs(
  newEls: Element[],
  newBlocks: Fragment,
  anchor: Element | null,
  ownerDoc: Document,
): void {
  if (!anchor) return
  // newEls and newBlocks are emitted from the same fragment in order — zip
  // by index. Block.type carries the agent's declared intent; only paragraph
  // blocks should inherit anchor pPr. image / page-break / horizontal-rule
  // blocks emit a <w:p> shell but skip MDF — they're structural, not text.
  for (let i = 0; i < newEls.length; i++) {
    const el = newEls[i]!
    const block = newBlocks[i]!
    if (block.type !== "paragraph") continue
    if (!isParagraphElement(el)) continue
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
  // Track parents whose children we removed; after the pass we re-check any
  // <w:tc> cells that became empty and insert a placeholder paragraph. ECMA-376
  // 17.4.66 requires every table cell to contain at least one <w:p> — Word's
  // loader prompts to repair on open when a cell is empty.
  const touchedParents = new Set<Element>()
  for (const p of target.paragraphs) {
    if (trackContext.enabled) {
      // Strikethrough-mode: keep the paragraph in the tree, mark its content
      // and paragraph-mark deleted so Word's review UI shows it as a deletion
      // for the user to accept or reject.
      wrapParagraphContentInDel(p, documentDoc, trackContext)
      markParagraphMarkDeleted(p, documentDoc, trackContext)
    } else {
      const parent = p.parentNode as Element | null
      if (parent) {
        touchedParents.add(parent)
        parent.removeChild(p)
      }
      stale.add(p)
    }
  }
  for (const parent of touchedParents) {
    if (parent.namespaceURI === w && parent.localName === "tc") {
      ensureCellHasParagraph(parent, documentDoc)
    }
  }
  return target.paragraphs.length
}

/** A `<w:tc>` must contain at least one `<w:p>` per ECMA-376 17.4.66. After
 * deletes, if the cell has no paragraph child left, append an empty one so
 * Word doesn't flag the file as needing repair. The placeholder lands at the
 * end of the cell, after any existing `<w:tcPr>`. */
function ensureCellHasParagraph(tc: Element, ownerDoc: Document): void {
  if (firstChildNS(tc, w, "p")) return
  tc.appendChild(ownerDoc.createElementNS(w, "w:p"))
}

/* ------------- insert ------------- */

function applyInsertBefore(
  target: ResolvedTarget,
  fragment: Fragment,
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
  inheritFormatForNewParagraphs(newEls, fragment, anchor, documentDoc)
  if (trackContext.enabled) {
    for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext)
  }
  if (anchor) {
    const parent = anchor.parentNode! as Element
    for (const el of newEls) parent.insertBefore(el, anchor)
    normalizeTableSequencing(parent, documentDoc)
  } else {
    insertAtContainerEnd(target.container, newEls, resolverCtx)
    normalizeTableSequencing(target.container, documentDoc)
  }
  return newEls.length
}

function applyInsertAfter(
  target: ResolvedTarget,
  fragment: Fragment,
  documentDoc: Document,
  trackContext: TrackContext,
  emitCtx: EmitContext,
  resolverCtx: ResolverContext,
): number {
  const newEls = emitFragment(fragment, documentDoc, emitCtx)
  if (newEls.length === 0) return 0
  const anchor = target.paragraphs[target.paragraphs.length - 1] ?? null
  inheritFormatForNewParagraphs(newEls, fragment, anchor, documentDoc)
  if (trackContext.enabled) {
    for (const el of newEls) markParagraphAsInserted(el, documentDoc, trackContext)
  }
  if (anchor) {
    const parent = anchor.parentNode! as Element
    const next = anchor.nextSibling
    if (next) for (const el of newEls) parent.insertBefore(el, next)
    else for (const el of newEls) parent.appendChild(el)
    normalizeTableSequencing(parent, documentDoc)
  } else {
    insertAtContainerEnd(target.container, newEls, resolverCtx)
    normalizeTableSequencing(target.container, documentDoc)
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
  fragment: Fragment,
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
  inheritFormatForNewParagraphs(newEls, fragment, anchor, documentDoc)

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
  normalizeTableSequencing(parent as Element, documentDoc)
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
        for (const c of buildPPrChildren(op.paraFormat, documentDoc))
          insertChildInOrder(pPr, c, PPR_CHILD_ORDER)
      }
      attachPPrChange(pPr, snapshot, documentDoc, trackContext)
    }
    touched++
  }
  return touched
}

/* ------------- set-run -------------
 *
 * Replace text in a single run while preserving the run's rPr (font /
 * underline / size / etc.) and all sibling runs in the paragraph. The
 * archetypal use is filling a form-fill placeholder: paragraph reads
 * `[label-bold] [whitespace-blank-with-underline]` and the agent wants to
 * provide the value text — the underline run's rPr (the placeholder
 * shape) carries through to the rendered value automatically. Optional
 * `format` lets the agent override specific rPr fields on the targeted
 * run when needed; absent, the run's existing rPr is preserved verbatim. */
function applySetRun(
  runEl: Element,
  op: Extract<EditOp, { op: "set-run" }>,
  documentDoc: Document,
  trackContext: TrackContext,
): number {
  // Take a snapshot of the run's rPr for tracked-changes recording before
  // any mutation, matching the order discipline used elsewhere in this
  // engine (snapshot pre-mutation; later cloneNode would capture the new
  // state and produce wrong rPrChange entries).
  const oldRPr = firstChildNS(runEl, w, "rPr")
  const rPrSnapshot = oldRPr ? (oldRPr.cloneNode(true) as Element) : null

  // Optional format override — same pattern as applyFormat but scoped to
  // this one run: clear managed children, re-emit from buildRPrChildren.
  if (op.format) {
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
      runEl.insertBefore(rPr, runEl.firstChild)
    }
    for (const c of buildRPrChildren(op.format, documentDoc)) rPr.appendChild(c)
    attachRPrChange(rPr, rPrSnapshot, documentDoc, trackContext)
  }

  // Replace the run's text content. A run can hold multiple <w:t> /
  // <w:tab> / <w:br> children; we collapse them into one <w:t> carrying
  // the new text with xml:space="preserve" so leading/trailing whitespace
  // (common in form values) survives serialization. tab / br elements are
  // dropped — they belonged to the placeholder run's structure, not the
  // value's; agent supplies a string, not structured runs.
  for (const c of Array.from(getChildren(runEl))) {
    if (c.namespaceURI !== w) continue
    if (c.localName === "rPr") continue
    runEl.removeChild(c)
  }
  const t = documentDoc.createElementNS(w, "w:t")
  t.setAttribute("xml:space", "preserve")
  t.appendChild(documentDoc.createTextNode(op.with))
  runEl.appendChild(t)

  return 1
}
