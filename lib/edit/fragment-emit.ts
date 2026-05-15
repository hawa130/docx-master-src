/**
 * Fragment → OOXML elements.
 *
 * Each Block emits exactly one <w:p>. The emitter's job is mechanical
 * serialization; the content model (RichText, RunFormat, ParagraphFormat)
 * is what the agent declared. New block types or new format properties = a
 * new case here, contained.
 *
 * The format helpers (`buildRPrChildren`, `buildPPrChildren`) are shared
 * with edit-engine.ts — the format op clears managed children and appends
 * the same elements this emitter would produce for fresh paragraphs.
 *
 * Image emission is wired through a callback (`ImageEmitter`) so this
 * module stays free of the image-asset / zip-mutation concern. The engine
 * supplies the callback when image assets are available; absent callback +
 * `image` block = error (Step 7 plumbs this in via image-asset.ts).
 */

import { NS } from "@lib/parse/types.ts"
import type * as z from "zod/mini"
import type { InlineRefSchema } from "@lib/config/edit-config-schema.ts"
import {
  assertNever,
  type Block,
  type Fragment,
  type ParagraphFormat,
  type RichText,
  type RunFormat,
} from "@lib/config/edit-types.ts"
import {
  type Length,
  type LineSpacingInput,
  parseIndent,
  parseLineSpacing,
  toHalfPt,
  toTwips,
} from "@lib/shared/units.ts"
import { emitInlineField, emitInlineStyleRef } from "@lib/edit/fields/inline-fields.ts"
import { RPR_CHILD_ORDER } from "@lib/xml/xml-order.ts"
import { emitTableBlock } from "@lib/edit/table-emit.ts"
import { emitEquationBlock, emitInlineEquation } from "@lib/edit/math/equation-emit.ts"
import { emitCaptionBlock, emitCaptionReset } from "@lib/edit/caption-emit.ts"
import type { BookmarkRange } from "@lib/edit/bookmark.ts"
import type {
  PendingCaptionFill,
  PendingCaptionReset,
  ResolvedCaptionConfig,
} from "@lib/edit/caption-counter.ts"

export type InlineRef = z.infer<typeof InlineRefSchema>

/** Callback the engine supplies for InlineRef nodes. Returns the run
 * elements that go inside the paragraph (typically the 5-run REF
 * sequence from emitRefField). Decoupled from this module so fragment-emit
 * stays free of bookmark / locator / numbering-counter concerns. */
export type RefEmitter = (
  ref: InlineRef,
  ownerDoc: Document,
  defaultFormat: RunFormat | undefined,
) => Element[]

const w = NS.w

/** OOXML toggle element: presence-only when `on=true`, val="0" when false. */
function toggleElement(ownerDoc: Document, qname: string, on: boolean): Element {
  const el = ownerDoc.createElementNS(w, qname)
  if (!on) el.setAttributeNS(w, "w:val", "0")
  return el
}

/* ------------- run-level format ------------- */

export function buildRPrChildren(fmt: RunFormat, ownerDoc: Document): Element[] {
  const out: Element[] = []
  if (fmt.fontLatin || fmt.fontCJK) {
    const rFonts = ownerDoc.createElementNS(w, "w:rFonts")
    const ascii = fmt.fontLatin ?? fmt.fontCJK ?? ""
    const ea = fmt.fontCJK ?? fmt.fontLatin ?? ""
    if (ascii) {
      rFonts.setAttributeNS(w, "w:ascii", ascii)
      rFonts.setAttributeNS(w, "w:hAnsi", ascii)
    }
    if (ea) rFonts.setAttributeNS(w, "w:eastAsia", ea)
    out.push(rFonts)
  }
  if (fmt.size !== undefined) {
    const halfPt = toHalfPt(fmt.size, "size")
    const sz = ownerDoc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(halfPt))
    out.push(sz)
    const szCs = ownerDoc.createElementNS(w, "w:szCs")
    szCs.setAttributeNS(w, "w:val", String(halfPt))
    out.push(szCs)
  }
  // Toggles emit both states: `true` adds the on-marker, `false` adds
  // `w:val="0"` so it overrides an inherited / cascaded `on`. Without the
  // explicit off, MDF inheritance from a bold anchor would silently
  // re-bold the new content. `undefined` skips the property (style
  // cascade decides).
  if (fmt.bold !== undefined) {
    out.push(toggleElement(ownerDoc, "w:b", fmt.bold))
    out.push(toggleElement(ownerDoc, "w:bCs", fmt.bold))
  }
  if (fmt.italic !== undefined) {
    out.push(toggleElement(ownerDoc, "w:i", fmt.italic))
    out.push(toggleElement(ownerDoc, "w:iCs", fmt.italic))
  }
  if (fmt.underline !== undefined) {
    const u = ownerDoc.createElementNS(w, "w:u")
    u.setAttributeNS(w, "w:val", fmt.underline ? "single" : "none")
    out.push(u)
  }
  if (fmt.strike !== undefined) {
    out.push(toggleElement(ownerDoc, "w:strike", fmt.strike))
  }
  if (fmt.color) {
    const c = ownerDoc.createElementNS(w, "w:color")
    c.setAttributeNS(w, "w:val", fmt.color)
    out.push(c)
  }
  if (fmt.vertAlign) {
    const va = ownerDoc.createElementNS(w, "w:vertAlign")
    va.setAttributeNS(w, "w:val", fmt.vertAlign)
    out.push(va)
  }
  // CT_RPr children must appear in EG_RPrBase order (b/bCs before sz/szCs,
  // etc.). Callers append in returned order, so sort here once instead of
  // making each call site care.
  const orderIdx = (el: Element) => {
    const i = (RPR_CHILD_ORDER as ReadonlyArray<string>).indexOf(el.localName!)
    return i < 0 ? RPR_CHILD_ORDER.length : i
  }
  out.sort((a, b) => orderIdx(a) - orderIdx(b))
  return out
}

/** Run-property children that `buildRPrChildren` can produce — exposed so
 * edit-engine's `format` op can clear exactly these before re-applying the
 * builder's output, preserving anything the user didn't ask to change (lang,
 * kern, w, ...).
 *
 * Scope: run-level (from `RunFormat`). Wider than style-mutation's
 * RPR_MANAGED_CHILDREN — RunFormat has `u` and `strike` which
 * `StyleConfigEntry` does not. Don't fold the two sets together; their
 * scopes are intentionally different. */
export const RPR_MANAGED_LOCAL_NAMES: ReadonlySet<string> = new Set([
  "rFonts",
  "sz",
  "szCs",
  "b",
  "bCs",
  "i",
  "iCs",
  "u",
  "strike",
  "color",
  "vertAlign",
])

/* ------------- paragraph-level format ------------- */

export function buildPPrChildren(fmt: ParagraphFormat, ownerDoc: Document): Element[] {
  const out: Element[] = []
  // OOXML schema order inside <w:pPr>: ... numPr, pBdr, shd, spacing, ind, jc,
  // outlineLvl, ..., rPr (last). We only emit the ones we manage; preserving
  // children we don't touch is the engine's job for in-place mutation.
  if (
    fmt.spaceBefore !== undefined ||
    fmt.spaceAfter !== undefined ||
    fmt.lineSpacing !== undefined
  ) {
    const spacing = ownerDoc.createElementNS(w, "w:spacing")
    if (fmt.spaceBefore !== undefined) {
      spacing.setAttributeNS(w, "w:before", String(toTwips(fmt.spaceBefore, "spaceBefore")))
    }
    if (fmt.spaceAfter !== undefined) {
      spacing.setAttributeNS(w, "w:after", String(toTwips(fmt.spaceAfter, "spaceAfter")))
    }
    if (fmt.lineSpacing !== undefined) {
      const ls = parseLineSpacing(fmt.lineSpacing as LineSpacingInput, "lineSpacing")
      spacing.setAttributeNS(w, "w:line", String(ls.value))
      spacing.setAttributeNS(w, "w:lineRule", ls.mode)
    }
    out.push(spacing)
  }
  if (
    fmt.firstLineIndent != null ||
    fmt.hangingIndent != null ||
    fmt.indentLeft != null ||
    fmt.indentRight != null
  ) {
    const ind = ownerDoc.createElementNS(w, "w:ind")
    const fli = parseIndent(fmt.firstLineIndent ?? null)
    if (fli && fli.value !== 0) {
      const attr = fli.kind === "char" ? "w:firstLineChars" : "w:firstLine"
      ind.setAttributeNS(w, attr, String(fli.value))
    }
    const hi = parseIndent(fmt.hangingIndent ?? null)
    if (hi && hi.value !== 0) {
      const attr = hi.kind === "char" ? "w:hangingChars" : "w:hanging"
      ind.setAttributeNS(w, attr, String(hi.value))
    }
    const il = parseIndent(fmt.indentLeft ?? null)
    if (il) {
      const attr = il.kind === "char" ? "w:leftChars" : "w:left"
      ind.setAttributeNS(w, attr, String(il.value))
    }
    const ir = parseIndent(fmt.indentRight ?? null)
    if (ir) {
      const attr = ir.kind === "char" ? "w:rightChars" : "w:right"
      ind.setAttributeNS(w, attr, String(ir.value))
    }
    out.push(ind)
  }
  if (fmt.alignment) {
    const jc = ownerDoc.createElementNS(w, "w:jc")
    jc.setAttributeNS(w, "w:val", fmt.alignment)
    out.push(jc)
  }
  if (fmt.outlineLevel !== undefined) {
    const ol = ownerDoc.createElementNS(w, "w:outlineLvl")
    ol.setAttributeNS(w, "w:val", String(fmt.outlineLevel))
    out.push(ol)
  }
  return out
}

/** pPr children that buildPPrChildren manages (excluding pStyle and numPr,
 * which are addressed separately by the engine — pStyle through styleId
 * assignment, numPr through numbering binding). */
export const PPR_MANAGED_LOCAL_NAMES: ReadonlySet<string> = new Set([
  "spacing",
  "ind",
  "jc",
  "outlineLvl",
])

/* ------------- richtext → runs ------------- */

function emitRun(text: string, format: RunFormat | undefined, ownerDoc: Document): Element {
  const r = ownerDoc.createElementNS(w, "w:r")
  if (format) {
    const rPr = ownerDoc.createElementNS(w, "w:rPr")
    for (const c of buildRPrChildren(format, ownerDoc)) rPr.appendChild(c)
    if (rPr.childNodes.length > 0) r.appendChild(rPr)
  }
  const t = ownerDoc.createElementNS(w, "w:t")
  // Word strips leading/trailing whitespace inside <w:t> unless xml:space is
  // preserved. Set unconditionally — cheap, avoids "trailing space before
  // newline got eaten" surprises.
  t.setAttribute("xml:space", "preserve")
  t.appendChild(ownerDoc.createTextNode(text))
  r.appendChild(t)
  return r
}

export function emitRichText(
  rt: RichText,
  ownerDoc: Document,
  ctx: EmitContext,
  defaultFormat?: RunFormat,
): Element[] {
  if (typeof rt === "string") {
    return rt.length === 0 ? [] : [emitRun(rt, defaultFormat, ownerDoc)]
  }
  const out: Element[] = []
  for (const piece of rt) {
    if ("refTo" in piece) {
      if (!ctx.emitRef) {
        throw new Error("InlineRef encountered but ctx.emitRef was not provided by the engine")
      }
      const fmt = piece.format ?? defaultFormat
      for (const r of ctx.emitRef(piece, ownerDoc, fmt)) out.push(r)
      continue
    }
    if ("math" in piece) {
      out.push(emitInlineEquation(piece.math, ownerDoc))
      continue
    }
    if ("link" in piece) {
      if (!ctx.emitHyperlink) {
        throw new Error(
          "HyperlinkNode encountered but ctx.emitHyperlink was not provided by the engine",
        )
      }
      const fmt = piece.format ?? defaultFormat
      out.push(ctx.emitHyperlink(piece.link, piece.text, fmt, ownerDoc))
      continue
    }
    if ("field" in piece) {
      const fmt = piece.format ?? defaultFormat
      for (const r of emitInlineField(ownerDoc, piece.field, fmt)) out.push(r)
      continue
    }
    if ("styleRef" in piece) {
      const fmt = piece.format ?? defaultFormat
      for (const r of emitInlineStyleRef(ownerDoc, piece.styleRef, piece.numberOnly ?? false, fmt)) {
        out.push(r)
      }
      continue
    }
    out.push(emitRun(piece.text, piece.format ?? defaultFormat, ownerDoc))
  }
  return out
}

/* ------------- block emitters ------------- */

/**
 * Optional callback that materialises an image asset (zip media/, rels,
 * Content_Types) and returns the inline <w:drawing> element. Wired by
 * edit-engine when image assets are in scope. Decoupled from this module
 * so fragment-emit stays a pure XML producer.
 */
export type ImageEmitter = (
  src: string,
  width: Length,
  height: Length,
  alt: string | undefined,
  ownerDoc: Document,
) => Element

export interface EmitContext {
  emitImage?: ImageEmitter
  /** Provided when edits[] contains InlineRef nodes. Engine builds this
   * over a bookmark allocator + locator resolver + pending-backfill queue.
   * Absent ctx.emitRef + an InlineRef in input = engine error at emit. */
  emitRef?: RefEmitter
  /** Provided when edits[] contains HyperlinkNode inline nodes. Returns the
   * complete `<w:hyperlink>` element ready to insert as a paragraph child.
   * Decoupled so fragment-emit stays free of rels-registry / character-style
   * concerns. Absent ctx.emitHyperlink + a HyperlinkNode in input = engine
   * error at emit. */
  emitHyperlink?: (link: string, text: string, format: RunFormat | undefined, ownerDoc: Document) => Element
  /** Called when a `ParagraphBlock.anchor` is set — registers the named
   * bookmark on the just-emitted paragraph Element. Absent ctx.adoptAnchor
   * + an anchor in input = engine error at emit. */
  adoptAnchor?: (name: string, pEl: Element) => void
  /** Usable content width (LaTeX `\textwidth`) of the section the current
   * op is targeting, in twips. Consumed by `emitTableBlock` to seed
   * autofit gridCol widths. Populated per-op by the engine (different ops
   * can target different sections); absent for ops targeting a `<w:tc>`
   * container or when the document's sectPr lacks pgSz/pgMar — emitters
   * fall back to a conservative constant in that case. */
  usableWidthTwips?: number
  /** Caption pipeline callbacks (spec §4.5-§4.7). Grouped as one
   * sub-object so all five callbacks travel together — caption blocks
   * dispatch behind a single presence check rather than five. Absent
   * `captions` → caption blocks throw at emit. */
  captions?: CaptionEmitCallbacks
}

export interface CaptionEmitCallbacks {
  /** Resolve an identifier (CaptionBlock.captionId / EquationBlock.captionId)
   * to its resolved config. Engine populates from apply config's
   * `captions` table at apply start. Returns undefined when the
   * identifier isn't declared. */
  resolve: (identifier: string) => ResolvedCaptionConfig | undefined
  /** Reserve a bookmark id+name. Caption emit writes bookmarkStart/End
   * inline around number runs; `bindBookmark` records the paragraph
   * binding so REF \h can resolve cross-references. */
  allocateBookmark: (name: string) => BookmarkRange
  /** Post-emit binding for `allocateBookmark`. */
  bindBookmark: (name: string, pEl: Element) => void
  /** Register a caption fill record so the counter sim can compute
   * rendered values post-emit. */
  registerFill: (fill: PendingCaptionFill) => void
  /** Register a caption counter reset marker. */
  registerReset: (reset: PendingCaptionReset) => void
}

function ensurePPr(p: Element, ownerDoc: Document): Element {
  for (const c of Array.from(p.childNodes)) {
    if (c.nodeType === 1) {
      const el = c as Element
      if (el.namespaceURI === w && el.localName === "pPr") return el
    }
  }
  const pPr = ownerDoc.createElementNS(w, "w:pPr")
  p.insertBefore(pPr, p.firstChild)
  return pPr
}

function emitParagraphBlock(
  block: Extract<Block, { type: "paragraph" }>,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  const p = ownerDoc.createElementNS(w, "w:p")
  const needsPPr =
    block.styleId !== undefined || block.paraFormat !== undefined || block.numbering !== undefined
  if (needsPPr) {
    const pPr = ensurePPr(p, ownerDoc)
    if (block.styleId) {
      const ps = ownerDoc.createElementNS(w, "w:pStyle")
      ps.setAttributeNS(w, "w:val", block.styleId)
      pPr.appendChild(ps)
    }
    if (block.numbering) {
      const numPr = ownerDoc.createElementNS(w, "w:numPr")
      const ilvl = ownerDoc.createElementNS(w, "w:ilvl")
      ilvl.setAttributeNS(w, "w:val", String(block.numbering.level))
      numPr.appendChild(ilvl)
      const numId = ownerDoc.createElementNS(w, "w:numId")
      numId.setAttributeNS(w, "w:val", block.numbering.numId)
      numPr.appendChild(numId)
      pPr.appendChild(numPr)
    }
    if (block.paraFormat) {
      for (const c of buildPPrChildren(block.paraFormat, ownerDoc)) pPr.appendChild(c)
    }
  }
  for (const r of emitRichText(block.text, ownerDoc, ctx, block.runFormat)) p.appendChild(r)
  if (block.anchor) {
    if (!ctx.adoptAnchor) {
      throw new Error(
        "ParagraphBlock.anchor encountered but ctx.adoptAnchor was not provided by the engine",
      )
    }
    ctx.adoptAnchor(block.anchor, p)
  }
  return p
}

function emitImageBlock(
  block: Extract<Block, { type: "image" }>,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  if (!ctx.emitImage) {
    throw new Error("image block encountered but no ImageEmitter wired (image asset path inactive)")
  }
  const drawing = ctx.emitImage(block.src, block.width, block.height, block.alt, ownerDoc)
  const p = ownerDoc.createElementNS(w, "w:p")
  if (block.styleId !== undefined || block.paraFormat !== undefined) {
    const pPr = ensurePPr(p, ownerDoc)
    if (block.styleId) {
      const ps = ownerDoc.createElementNS(w, "w:pStyle")
      ps.setAttributeNS(w, "w:val", block.styleId)
      pPr.appendChild(ps)
    }
    if (block.paraFormat) {
      for (const c of buildPPrChildren(block.paraFormat, ownerDoc)) pPr.appendChild(c)
    }
  }
  const r = ownerDoc.createElementNS(w, "w:r")
  r.appendChild(drawing)
  p.appendChild(r)
  return p
}

function emitPageBreakBlock(ownerDoc: Document): Element {
  const p = ownerDoc.createElementNS(w, "w:p")
  const r = ownerDoc.createElementNS(w, "w:r")
  const br = ownerDoc.createElementNS(w, "w:br")
  br.setAttributeNS(w, "w:type", "page")
  r.appendChild(br)
  p.appendChild(r)
  return p
}

function emitHorizontalRuleBlock(ownerDoc: Document): Element {
  // Convention: empty paragraph carrying a bottom border. Word renders this
  // as a horizontal line; keeps the structure as a single <w:p> like every
  // other block.
  const p = ownerDoc.createElementNS(w, "w:p")
  const pPr = ensurePPr(p, ownerDoc)
  const pBdr = ownerDoc.createElementNS(w, "w:pBdr")
  const bottom = ownerDoc.createElementNS(w, "w:bottom")
  bottom.setAttributeNS(w, "w:val", "single")
  bottom.setAttributeNS(w, "w:sz", "6")
  bottom.setAttributeNS(w, "w:space", "1")
  bottom.setAttributeNS(w, "w:color", "auto")
  pBdr.appendChild(bottom)
  pPr.appendChild(pBdr)
  return p
}

/* ------------- public emit ------------- */

export function emitBlock(block: Block, ownerDoc: Document, ctx: EmitContext): Element {
  switch (block.type) {
    case "paragraph":
      return emitParagraphBlock(block, ownerDoc, ctx)
    case "image":
      return emitImageBlock(block, ownerDoc, ctx)
    case "page-break":
      return emitPageBreakBlock(ownerDoc)
    case "horizontal-rule":
      return emitHorizontalRuleBlock(ownerDoc)
    case "table":
      return emitTableBlock(block, ownerDoc, ctx)
    case "equation":
      return emitEquationBlock(block, ownerDoc, ctx)
    case "caption":
      return dispatchCaption(block, ownerDoc, ctx)
    case "caption-counter-reset":
      return dispatchCaptionReset(block, ownerDoc, ctx)
    default:
      return assertNever(block)
  }
}

function dispatchCaption(
  block: Extract<Block, { type: "caption" }>,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  if (!ctx.captions) {
    throw new Error(
      "CaptionBlock: ctx.captions callbacks not provided by the engine. " +
        "Caption blocks require the captions table to be declared in the apply config.",
    )
  }
  const config = ctx.captions.resolve(block.captionId)
  if (!config) {
    throw new Error(
      `CaptionBlock: captionId "${block.captionId}" is not declared in captions table.`,
    )
  }
  const bookmark =
    block.anchor !== undefined ? ctx.captions.allocateBookmark(block.anchor) : undefined
  const { paragraph, fill } = emitCaptionBlock(ownerDoc, {
    captionConfig: config,
    text: block.text,
    bookmark,
  })
  if (block.anchor !== undefined) {
    ctx.captions.bindBookmark(block.anchor, paragraph)
  }
  ctx.captions.registerFill(fill)
  return paragraph
}

function dispatchCaptionReset(
  block: Extract<Block, { type: "caption-counter-reset" }>,
  ownerDoc: Document,
  ctx: EmitContext,
): Element {
  if (!ctx.captions) {
    throw new Error("CaptionCounterReset: ctx.captions callbacks not provided by the engine.")
  }
  const config = ctx.captions.resolve(block.captionId)
  if (!config) {
    throw new Error(
      `CaptionCounterReset: captionId "${block.captionId}" is not declared in captions table.`,
    )
  }
  const { paragraph, reset } = emitCaptionReset(ownerDoc, {
    identifier: block.captionId,
    newValue: block.newValue ?? 1,
  })
  ctx.captions.registerReset(reset)
  return paragraph
}

export function emitFragment(fragment: Fragment, ownerDoc: Document, ctx: EmitContext): Element[] {
  return fragment.map((b) => emitBlock(b, ownerDoc, ctx))
}
