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
import { parseLineSpacing } from "@lib/apply/style-mutation.ts"
import { RPR_CHILD_ORDER } from "@lib/xml/xml-order.ts"

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
    const sz = ownerDoc.createElementNS(w, "w:sz")
    sz.setAttributeNS(w, "w:val", String(Math.round(fmt.size * 2)))
    out.push(sz)
    const szCs = ownerDoc.createElementNS(w, "w:szCs")
    szCs.setAttributeNS(w, "w:val", String(Math.round(fmt.size * 2)))
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

interface IndentParts {
  kind: "twip" | "char"
  value: number
}

function parseIndent(v: string | number | null): IndentParts | null {
  if (v === null) return null
  if (typeof v === "number") return { kind: "twip", value: Math.round(v * 20) }
  const m = v.trim().match(/^(-?\d+(?:\.\d+)?)\s*(char|chars|pt)?$/i)
  if (!m) return null
  const n = parseFloat(m[1]!)
  const unit = (m[2] || "").toLowerCase()
  if (unit.startsWith("char")) return { kind: "char", value: Math.round(n * 100) }
  return { kind: "twip", value: Math.round(n * 20) }
}

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
      spacing.setAttributeNS(w, "w:before", String(Math.round(fmt.spaceBefore * 20)))
    }
    if (fmt.spaceAfter !== undefined) {
      spacing.setAttributeNS(w, "w:after", String(Math.round(fmt.spaceAfter * 20)))
    }
    if (fmt.lineSpacing !== undefined) {
      const ls = parseLineSpacing(fmt.lineSpacing)
      const rule = fmt.lineRule ?? (ls.explicitPt || ls.value >= 10 ? "exact" : "auto")
      if (rule === "auto") {
        spacing.setAttributeNS(w, "w:line", String(Math.round(ls.value * 240)))
        spacing.setAttributeNS(w, "w:lineRule", "auto")
      } else {
        spacing.setAttributeNS(w, "w:line", String(Math.round(ls.value * 20)))
        spacing.setAttributeNS(w, "w:lineRule", rule)
      }
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

function emitRichText(
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
  widthPt: number,
  heightPt: number,
  alt: string | undefined,
  ownerDoc: Document,
) => Element

export interface EmitContext {
  emitImage?: ImageEmitter
  /** Provided when edits[] contains InlineRef nodes. Engine builds this
   * over a bookmark allocator + locator resolver + pending-backfill queue.
   * Absent ctx.emitRef + an InlineRef in input = engine error at emit. */
  emitRef?: RefEmitter
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
  const drawing = ctx.emitImage(block.src, block.widthPt, block.heightPt, block.alt, ownerDoc)
  const p = ownerDoc.createElementNS(w, "w:p")
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
    default:
      return assertNever(block)
  }
}

export function emitFragment(fragment: Fragment, ownerDoc: Document, ctx: EmitContext): Element[] {
  return fragment.map((b) => emitBlock(b, ownerDoc, ctx))
}
