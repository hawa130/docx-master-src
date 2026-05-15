/**
 * One docx part's `_rels/<part>.rels` state.
 *
 * Each OOXML part (document.xml, header1.xml, footer1.xml, …) has its own
 * rels file with a per-part rId namespace. This class owns one such file —
 * parsing, allocating fresh rIds, appending Relationship entries, and
 * flushing the modified text back into the writer's replacement map.
 *
 * Mutation is string-level: the rels XML is small enough that DOM-mutate
 * + serialize would add machinery for no benefit. `appendRel` rewrites
 * `</Relationships>` with the new entry preceding it.
 */

import { type DocxReader, parseXml } from "@lib/xml/reader.ts"
import type { WritableArchive } from "@lib/xml/writable-archive.ts"

interface RelEntry {
  id: string
  target: string
  type: string
}

const DEFAULT_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`

export class PartRels {
  private relsText: string
  private rels: RelEntry[] = []
  private nextRId = 1
  private dirty = false
  /** Highest seen `media/imageN.<ext>` index + 1 — used by `nextImageNum()`
   *  helper so the asset-registry layer can dedup with existing media. */
  private highestImageNum = 1

  constructor(initialRelsXml: string) {
    this.relsText = initialRelsXml
    this.parseRels()
  }

  /** Open from a docx reader. `relsPath` is the in-archive path
   *  (e.g. `word/_rels/document.xml.rels`). Falls back to an empty
   *  Relationships skeleton when the part has no existing rels file. */
  static async open(reader: DocxReader, relsPath: string): Promise<PartRels> {
    const text = (await reader.readText(relsPath)) ?? DEFAULT_RELS_XML
    return new PartRels(text)
  }

  /** Append a Relationship, returning the freshly allocated rId. `mode`
   *  is "External" for hyperlinks / external links; omit for internal
   *  part-to-part references (images, headers, footers, …). Targets are
   *  XML-attr-escaped for agent-supplied values (hrefs); the caller is
   *  responsible for pre-encoding URLs if they want percent encoding. */
  appendRel(type: string, target: string, mode?: "External"): { rId: string } {
    const rId = this.allocateRId()
    const modeAttr = mode ? ` TargetMode="${mode}"` : ""
    const escapedTarget = escapeXmlAttr(target)
    const insert = `<Relationship Id="${rId}" Type="${type}" Target="${escapedTarget}"${modeAttr}/>`
    this.relsText = this.relsText.replace("</Relationships>", `${insert}</Relationships>`)
    this.rels.push({ id: rId, target, type })
    this.dirty = true
    return { rId }
  }

  /** Returns whether any rels were appended since construction. Callers
   *  gate flush on this so untouched parts don't rewrite their rels file. */
  isDirty(): boolean {
    return this.dirty
  }

  /** Check whether a Relationship with the given Target string already
   *  exists. Used by callers that want to be idempotent about adding a
   *  named-part rel (e.g. the numbering.xml registration, which apply
   *  re-runs on a doc that already has the rel from a prior apply). */
  hasRelTo(target: string): boolean {
    return this.rels.some((r) => r.target === target)
  }

  /** Stage the modified rels text at `path` in the writer's replacement
   *  map. No-op when nothing was appended. The body's rels path
   *  (`word/_rels/document.xml.rels`) is one of the forbidden direct-
   *  write targets on `WritableArchive`; uses the `_setFromAccumulator`
   *  escape hatch since this IS the legitimate accumulator path.
   *  Per-part HF rels paths (header1.xml.rels etc.) aren't forbidden
   *  and could go through `.set` too, but routing both through one
   *  method keeps the call sites uniform. */
  flushTo(replacements: WritableArchive, path: string): void {
    if (!this.dirty) return
    replacements.setFromAccumulator(path, this.relsText)
  }

  /** Returns the next free `imageN` suffix for media path allocation.
   *  Used by the asset registry to avoid clashing with images already
   *  registered in this part's rels. */
  nextImageNum(): number {
    return this.highestImageNum
  }

  /** Bump the image-number cursor past `n`. Called by the asset registry
   *  after it assigns `media/imageN.<ext>` for a fresh upload, so the
   *  next allocation skips ahead. */
  advanceImageNum(n: number): void {
    if (n + 1 > this.highestImageNum) this.highestImageNum = n + 1
  }

  /* ------------- internals ------------- */

  private parseRels(): void {
    const doc = parseXml(this.relsText)
    const root = doc.documentElement
    if (!root) return
    let maxNum = 0
    for (let i = 0; i < root.childNodes.length; i++) {
      const c = root.childNodes[i]
      if (!c || c.nodeType !== 1) continue
      const el = c as Element
      if (el.localName !== "Relationship") continue
      const id = el.getAttribute("Id") ?? ""
      const target = el.getAttribute("Target") ?? ""
      const type = el.getAttribute("Type") ?? ""
      if (id) {
        this.rels.push({ id, target, type })
        const m = id.match(/^rId(\d+)$/)
        if (m) maxNum = Math.max(maxNum, parseInt(m[1]!, 10))
      }
      const mt = target.match(/^(?:\.\.\/)?media\/image(\d+)\./)
      if (mt) {
        const n = parseInt(mt[1]!, 10) + 1
        if (n > this.highestImageNum) this.highestImageNum = n
      }
    }
    this.nextRId = maxNum + 1
  }

  private allocateRId(): string {
    return `rId${this.nextRId++}`
  }
}

/** Minimal XML-attribute escape — Relationship targets like hyperlink hrefs
 *  may contain `&` (query strings), `"`, `<`, `>`. Encode just enough to
 *  keep the rels XML valid. */
function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
}
