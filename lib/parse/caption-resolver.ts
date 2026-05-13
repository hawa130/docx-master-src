/**
 * Resolve the raw `captions` config (agent-facing) into per-identifier
 * `ResolvedCaptionConfig` records consumed by caption-emit + counter sim.
 *
 * Resolution work:
 *   - Apply defaults: prefix/suffix → "", format → "arabic",
 *     chapterPrefix → [], chapterSeparator → ".", bodySeparator → " ",
 *     subCounter format → "alphabetic", subCounter prefix/suffix → ""
 *   - For each chapterPrefix styleId, look up styleName from styles.xml's
 *     `<w:name w:val="..."/>` and outlineLvl from `<w:pPr><w:outlineLvl/>`
 *   - Derive SEQ \s switch's outline level from the LAST chapterPrefix
 *     entry (deepest chapter level controls counter restart)
 *
 * Throws on:
 *   - chapterPrefix references unknown styleId
 *   - chapterPrefix references a style without `<w:name>` (engine bug
 *     for built-in styles; agent error for hand-written styles)
 *
 * Does NOT throw on:
 *   - chapterPrefix style without outlineLvl binding (Word will render
 *     "0"; emit pre-scan can warn separately)
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"
import type { CaptionsConfig } from "@lib/config/config-types.ts"
import type { ResolvedCaptionConfig } from "@lib/edit/caption-counter.ts"

const w = NS.w

export interface ResolvedCaptions {
  byIdentifier: Map<string, ResolvedCaptionConfig>
  /** styleId → styleName lookup used by the runtime caption pipeline
   * (standardize re-emit needs to match paragraph pStyle against
   * caption styleIds; outline simulator output annotation needs the
   * styleName for STYLEREF resolution). */
  styleIdToName: Map<string, string>
}

export function resolveCaptions(
  raw: CaptionsConfig | undefined,
  stylesDoc: Document | null,
): ResolvedCaptions {
  const byIdentifier = new Map<string, ResolvedCaptionConfig>()
  const styleIdToName = new Map<string, string>()
  if (!raw) return { byIdentifier, styleIdToName }

  // Build styleId → { name, outlineLevel } from styles.xml once.
  const styleIndex = indexStyles(stylesDoc)

  for (const [identifier, entry] of Object.entries(raw)) {
    const chapterPrefix: Array<{ styleName: string; outlineLevel: number }> = []
    for (const styleId of entry.chapterPrefix ?? []) {
      const info = styleIndex.get(styleId)
      if (!info) {
        throw new Error(
          `captions["${identifier}"].chapterPrefix references unknown styleId "${styleId}". ` +
            `Declare the style in styles[] or fix the reference.`,
        )
      }
      chapterPrefix.push({
        styleName: info.name,
        outlineLevel: info.outlineLevel ?? 0,
      })
      styleIdToName.set(styleId, info.name)
    }
    const restartAtOutlineLevel =
      chapterPrefix.length > 0
        ? chapterPrefix[chapterPrefix.length - 1]!.outlineLevel || undefined
        : undefined

    // Note paragraph styleId mapping too — standardize re-emit needs it.
    const paraInfo = styleIndex.get(entry.styleId)
    if (paraInfo) styleIdToName.set(entry.styleId, paraInfo.name)

    const subCounter = entry.subCounter
      ? {
          format: entry.subCounter.format ?? "alphabetic",
          prefix: entry.subCounter.prefix ?? "",
          suffix: entry.subCounter.suffix ?? "",
        }
      : undefined

    byIdentifier.set(identifier, {
      identifier,
      prefix: entry.prefix ?? "",
      suffix: entry.suffix ?? "",
      format: entry.format ?? "arabic",
      chapterPrefix,
      chapterSeparator: entry.chapterSeparator ?? ".",
      bodySeparator: entry.bodySeparator ?? " ",
      paragraphStyleId: entry.styleId,
      restartAtOutlineLevel,
      subCounter,
    })
  }

  return { byIdentifier, styleIdToName }
}

interface StyleInfo {
  name: string
  /** 1-indexed outline level (matches SEQ `\s` and Word UI convention).
   * `<w:outlineLvl w:val="0"/>` in styles.xml maps to outlineLevel 1. */
  outlineLevel: number | undefined
}

function indexStyles(stylesDoc: Document | null): Map<string, StyleInfo> {
  const out = new Map<string, StyleInfo>()
  if (!stylesDoc) return out
  const root = stylesDoc.documentElement
  if (!root) return out
  for (const styleEl of getChildrenNS(root, w, "style")) {
    const id = wAttr(styleEl, "styleId")
    if (!id) continue
    const nameEl = firstChildNS(styleEl, w, "name")
    const name = nameEl ? (wAttr(nameEl, "val") ?? id) : id
    let outlineLevel: number | undefined
    const pPr = firstChildNS(styleEl, w, "pPr")
    if (pPr) {
      const lvlEl = firstChildNS(pPr, w, "outlineLvl")
      if (lvlEl) {
        const v = wAttr(lvlEl, "val")
        const n = v ? parseInt(v, 10) : NaN
        // OOXML stores 0-indexed; SEQ \s and our internal convention
        // are 1-indexed.
        if (Number.isFinite(n) && n >= 0 && n <= 8) outlineLevel = n + 1
      }
    }
    out.set(id, { name, outlineLevel })
  }
  return out
}
