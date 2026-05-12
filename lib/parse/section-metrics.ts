/**
 * Derived metrics over parsed section info.
 *
 * Both helpers are pure functions over `SectionInfo` — no XML access, no
 * mutation. They live here (rather than in `document-parser.ts`) so the
 * parser file stays focused on read-side parsing and these can be reused
 * without pulling in the parser.
 */

import type { SectionInfo } from "@lib/parse/types.ts"

/** Content-area width in twips: `pgSz.w − pgMar.left − pgMar.right`.
 *
 * The result is the value LaTeX would call `\textwidth` — the horizontal
 * box block-level content can occupy in this section before colliding
 * with the page margin. Returns 0 when `pageSize.width` is unavailable
 * (rare hand-built fixtures without `<w:pgSz>`).
 */
export function sectionUsableWidthTwips(section: SectionInfo): number {
  return Math.max(0, section.pageSize.width - section.margins.left - section.margins.right)
}

/** Locate the section whose `paraRange` contains the given 1-based
 * paragraph index. Returns `null` when no section matches (index out of
 * any range, e.g. against a fresh paragraph index past the original doc
 * length) or when `sections` is empty. Callers handling "insert at end"
 * should pass the last existing paragraph index, or fall back to
 * `sections[sections.length - 1]` directly when paragraphs is empty.
 */
export function sectionForParagraph(
  sections: SectionInfo[],
  paragraphIndex: number,
): SectionInfo | null {
  for (const s of sections) {
    if (paragraphIndex >= s.paraRange[0] && paragraphIndex <= s.paraRange[1]) return s
  }
  return null
}
