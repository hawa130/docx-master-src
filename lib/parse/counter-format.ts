/**
 * Shared counter-rendering primitives.
 *
 * Two distinct counter renderers live in this repo:
 *   - `lib/apply/numbering-counter.ts` (Word `numFmt` vocabulary —
 *     `chineseCounting` uses `〇`, `ideographTraditional` = heavenly
 *     stems, etc.)
 *   - `lib/edit/caption-counter.ts` (project `SeqFormat` vocabulary —
 *     `chinese` uses `零`, `chinese-formal` uses 壹贰叁, etc.)
 *
 * Their format dispatch tables don't overlap (different vocabularies),
 * so they stay separate. But the underlying numeric → glyph primitives
 * (roman, spreadsheet-style alpha) are identical and live here.
 *
 * Chinese-digit tables intentionally do NOT live here — the two
 * renderers spell zero differently (`〇` for Word's numFmt vs `零` for
 * the project's caption format) and that divergence is by design.
 */

/** Upper-case "I, II, III, IV, V, ...". Callers wanting lower-case
 * lowercase the result. Returns "" for n ≤ 0. */
export function toRoman(n: number): string {
  if (n <= 0) return ""
  const pairs: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ]
  let out = ""
  let m = n
  for (const [v, r] of pairs) {
    while (m >= v) {
      out += r
      m -= v
    }
  }
  return out
}

/** Spreadsheet-column style: A..Z, AA..ZZ, AAA... `upper` picks A-Z vs
 * a-z. Returns "" for n ≤ 0. */
export function toAlphaCounter(n: number, upper: boolean): string {
  if (n <= 0) return ""
  const base = upper ? 65 : 97
  let result = ""
  let m = n
  while (m > 0) {
    m--
    result = String.fromCharCode(base + (m % 26)) + result
    m = Math.floor(m / 26)
  }
  return result
}
