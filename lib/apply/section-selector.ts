/**
 * Section selectors — `"N"` or `"N-M"` (1-based, inclusive).
 *
 * Shared by `pageSetup.sections` and `headerFooter.sections`. Schema-level
 * validation uses `SECTION_SELECTOR_KEY_RE`; runtime expansion to concrete
 * section indices uses `expandSectionSelector`, which throws with a
 * caller-supplied `fieldPath` so the error names the config root the user
 * wrote (e.g. `pageSetup.sections` vs `headerFooter.sections`).
 */

export const SECTION_SELECTOR_KEY_RE = /^\d+(?:-\d+)?$/

export function expandSectionSelector(
  key: string,
  sectionCount: number,
  fieldPath: string,
): number[] {
  const range = key.match(/^(\d+)-(\d+)$/)
  if (range) {
    const lo = parseInt(range[1]!, 10)
    const hi = parseInt(range[2]!, 10)
    if (lo < 1 || hi > sectionCount || lo > hi) {
      throw new Error(
        `${fieldPath}: range "${key}" out of bounds. Document has ${sectionCount} section(s); valid 1..${sectionCount}.`,
      )
    }
    return Array.from({ length: hi - lo + 1 }, (_, i) => lo + i)
  }
  const n = parseInt(key, 10)
  if (n < 1 || n > sectionCount) {
    throw new Error(
      `${fieldPath}: section ${n} out of bounds. Document has ${sectionCount} section(s); valid 1..${sectionCount}.`,
    )
  }
  return [n]
}
