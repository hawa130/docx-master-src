/**
 * Shared zod validators used across both `ApplyConfigSchema` and
 * `EditConfigSchema`. Both files repeated these identically; consolidating
 * here so future tightening (e.g. trim semantics for NonEmptyString)
 * lands in one place.
 */

import * as z from "zod/mini"
import { parseIndent, parseLengthPt, parseLineSpacing } from "@lib/shared/units.ts"

/** Non-empty string. */
export const NonEmptyString = z.string().check(z.minLength(1))

/** Length value:
 *   - `number` — pt (bare).
 *   - `"Npt" / "Ncm" / "Nmm" / "Nin"` — explicit unit.
 *  Validated eagerly at parse time so unit typos surface at config-read,
 *  not at the emit site.
 */
export const LengthValue = z
  .union([z.number(), z.string()])
  .check(
    z.refine(
      (v) => {
        try {
          parseLengthPt(v)
          return true
        } catch {
          return false
        }
      },
      {
        error: (issue) => {
          try {
            parseLengthPt(issue.input as number | string)
            return "length parse failed"
          } catch (e) {
            return (e as Error).message
          }
        },
      },
    ),
  )

/** Paragraph indent value: a Length plus the special `"Nchar"` form (round-trips
 *  Word's `w:firstLineChars` / `w:hangingChars`, auto-scales with font size).
 *  Explicit `null` opts out of any indent attribute. */
export const IndentValue = z
  .union([z.number(), z.string(), z.null()])
  .check(
    z.refine(
      (v) => {
        try {
          parseIndent(v)
          return true
        } catch {
          return false
        }
      },
      {
        error: (issue) => {
          try {
            parseIndent(issue.input as number | string | null)
            return "indent parse failed"
          } catch (e) {
            return (e as Error).message
          }
        },
      },
    ),
  )

/** Line spacing value carries the OOXML `lineRule` choice via type:
 *   - `number` — multiplier (auto), e.g. `1.5`
 *   - `"Npt" / "Ncm" / ...` — exact line height
 *   - `{ atLeast: <Length> }` — at-least line height
 *  No magnitude heuristic — bare numbers are always multipliers. */
export const LineSpacingValue = z
  .union([z.number(), z.string(), z.strictObject({ atLeast: z.union([z.number(), z.string()]) })])
  .check(
    z.refine(
      (v) => {
        try {
          parseLineSpacing(v)
          return true
        } catch {
          return false
        }
      },
      {
        error: (issue) => {
          try {
            parseLineSpacing(issue.input as number | string | { atLeast: number | string })
            return "lineSpacing parse failed"
          } catch (e) {
            return (e as Error).message
          }
        },
      },
    ),
  )
