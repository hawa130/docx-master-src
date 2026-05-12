/**
 * Shared zod validators used across both `ApplyConfigSchema` and
 * `EditConfigSchema`. Both files repeated these identically; consolidating
 * here so future tightening (e.g. trim semantics for NonEmptyString)
 * lands in one place.
 */

import * as z from "zod/mini"

/** Non-empty string. */
export const NonEmptyString = z.string().check(z.minLength(1))

/** Paragraph indent value:
 *   - `"Nchar"` — character-units (e.g. `"2char"`), round-trips Word's
 *     `w:firstLineChars` / `w:hangingChars` and auto-scales with font size.
 *   - `"Npt"` — fixed points (e.g. `"12pt"`).
 *   - `number` — fixed points as a bare value.
 *   - `null` — explicit zero / no indent.
 */
export const IndentValue = z.union([z.string(), z.number(), z.null()])
