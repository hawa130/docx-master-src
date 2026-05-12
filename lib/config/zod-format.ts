/**
 * Shared zod error formatter.
 *
 * Both `ApplyConfigSchema` and `EditConfigSchema` parsers translate
 * `z.core.$ZodError` into multi-line, agent-readable text. The path
 * walking, issue rendering, and union-variant drilling are identical;
 * only the domain-specific hints differ. This module owns the generic
 * machinery; each parser supplies its own `HintFn` for domain hints.
 *
 * `flattenUnionIssues` is the load-bearing piece: zod surfaces a deeply
 * nested validation failure as `invalid_union` at the union's path with
 * the real per-variant issues hidden in `issue.errors`. Without
 * flattening, an agent sees `edits[0]: Invalid input` instead of
 * `edits[0].content[1]: unknown keys "rows", "alignment"`.
 */

import type * as z from "zod/mini"

export type HintFn = (
  issue: z.core.$ZodIssue,
  pathStr: string,
  raw: unknown,
) => string | null

/** Convert a zod issue.path (PropertyKey[]) to dotted/bracketed JS form,
 * e.g. `["numbering","levels",0,"styleId"]` → `numbering.levels[0].styleId`. */
export function formatPath(path: readonly PropertyKey[]): string {
  let out = ""
  for (const seg of path) {
    if (typeof seg === "number") out += `[${seg}]`
    else out += out ? `.${String(seg)}` : String(seg)
  }
  return out
}

/** Walk `path` into `raw` and return the value at that location, or undefined
 * if any segment is missing. Used so error messages can echo the actual
 * value the agent supplied — `issue.input` is unreliable across issue codes
 * in zod mini, the raw config is always the ground truth. */
export function valueAtPath(raw: unknown, path: readonly PropertyKey[]): unknown {
  let cur: unknown = raw
  for (const seg of path) {
    if (cur == null || typeof cur !== "object") return undefined
    cur = (cur as Record<PropertyKey, unknown>)[seg]
  }
  return cur
}

/** Format a single zod issue into an agent-readable line. Falls back to the
 * issue's default message only when no specific code matches; for the common
 * codes we render concrete details (unknown keys, expected types, allowed
 * enum values) instead of the bare "Invalid input" zod emits at low cost. */
export function formatIssue(issue: z.core.$ZodIssue, raw: unknown): string {
  const got = valueAtPath(raw, issue.path)
  const gotStr = got === undefined ? "(missing)" : JSON.stringify(got)
  switch (issue.code) {
    case "unrecognized_keys": {
      const keys = (issue as { keys?: string[] }).keys ?? []
      const noun = keys.length === 1 ? "key" : "keys"
      return `unknown ${noun} ${keys.map((k) => `"${k}"`).join(", ")}`
    }
    case "invalid_type": {
      const expected = (issue as { expected?: string }).expected ?? "value"
      return got === undefined
        ? `missing required field (expected ${expected})`
        : `expected ${expected}, got ${gotStr}`
    }
    case "invalid_value": {
      const allowed = (issue as { values?: readonly unknown[] }).values ?? []
      return `invalid value ${gotStr}. Allowed: [${allowed.map((v) => JSON.stringify(v)).join(", ")}]`
    }
    case "too_small": {
      const min = (issue as { minimum?: unknown }).minimum
      const origin = (issue as { origin?: string }).origin
      if (origin === "string") return `must be a non-empty string (got ${gotStr})`
      if (origin === "array") return `must contain at least ${String(min)} item(s)`
      return `value too small (minimum ${String(min)})`
    }
    case "too_big": {
      const max = (issue as { maximum?: unknown }).maximum
      return `value too large (maximum ${String(max)})`
    }
    case "invalid_format":
      return `format check failed (${(issue as { format?: string }).format ?? "unknown"})`
    default:
      return issue.message
  }
}

/** When zod reports `invalid_union`, the real issues are nested per-variant.
 * Pick the variant that matched the discriminator (its issues are not just
 * `invalid_value` on a `type` / `op` field) and surface those. Recurse if
 * the matched variant itself fails another union. */
export function flattenUnionIssues(issue: z.core.$ZodIssue): z.core.$ZodIssue[] {
  if (issue.code !== "invalid_union") return [issue]
  const variants = (issue as { errors?: z.core.$ZodIssue[][] }).errors ?? []
  // Score each variant by "real issue count" (excluding type-discriminator
  // mismatches). Prefer the variant with the lowest non-discriminator score
  // and at least one such issue — that's the variant the agent intended.
  let best: { issues: z.core.$ZodIssue[]; score: number } | null = null
  for (const variant of variants) {
    const realIssues = variant.filter(
      (sub) =>
        !(
          sub.code === "invalid_value" &&
          sub.path.length === 1 &&
          (sub.path[0] === "type" || sub.path[0] === "op")
        ),
    )
    if (realIssues.length === 0) continue
    const score = realIssues.length
    if (!best || score < best.score) best = { issues: realIssues, score }
  }
  if (!best) {
    // All variants failed only at the discriminator — surface those raw so
    // the agent sees which `type` / `op` values are accepted.
    const first = variants.flat()[0]
    if (first) return [first]
    return [issue]
  }
  return best.issues.flatMap((sub) => {
    const merged: z.core.$ZodIssue = {
      ...sub,
      // Merge the union's path prefix with the variant-internal path.
      path: [...issue.path, ...sub.path],
    } as z.core.$ZodIssue
    return flattenUnionIssues(merged)
  })
}

/** Format a ZodError as a multi-line, agent-readable string. Each issue
 * gets a path prefix and either the supplied domain hint, a code-specific
 * render from `formatIssue`, or zod's default message as a final fallback.
 * `invalid_union` issues drill into the matched variant so the agent sees
 * the inner field path, not just `edits[0]: Invalid input`. */
export function formatZodError(
  error: z.core.$ZodError,
  raw: unknown,
  hint?: HintFn,
): string {
  const lines: string[] = []
  for (const top of error.issues) {
    for (const issue of flattenUnionIssues(top)) {
      const pathStr = formatPath(issue.path)
      const custom = hint?.(issue, pathStr, raw) ?? null
      const msg = custom ?? formatIssue(issue, raw)
      lines.push(pathStr ? `${pathStr}: ${msg}` : msg)
    }
  }
  return lines.join("\n")
}
