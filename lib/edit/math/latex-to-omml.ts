/**
 * LaTeX → OMML pipeline.
 *
 * Two-stage conversion, isolated here so v2 can swap the MML→OMML half:
 *
 *   stage 1: Temml renders LaTeX → MathML XML string. MIT-licensed, browser
 *            + Node. `xml: true` keeps the renderer DOM-free. Loaded via
 *            dynamic import + fallback because tsdown corrupts the
 *            surrogate-range literals in temml's tokenizer when bundling;
 *            build-skill.ts copies the runtime into `_shared/temml/`. Same
 *            pattern as xmllint-wasm.
 *
 *   stage 2: mathml2omml translates MathML → OMML (`<m:oMath>...` string).
 *            LGPL-3.0; bundled into the dist for v1. Replacement target in
 *            v2 — see references/equations.md "Not supported".
 *
 * Known mathml2omml bugs we live with at v1:
 *   - n-ary operators (Σ, ∫, ∏) emit an empty `<m:e/>` with the operand
 *     shifted to a sibling. Renders in Word as a dashed placeholder box
 *     before the operand. No workaround at this layer.
 *
 * Synchronization model:
 *   The emit chain is synchronous (one call per Block, recursive into
 *   tables). Temml loads asynchronously via dynamic import. The engine
 *   pre-walks every Block in `edits[]`, calls `prepareLatex` for each
 *   `latex` it finds, then emit calls `getOmmlSync` from the populated
 *   cache. The cache is module-scoped — keeping it stateful is acceptable
 *   for a per-invocation CLI.
 */

import { fileURLToPath, pathToFileURL } from "node:url"
import { dirname, join } from "node:path"
import { mml2omml } from "mathml2omml"
import { validateOMath } from "@lib/shared/docx-validate.ts"

interface TemmlModule {
  renderToString: (
    expression: string,
    options?: { xml?: boolean; displayMode?: boolean; throwOnError?: boolean },
  ) => string
}

const MODULE_DIR = dirname(fileURLToPath(import.meta.url))

let cachedTemml: TemmlModule | null = null
async function loadTemml(): Promise<TemmlModule> {
  if (cachedTemml) return cachedTemml
  // Bare specifier resolves through node_modules in dev. The colocated path
  // resolves the runtime copied alongside cli-helpers.js by build-skill.
  const candidates = ["temml", pathToFileURL(join(MODULE_DIR, "temml", "dist", "temml.cjs")).href]
  let lastErr: unknown = null
  for (const spec of candidates) {
    try {
      const mod = (await import(spec)) as TemmlModule | { default: TemmlModule }
      // CJS interop: when imported from ESM, the Temml namespace object can
      // land on `.default` (CJS build) or on the module record itself (ESM
      // build). Probe both.
      const resolved: TemmlModule = "renderToString" in mod ? mod : mod.default
      cachedTemml = resolved
      return resolved
    } catch (err) {
      lastErr = err
    }
  }
  throw new Error(
    `temml runtime not found. Looked for: ${candidates.join(", ")}. ` +
      `Last error: ${(lastErr as Error)?.message ?? "(unknown)"}.`,
  )
}

/* ---------- pre-resolution cache ---------- */

function cacheKey(latex: string, displayMode: boolean): string {
  return `${displayMode ? "B" : "I"}:${latex}`
}

const ommlCache = new Map<string, string>()

/** Pre-resolve a batch of LaTeX expressions into OMML strings, populating
 * the module cache the synchronous emit functions read from. The engine
 * calls this once before running edits; idempotent + safe to call with
 * duplicate items. */
export async function prepareLatex(
  items: ReadonlyArray<{ latex: string; displayMode: boolean }>,
): Promise<void> {
  if (items.length === 0) return
  const temml = await loadTemml()
  for (const { latex, displayMode } of items) {
    const key = cacheKey(latex, displayMode)
    if (ommlCache.has(key)) continue
    let mathml: string
    try {
      mathml = temml.renderToString(latex, { xml: true, displayMode, throwOnError: true })
    } catch (err) {
      throw new Error(`LaTeX parse error in ${truncateLatex(latex)}: ${(err as Error).message}`, {
        cause: err,
      })
    }
    // mml2omml degrades silently on unsupported MathML nodes (e.g. mpadded
    // from \mathrm{}): it writes "Type not supported: X" to stderr and
    // returns OMML with that subtree dropped. Result is schema-valid but
    // visually wrong, so the math-schema check below won't catch it.
    // Capture stderr around the call and treat any warning as a hard fail
    // — the agent should know they need the omml escape hatch.
    let omml: string
    let stderrCapture: string
    try {
      ;({ result: omml, stderr: stderrCapture } = captureStderrSync(() => mml2omml(mathml)))
    } catch (err) {
      throw new Error(
        `MathML → OMML conversion failed for ${truncateLatex(latex)}: ${(err as Error).message}. ` +
          `Known fragile tokens — see references/equations.md "Known fragile LaTeX tokens"; use the omml escape hatch on the EquationBlock if the failure is unrecoverable.`,
        { cause: err },
      )
    }
    if (stderrCapture.length > 0) {
      const warnings = [
        ...new Set(
          stderrCapture
            .split("\n")
            .map((l) => l.trim())
            .filter(Boolean),
        ),
      ].join("; ")
      throw new Error(
        `MathML → OMML conversion emitted warnings for ${truncateLatex(latex)}: ${warnings} — the affected subtree was silently dropped, so the equation would render incomplete. Switch this equation to the omml escape hatch on the EquationBlock; see references/equations.md "Known fragile LaTeX tokens".`,
      )
    }
    // mml2omml occasionally emits schema-invalid OMML (e.g. <m:rPr> where
    // the math schema rejects it). Catch it here so the throw carries the
    // source LaTeX; post-write validation would surface a bare schema
    // error against input_0.xml with no edit context.
    const schemaErrors = await validateOMath(omml)
    if (schemaErrors.length > 0) {
      throw new Error(
        `OOXML math schema validation failed for ${truncateLatex(latex)}: ${schemaErrors[0]}` +
          (schemaErrors.length > 1 ? ` (+${schemaErrors.length - 1} more)` : "") +
          ` — switch this equation to the omml escape hatch on the EquationBlock; see references/equations.md "Known fragile LaTeX tokens".`,
      )
    }
    ommlCache.set(key, omml)
  }
}

/** Short, JSON-safe rendering of a LaTeX source for error messages.
 * Cap at 80 chars so a long expression doesn't drown the stderr line. */
function truncateLatex(latex: string): string {
  const trimmed = latex.length > 80 ? latex.slice(0, 80) + "…" : latex
  return JSON.stringify(trimmed)
}

/** Run a synchronous fn with process.stderr.write redirected to an
 *  in-memory buffer; return the fn's result plus captured stderr text.
 *  Used to catch mml2omml's warn-and-degrade output that bypasses its
 *  throw path. mml2omml is synchronous, so the global-state hijack stays
 *  scoped to one call. */
function captureStderrSync<T>(fn: () => T): { result: T; stderr: string } {
  let captured = ""
  const orig = process.stderr.write.bind(process.stderr)
  process.stderr.write = ((chunk: unknown) => {
    captured += typeof chunk === "string" ? chunk : String(chunk)
    return true
  }) as typeof process.stderr.write
  try {
    const result = fn()
    return { result, stderr: captured }
  } finally {
    process.stderr.write = orig
  }
}

/** Read a pre-resolved OMML string. Throws when the latex wasn't passed
 * through prepareLatex first — this is an engine bug (the pre-walk missed
 * a code path), not user-actionable. */
export function getOmmlSync(latex: string, displayMode: boolean): string {
  const key = cacheKey(latex, displayMode)
  const omml = ommlCache.get(key)
  if (!omml) {
    throw new Error(
      `internal: equation not pre-resolved (latex=${JSON.stringify(latex)}, displayMode=${displayMode}). ` +
        `Engine must call prepareLatex on all edits before emit.`,
    )
  }
  return omml
}
