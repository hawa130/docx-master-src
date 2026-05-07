import { existsSync, readFileSync } from "node:fs"
import { resolve } from "node:path"
import { applyEdits } from "@lib/edit-engine.ts"

/**
 * CLI entry for `apply_edits` — surgical edits on an existing docx.
 *
 * Config shape (validated by lib/edit-config-schema.ts):
 *   - source / output: docx paths
 *   - edits[]: ordered ops (replace / insert-before / insert-after / delete / format)
 *   - trackChanges?: emit ins/del/rPrChange/pPrChange so Word's review UI
 *     shows accept/reject; default false (silent mutation)
 *
 * Output: single-line summary plus per-op touched count. Failures abort
 * with no file written; the original is never modified.
 */

async function main() {
  const args = process.argv.slice(2)
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error("Usage: node scripts/apply_edits.js <config.json>")
    process.exit(1)
  }
  let raw: unknown
  try {
    raw = JSON.parse(readFileSync(configPath, "utf8"))
  } catch (err) {
    console.error(`Cannot read config: ${(err as Error).message}`)
    process.exit(1)
  }

  // Universal source/output sanity (mirroring cli-helpers.ts for the
  // standardize family). The schema enforces non-empty strings; we add the
  // exists / not-same-file checks here against the resolved paths.
  const obj = raw as { source?: string; output?: string }
  if (!obj.source || !obj.output) {
    console.error("config must include source and output paths")
    process.exit(1)
  }
  const source = resolve(obj.source)
  const output = resolve(obj.output)
  if (source === output) {
    console.error("output must differ from source")
    process.exit(1)
  }
  if (!existsSync(source)) {
    console.error(`source not found: ${source}`)
    process.exit(1)
  }
  // Engine reopens via DocxReader.open(output) after copying, so configs
  // referencing source via relative path need the resolved value reflected
  // back on the raw object.
  ;(raw as { source: string }).source = source
  ;(raw as { output: string }).output = output

  try {
    const report = await applyEdits(source, output, raw)
    const blockerStr =
      report.blockerCounts["tracked-change"] ||
      report.blockerCounts.field ||
      report.blockerCounts.sdt
        ? `  (skipped zones: ${report.blockerCounts["tracked-change"]} tracked, ${report.blockerCounts.field} field, ${report.blockerCounts.sdt} sdt)`
        : ""
    console.log(
      `✓ ${report.applied} edit(s) applied${report.trackChanges ? " (track-changes)" : ""}${blockerStr}`,
    )
    for (const r of report.perOp) {
      console.log(`  edits[${r.index}] ${r.op}: ${r.touched} paragraph(s)`)
    }
    console.log(`Output: ${output}`)
  } catch (err) {
    console.error((err as Error).message)
    process.exit(1)
  }
}

void main()
