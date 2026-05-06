import { readFileSync, existsSync, unlinkSync } from "node:fs"
import { resolve } from "node:path"
import { applyStyles, type ApplyConfig } from "./apply-styles.ts"

/**
 * Narrow entry point for paragraph restyling.
 *
 * Accepts the same config shape as `apply_styles` but rejects the `template`
 * and `numbering` blocks — those have their own dedicated CLIs
 * (`import_template`, `migrate_numbering`). This narrowing exists so the
 * Targeted Edit path doesn't need to mentally hold a schema with operations
 * it isn't using; a focused tool with a focused config is faster to compose
 * and faster to verify.
 *
 * For the full-standardization workflow that combines restyle + numbering +
 * template in one pass, use `apply_styles` directly.
 */
async function main() {
  const args = process.argv.slice(2)
  const dryRun = args.includes("--dry-run")
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error("Usage: node scripts/restyle.js [--dry-run] <config.json>")
    process.exit(1)
  }
  let config: ApplyConfig
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"))
  } catch (err) {
    console.error(`Cannot read config: ${(err as Error).message}`)
    process.exit(1)
  }
  if (dryRun) (config as ApplyConfig & { dryRun?: boolean }).dryRun = true

  if (!config.source || !config.output) {
    console.error("config.source and config.output are required")
    process.exit(1)
  }
  if (config.template) {
    console.error(
      "restyle: `template` is not allowed here. Use `import_template` first (or `apply_styles` for combined operations).",
    )
    process.exit(1)
  }
  if (config.numbering) {
    console.error(
      "restyle: `numbering` is not allowed here. Use `migrate_numbering` (or `apply_styles` for combined operations).",
    )
    process.exit(1)
  }
  if (!Array.isArray(config.styles) || config.styles.length === 0) {
    console.error("restyle: config.styles must be a non-empty array")
    process.exit(1)
  }

  const source = resolve(config.source)
  const output = resolve(config.output)
  if (!config.dryRun && source === output) {
    console.error("output must differ from source")
    process.exit(1)
  }
  if (!existsSync(source)) {
    console.error(`source not found: ${source}`)
    process.exit(1)
  }

  try {
    await applyStyles(source, output, config)
  } catch (err) {
    if (existsSync(output)) {
      try {
        unlinkSync(output)
      } catch {}
    }
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

void main()
