import { readFileSync, existsSync, unlinkSync } from "node:fs"
import { resolve } from "node:path"
import { applyStyles, type ApplyConfig } from "./apply-styles.ts"

/**
 * CLI entry for `apply_styles` — the unified standardization orchestrator.
 *
 * Accepts a config that can combine `styles[]`, `numbering`, and `template`
 * in one pass. The orchestrator runs them in the correct order and produces
 * a single change report. Use this when you have a complete picture of the
 * transformation; for narrower changes, the dedicated CLIs (`restyle`,
 * `migrate_numbering`, `import_template`) accept smaller configs.
 */
async function main() {
  const args = process.argv.slice(2)
  const dryRun = args.includes("--dry-run")
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error("Usage: node scripts/apply_styles.js [--dry-run] <config.json>")
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
  // At least one operation must be specified. Pure restyle needs styles[];
  // pure numbering migration needs numbering; pure template import needs
  // template. With nothing, the script would just produce a copy. Narrow
  // wrappers (restyle / migrate_numbering / import_template) further
  // constrain which operations are allowed via their own gates.
  if (!Array.isArray(config.styles)) {
    console.error("config.styles must be an array (may be empty if `numbering` or `template` is supplied)")
    process.exit(1)
  }
  const hasStyles = config.styles.length > 0
  const hasNumbering = !!(config.numbering?.levels?.length)
  const hasTemplate = !!(config.template?.styles?.length)
  if (!hasStyles && !hasNumbering && !hasTemplate) {
    console.error(
      "config has no operation: provide at least one of styles[] (non-empty), numbering, or template",
    )
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
