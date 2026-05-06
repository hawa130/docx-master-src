import { readFileSync, existsSync, unlinkSync } from "node:fs"
import { resolve } from "node:path"
import { applyStyles, type ApplyConfig } from "./apply-styles.ts"

/**
 * Narrow entry point for installing / replacing a multi-level auto-numbering
 * scheme bound to existing heading styles.
 *
 * Accepts a config that requires `numbering`. `styles[]` is optional — when
 * the doc already defines the heading styles you're binding to (Heading1,
 * Heading2, ...), no redeclaration is needed. The orchestrator validates each
 * `numbering.levels[].styleId` against both `config.styles[]` and the source
 * doc's existing styles.xml, so typos are still caught.
 *
 * `template` is rejected here; chain `import_template` first if you need to
 * pull in heading styles from a template, then run `migrate_numbering` to
 * bind the numbering. (Or use `apply_styles` for combined operations.)
 */
async function main() {
  const args = process.argv.slice(2)
  const dryRun = args.includes("--dry-run")
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error("Usage: node scripts/migrate_numbering.js [--dry-run] <config.json>")
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
      "migrate_numbering: `template` is not allowed here. Run `import_template` first if you need template styles, then `migrate_numbering`.",
    )
    process.exit(1)
  }
  if (!config.numbering || !Array.isArray(config.numbering.levels) || config.numbering.levels.length === 0) {
    console.error("migrate_numbering: config.numbering.levels must be a non-empty array")
    process.exit(1)
  }
  // styles[] is optional on this path. Default to empty so applyStyles'
  // gate ("at least one operation") sees `numbering` and accepts the run.
  if (!Array.isArray(config.styles)) {
    config.styles = []
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
