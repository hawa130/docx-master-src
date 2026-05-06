import { readFileSync, existsSync, unlinkSync } from "node:fs"
import { resolve } from "node:path"
import { applyStyles, type ApplyConfig } from "./apply-styles.ts"

/**
 * Narrow entry point for importing named styles from a template document
 * into the source's styles.xml.
 *
 * Accepts a config that requires `template`. `styles[]` is optional — when
 * you only want to pull in the template's style system without overriding
 * any of its imported definitions, leave `styles[]` empty. If you do declare
 * `styles[]`, those entries override matching imported styles (template is
 * the baseline, user-declared wins on conflict).
 *
 * `numbering` is rejected here; if the template's imported styles reference
 * numbering, `importNumbering` (default true) handles that automatically. To
 * install a *new* numbering scheme separately, run `migrate_numbering` after.
 */
async function main() {
  const args = process.argv.slice(2)
  const dryRun = args.includes("--dry-run")
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error("Usage: node scripts/import_template.js [--dry-run] <config.json>")
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
  if (!config.template || !Array.isArray(config.template.styles) || config.template.styles.length === 0) {
    console.error("import_template: config.template.styles must be a non-empty array of styleIds")
    process.exit(1)
  }
  if (config.numbering) {
    console.error(
      "import_template: `numbering` is not allowed here. Imported styles' numbering references are migrated automatically; for a new numbering scheme, run `migrate_numbering` after.",
    )
    process.exit(1)
  }
  // styles[] is optional. Default to empty so applyStyles' gate sees
  // `template` and accepts the run.
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
