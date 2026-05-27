import { existsSync, readFileSync, unlinkSync } from "node:fs"
import { resolve } from "node:path"
import { applyStyles } from "@lib/apply/apply-styles.ts"
import { getBlankTemplatePath } from "@lib/apply/blank-source.ts"
import { parseConfig, type ApplyConfig } from "@lib/config/config-schema.ts"

/**
 * Shared scaffolding for the apply-family CLIs (`apply` and the narrow
 * variants that share its config schema). Each one parses the same
 * `[--dry-run] <config.json>` argv shape, validates the same source/output
 * fields, and runs `applyStyles` with the same error-cleanup behavior. The
 * differences are which config blocks are required vs forbidden — handled
 * per-CLI by passing a `validate` callback.
 *
 * The exact stderr wording and exit codes are part of the agent-facing
 * contract; downstream agents pattern-match on them. Don't tweak them
 * without coordinating a SKILL.md update.
 */

export interface CliSpec {
  /** CLI name as it appears in usage strings, e.g. "apply". */
  command: string
  /** Script filename for the usage line, e.g. "apply.js". */
  script: string
  /** Validates the parsed config beyond the universal source/output checks.
   * Throw an Error with a CLI-friendly message on failure. May mutate
   * `config` (e.g. to default `styles[]` to `[]`). */
  validate: (config: ApplyConfig) => void
}

export async function runCli(spec: CliSpec): Promise<void> {
  const args = process.argv.slice(2)
  const dryRun = args.includes("--dry-run")
  const allowValidationWarnings = args.includes("--allow-validation-warnings")
  const configPath = args.filter((a) => !a.startsWith("--"))[0]
  if (!configPath) {
    console.error(`Usage: node scripts/${spec.script} [--dry-run] [--allow-validation-warnings] <config.json>`)
    process.exit(1)
  }
  let raw: unknown
  try {
    raw = JSON.parse(readFileSync(configPath, "utf8"))
  } catch (err) {
    console.error(`Cannot read config: ${(err as Error).message}`)
    process.exit(1)
  }
  // Apply --dry-run / --allow-validation-warnings before schema parse so
  // they're captured by the schema's boolean check (and survive
  // strictObject's unknown-key rejection).
  if (raw && typeof raw === "object") {
    const rawObj = raw as { dryRun?: boolean; allowValidationWarnings?: boolean }
    if (dryRun) rawObj.dryRun = true
    if (allowValidationWarnings) rawObj.allowValidationWarnings = true
  }

  let config: ApplyConfig
  try {
    config = parseConfig(raw)
  } catch (err) {
    console.error((err as Error).message)
    process.exit(1)
  }

  try {
    spec.validate(config)
  } catch (err) {
    console.error((err as Error).message)
    process.exit(1)
  }

  // Omitted source → scaffold from the bundled blank.docx template. The
  // template ships at lib/apply/_assets/blank.docx (dev) and is copied
  // into <scripts>/_assets/ by build-skill.
  const source = config.source !== undefined ? resolve(config.source) : getBlankTemplatePath()
  const output = resolve(config.output)
  if (!config.dryRun && source === output) {
    console.error("output must differ from source")
    process.exit(1)
  }
  if (!existsSync(source)) {
    console.error(
      config.source !== undefined
        ? `source not found: ${source}`
        : `bundled blank template not found at ${source} — run \`bun run build:skill\` to stage it`,
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
