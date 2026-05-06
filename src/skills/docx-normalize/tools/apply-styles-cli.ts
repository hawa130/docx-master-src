import { runCli } from "../lib/cli-helpers.ts"

/**
 * CLI entry for `apply_styles` — the unified standardization orchestrator.
 *
 * Accepts a config that can combine `styles[]`, `numbering`, and `template`
 * in one pass. The orchestrator runs them in the correct order and produces
 * a single change report. Use this when you have a complete picture of the
 * transformation; for narrower changes, the dedicated CLIs (`restyle`,
 * `migrate_numbering`, `import_template`) accept smaller configs.
 */
void runCli({
  command: "apply_styles",
  script: "apply_styles.js",
  // At least one operation must be specified. Pure restyle needs styles[];
  // pure numbering migration needs numbering; pure template import needs
  // template. With nothing, the script would just produce a copy.
  validate(config) {
    if (!Array.isArray(config.styles)) {
      throw new Error(
        "config.styles must be an array (may be empty if `numbering` or `template` is supplied)",
      )
    }
    const hasStyles = config.styles.length > 0
    const hasNumbering = !!config.numbering?.levels?.length
    const hasTemplate = !!config.template?.styles?.length
    if (!hasStyles && !hasNumbering && !hasTemplate) {
      throw new Error(
        "config has no operation: provide at least one of styles[] (non-empty), numbering, or template",
      )
    }
  },
})
