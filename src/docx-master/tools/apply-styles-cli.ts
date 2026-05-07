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
  // (`config.styles` is normalized to `[]` by the engine when omitted, so we
  // read it via optional chaining here.)
  validate(config) {
    const hasStyles = (config.styles?.length ?? 0) > 0
    const hasNumbering = !!config.numbering?.levels?.length
    const hasTemplate = !!config.template?.styles?.length
    const hasThemeOverride =
      !!config.theme?.fonts &&
      Object.values(config.theme.fonts).some((v) => typeof v === "string" && v)
    if (!hasStyles && !hasNumbering && !hasTemplate && !hasThemeOverride) {
      throw new Error(
        "config has no operation: provide at least one of styles[] (non-empty), numbering, template, or theme.fonts",
      )
    }
  },
})
