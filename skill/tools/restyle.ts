import { runCli } from "@lib/cli-helpers.ts"

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
void runCli({
  command: "restyle",
  script: "restyle.js",
  validate(config) {
    if (config.template) {
      throw new Error(
        "restyle: `template` is not allowed here. Use `import_template` first (or `apply_styles` for combined operations).",
      )
    }
    if (config.numbering) {
      throw new Error(
        "restyle: `numbering` is not allowed here. Use `migrate_numbering` (or `apply` for combined operations).",
      )
    }
    if (config.edits) {
      throw new Error(
        "restyle: `edits` is not allowed here. Use `apply` for restyle + content insertion in one pass, or `apply_edits` for pure content edits.",
      )
    }
    if (!Array.isArray(config.styles) || config.styles.length === 0) {
      throw new Error("restyle: config.styles must be a non-empty array")
    }
  },
})
