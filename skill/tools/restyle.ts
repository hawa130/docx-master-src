import { runCli } from "@lib/cli-helpers.ts"

/**
 * Narrow entry point for paragraph restyling.
 *
 * Accepts the same config shape as `apply` but rejects the `template`,
 * `numbering`, and `edits` blocks — those have their own dedicated CLIs
 * (`import_template`, `migrate_numbering`) or belong on `apply` for combined
 * operations. This narrowing exists so the Targeted Edit path doesn't need
 * to mentally hold a schema with operations it isn't using; a focused tool
 * with a focused config is faster to compose and faster to verify.
 *
 * For the full-standardization workflow that combines restyle + numbering +
 * template + edits in one pass, use `apply` directly.
 */
void runCli({
  command: "restyle",
  script: "restyle.js",
  validate(config) {
    if (config.template) {
      throw new Error(
        "restyle: `template` is not allowed here. Use `import_template` first (or `apply` for combined operations).",
      )
    }
    if (config.numbering) {
      throw new Error(
        "restyle: `numbering` is not allowed here. Use `migrate_numbering` (or `apply` for combined operations).",
      )
    }
    if (config.edits) {
      throw new Error(
        "restyle: `edits` is not allowed here. Use `apply` for restyle + content insertion in one pass.",
      )
    }
    if (!Array.isArray(config.styles) || config.styles.length === 0) {
      throw new Error("restyle: config.styles must be a non-empty array")
    }
  },
})
