import { runCli } from "../lib/cli-helpers.ts"

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
void runCli({
  command: "migrate_numbering",
  script: "migrate_numbering.js",
  validate(config) {
    if (config.template) {
      throw new Error(
        "migrate_numbering: `template` is not allowed here. Run `import_template` first if you need template styles, then `migrate_numbering`.",
      )
    }
    if (
      !config.numbering ||
      !Array.isArray(config.numbering.levels) ||
      config.numbering.levels.length === 0
    ) {
      throw new Error(
        "migrate_numbering: config.numbering.levels must be a non-empty array",
      )
    }
    // styles[] is optional on this path. Default to empty so applyStyles'
    // gate ("at least one operation") sees `numbering` and accepts the run.
    if (!Array.isArray(config.styles)) config.styles = []
  },
})
