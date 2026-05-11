import { runCli } from "@lib/shared/cli-helpers.ts"

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
 * bind the numbering. (Or use `apply` for combined operations.)
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
    if (!config.numbering) {
      throw new Error("migrate_numbering: config.numbering is required")
    }
    const schemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
    if (schemes.length === 0 || schemes.every((s) => s.levels.length === 0)) {
      throw new Error(
        "migrate_numbering: config.numbering must contain at least one scheme with levels",
      )
    }
    if (config.edits) {
      throw new Error(
        "migrate_numbering: `edits` is not allowed here. Use `apply` for numbering + content insertion in one pass.",
      )
    }
    // styles[] is optional on this path; the engine defaults it to [].
  },
})
