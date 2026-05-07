import { runCli } from "../lib/cli-helpers.ts"

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
void runCli({
  command: "import_template",
  script: "import_template.js",
  validate(config) {
    if (
      !config.template ||
      !Array.isArray(config.template.styles) ||
      config.template.styles.length === 0
    ) {
      throw new Error(
        "import_template: config.template.styles must be a non-empty array of styleIds",
      )
    }
    if (config.numbering) {
      throw new Error(
        "import_template: `numbering` is not allowed here. Imported styles' numbering references are migrated automatically; for a new numbering scheme, run `migrate_numbering` after.",
      )
    }
    // styles[] is optional; the engine defaults it to [].
  },
})
