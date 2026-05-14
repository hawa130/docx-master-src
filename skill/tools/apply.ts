import { runCli } from "@lib/shared/cli-helpers.ts"

/**
 * CLI entry for `apply` — the unified docx mutation orchestrator.
 *
 * Accepts a config that can combine any of:
 *   - `styles[]` — install / override named paragraph styles
 *   - `numbering` — install / extend multi-level numbering schemes (single
 *     scheme or array)
 *   - `template` — import named styles from a template document
 *   - `theme.fonts` — set the document's design-layer font scheme
 *   - `assignments` / `pattern_rules` / `bulk_rules` / `exclude` — paragraph
 *     restyling rules (rule precedence: exclude > assignments > pattern_rules
 *     > bulk_rules)
 *   - `edits[]` — surgical content insertions / replacements / deletions /
 *     format ops
 *   - `trackChanges` — emit edits as Word revision markup
 *
 * Pipeline order: see SKILL.md §Commands.
 *
 * Sparse by design: only declared blocks are applied. Untouched styles,
 * numbering schemes, paragraphs, and theme stay as they are.
 *
 * Pure content-only edits (no style/numbering install) still go through
 * `apply` with `edits[]` as the only populated block.
 */
void runCli({
  command: "apply",
  script: "apply.js",
  validate(config) {
    const hasStyles = (config.styles?.length ?? 0) > 0
    const hasNumbering = (() => {
      if (!config.numbering) return false
      const schemes = Array.isArray(config.numbering) ? config.numbering : [config.numbering]
      return schemes.some((s) => s.levels.length > 0)
    })()
    const hasTemplate = !!config.template?.styles?.length
    const hasThemeOverride =
      !!config.theme?.fonts &&
      Object.values(config.theme.fonts).some((v) => typeof v === "string" && v)
    const hasEdits = (config.edits?.length ?? 0) > 0
    const hasCaptions = !!config.captions && Object.keys(config.captions).length > 0
    if (
      !hasStyles &&
      !hasNumbering &&
      !hasTemplate &&
      !hasThemeOverride &&
      !hasEdits &&
      !hasCaptions
    ) {
      throw new Error(
        "config has no operation: provide at least one of styles[] (non-empty), numbering, template, theme.fonts, edits[], or captions",
      )
    }
  },
})
