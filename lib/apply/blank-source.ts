/**
 * Locator for the bundled `blank.docx` template that apply uses when a
 * config omits `source`. The asset ships two ways:
 *
 *   1. Dev mode (running TS directly from the repo) — sits at
 *      `lib/apply/_assets/blank.docx` relative to the repo root.
 *   2. Bundled mode (running from `dist/docx-master/scripts/...`) —
 *      `build-skill.ts` copies the asset into `<scripts>/_assets/blank.docx`
 *      so a single basename-relative lookup resolves to the same file.
 *
 * Resolution mirrors `findSchemasDir` in `docx-validate.ts` — two
 * candidates, first hit wins. Throws on miss with a build-skill hint.
 */

import { existsSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

const MODULE_DIR = dirname(fileURLToPath(import.meta.url))

/** Absolute path to the bundled blank.docx template. The dev candidate
 *  resolves against this file's own location (`lib/apply/_assets/`);
 *  the bundled candidate accounts for tsdown's chunking — depending on
 *  whether this module gets inlined into the apply.js entry or split
 *  into `_shared/`, `import.meta.url` resolves to a different output
 *  file, so we look both next to the resolved location and one hop up.
 *  Throws on miss. */
export function getBlankTemplatePath(): string {
  const candidates = [
    join(MODULE_DIR, "_assets", "blank.docx"),
    // Shared-chunk landing (`<scripts>/_shared/...js`) → asset is one
    // level up in `<scripts>/_assets/`.
    join(MODULE_DIR, "..", "_assets", "blank.docx"),
  ]
  for (const p of candidates) {
    if (existsSync(p)) return p
  }
  throw new Error(`bundled blank.docx not found. Looked in: ${candidates.join(", ")}.`)
}
