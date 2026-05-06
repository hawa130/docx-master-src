import { defineConfig } from "tsdown"
import { resolve } from "node:path"

// Each skill's tools build into dist/<skill-name>/scripts/. Shared `src/core/`
// modules are imported via the "@core/*" alias (also declared in tsconfig.json
// for editor / type-check support); rolldown rewrites these at build time.
//
// To add a new skill: drop a new directory under src/skills/<name>/tools/ and
// append its entries below. The shared core gets bundled into each skill's
// scripts/_shared/ chunk independently, so skills are deliverable on their own.

const SHARED_ALIAS = { "@core": resolve(import.meta.dirname, "src/core") }

const docxNormalize = defineConfig({
  entry: {
    overview: "src/skills/docx-normalize/tools/overview.ts",
    inspect_range: "src/skills/docx-normalize/tools/inspect-range.ts",
    inspect_runs: "src/skills/docx-normalize/tools/inspect-runs.ts",
    inspect_neighbors: "src/skills/docx-normalize/tools/inspect-neighbors.ts",
    inspect_style: "src/skills/docx-normalize/tools/inspect-style.ts",
    inspect_style_def: "src/skills/docx-normalize/tools/inspect-style-def.ts",
    inspect_section: "src/skills/docx-normalize/tools/inspect-section.ts",
    find_paragraphs: "src/skills/docx-normalize/tools/find-paragraphs.ts",
    apply_styles: "src/skills/docx-normalize/tools/apply-styles-cli.ts",
    restyle: "src/skills/docx-normalize/tools/restyle.ts",
    migrate_numbering: "src/skills/docx-normalize/tools/migrate-numbering.ts",
    import_template: "src/skills/docx-normalize/tools/import-template.ts",
  },
  alias: SHARED_ALIAS,
  format: "esm",
  platform: "node",
  target: "node18",
  outDir: "dist/docx-normalize/scripts",
  outExtensions: () => ({ js: ".js" }),
  clean: true,
  dts: false,
  shims: true,
  sourcemap: false,
  minify: false,
  deps: { alwaysBundle: [/.*/] },
  outputOptions: {
    chunkFileNames: "_shared/[name].js",
    hashCharacters: "base36",
  },
})

export default docxNormalize
