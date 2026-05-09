import { defineConfig } from "tsdown"
import { resolve } from "node:path"

// Single-skill build. The `docx-master` skill source lives at top-level
// `skill/`; tsdown bundles each entry under `skill/tools/` into
// `dist/docx-master/scripts/<name>.js`. All non-tool TypeScript — OOXML
// primitives, skill engine, config schema, CLI scaffolding — lives in
// top-level `lib/` and is reachable via the `@lib/*` alias.
//
// Adding a tool: add a `<scriptName>: "<file>.ts"` entry to `tools` below.
// Tool source files live under `skill/tools/`.

const ROOT = import.meta.dirname

const tools: Record<string, string> = {
  // inspect
  overview: "overview.ts",
  inspect_range: "inspect-range.ts",
  inspect_runs: "inspect-runs.ts",
  inspect_neighbors: "inspect-neighbors.ts",
  inspect_style: "inspect-style.ts",
  inspect_style_def: "inspect-style-def.ts",
  inspect_section: "inspect-section.ts",
  inspect_table: "inspect-table.ts",
  inspect_blockers: "inspect-blockers.ts",
  find_paragraphs: "find-paragraphs.ts",
  // standardize sub-command
  apply: "apply.ts",
  restyle: "restyle.ts",
  migrate_numbering: "migrate-numbering.ts",
  import_template: "import-template.ts",
}

export default defineConfig({
  entry: Object.fromEntries(
    Object.entries(tools).map(([scriptName, file]) => [scriptName, `skill/tools/${file}`]),
  ),
  alias: { "@lib": resolve(ROOT, "lib") },
  format: "esm",
  platform: "node",
  target: "node18",
  outDir: "dist/docx-master/scripts",
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
