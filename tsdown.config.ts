import { defineConfig } from "tsdown"
import { resolve } from "node:path"

// Single-skill build. The `docx-master` skill source lives at
// `src/docx-master/`; tsdown bundles each entry under `tools/` into
// `dist/docx-master/scripts/<name>.js`. Cross-cutting OOXML primitives are
// imported via the `@core/*` alias.
//
// Adding a tool: add a `<scriptName>: "<file>.ts"` entry to `tools` below.
// Tool source files live under `src/docx-master/tools/`.

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
  find_paragraphs: "find-paragraphs.ts",
  // normalize sub-command
  apply_styles: "apply-styles-cli.ts",
  restyle: "restyle.ts",
  migrate_numbering: "migrate-numbering.ts",
  import_template: "import-template.ts",
}

export default defineConfig({
  entry: Object.fromEntries(
    Object.entries(tools).map(([scriptName, file]) => [
      scriptName,
      `src/docx-master/tools/${file}`,
    ]),
  ),
  alias: { "@core": resolve(ROOT, "src/core") },
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
