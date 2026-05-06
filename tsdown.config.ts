import { defineConfig } from "tsdown"

// Single build with code splitting: shared deps (jszip, @xmldom, core/) are
// extracted into chunks under scripts/_shared/ so each tool stays small.
export default defineConfig({
  entry: {
    overview: "src/tools/overview.ts",
    inspect_range: "src/tools/inspect-range.ts",
    inspect_runs: "src/tools/inspect-runs.ts",
    inspect_neighbors: "src/tools/inspect-neighbors.ts",
    inspect_style: "src/tools/inspect-style.ts",
    inspect_style_def: "src/tools/inspect-style-def.ts",
    inspect_section: "src/tools/inspect-section.ts",
    find_paragraphs: "src/tools/find-paragraphs.ts",
    apply_styles: "src/tools/apply-styles-cli.ts",
    restyle: "src/tools/restyle.ts",
    migrate_numbering: "src/tools/migrate-numbering.ts",
    import_template: "src/tools/import-template.ts",
  },
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
