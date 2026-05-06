import { defineConfig, type UserConfig } from "tsdown"
import { resolve } from "node:path"

// Multi-skill build. tsdown's `defineConfig` accepts an array of UserConfigs;
// each entry below is one skill, built into its own dist/<name>/scripts/.
//
// Adding a skill: append another `skill(name, tools)` block. Adding a tool:
// add a key to the existing block. The helper handles the path prefix and
// shared options so per-skill blocks stay focused on what's actually skill-
// specific (the script-name → source-file map).

const ROOT = import.meta.dirname
const SHARED_ALIAS = { "@core": resolve(ROOT, "src/core") }

function skill(name: string, tools: Record<string, string>): UserConfig {
  return {
    entry: Object.fromEntries(
      Object.entries(tools).map(([scriptName, file]) => [
        scriptName,
        `src/skills/${name}/tools/${file}`,
      ]),
    ),
    alias: SHARED_ALIAS,
    format: "esm",
    platform: "node",
    target: "node18",
    outDir: `dist/${name}/scripts`,
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
  }
}

export default defineConfig([
  skill("docx-normalize", {
    overview: "overview.ts",
    inspect_range: "inspect-range.ts",
    inspect_runs: "inspect-runs.ts",
    inspect_neighbors: "inspect-neighbors.ts",
    inspect_style: "inspect-style.ts",
    inspect_style_def: "inspect-style-def.ts",
    inspect_section: "inspect-section.ts",
    find_paragraphs: "find-paragraphs.ts",
    apply_styles: "apply-styles-cli.ts",
    restyle: "restyle.ts",
    migrate_numbering: "migrate-numbering.ts",
    import_template: "import-template.ts",
  }),
  // Future:
  // skill("docx-cleanup", { accept_changes: "accept-changes-cli.ts", ... }),
  // skill("docx-augment", { manage_captions: "manage-captions-cli.ts", ... }),
])
