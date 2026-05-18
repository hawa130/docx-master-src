import { defineConfig } from "tsdown"
import { resolve } from "node:path"

// Single-skill build. The `docx-master` skill source lives at top-level
// `skill/`; tsdown bundles each entry under `skill/tools/` into
// `dist/plugin/skills/docx-master/scripts/<name>.js`. All non-tool TypeScript
// — OOXML primitives, skill engine, config schema, CLI scaffolding — lives
// in top-level `lib/` and is reachable via the `@lib/*` alias.
//
// The outDir nests directly inside the publish-repo layout (`dist/plugin/`
// is a complete image of what hawa130/docx-master will hold), so the
// canonical skill location is `dist/plugin/skills/docx-master/` and no
// post-build "fan-out" copy is needed.
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
  inspect_caption: "inspect-caption.ts",
  find_paragraphs: "find-paragraphs.ts",
  find_text: "find-text.ts",
  validate: "validate.ts",
  // standardize sub-command
  apply: "apply.ts",
  migrate_captions: "migrate-captions.ts",
}

export default defineConfig({
  entry: Object.fromEntries(
    Object.entries(tools).map(([scriptName, file]) => [scriptName, `skill/tools/${file}`]),
  ),
  alias: { "@lib": resolve(ROOT, "lib") },
  format: "esm",
  platform: "node",
  target: "node18",
  outDir: "dist/plugin/skills/docx-master/scripts",
  outExtensions: () => ({ js: ".js" }),
  clean: true,
  dts: false,
  shims: true,
  sourcemap: false,
  minify: false,
  // Bundle every dep except a handful that bundling corrupts or breaks:
  //   - xmllint-wasm: index-node.js does runtime `require("./xmllint-node.js")`
  //     and references the `.wasm` by relative path; bundling rewrites those
  //     paths out of existence.
  //   - temml: source uses `\uD800-\uDFFF` lone-surrogate ranges in string
  //     literals to drive its tokenizer regex. oxc/rolldown replaces lone
  //     surrogates with U+FFFD on output, corrupting the regex so `\frac`
  //     and friends fail to tokenize.
  // build-skill.ts copies each package's runtime files into _shared/
  // alongside the bundled scripts so the dynamic-import fallback resolves.
  external: ["xmllint-wasm", "temml"],
  deps: { alwaysBundle: [/.*/] },
  outputOptions: {
    chunkFileNames: "_shared/[name].js",
    hashCharacters: "base36",
  },
})
