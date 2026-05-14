#!/usr/bin/env bun
/**
 * Stage SKILL.md + references/ from top-level `skill/` alongside tsdown's
 * scripts/ output, then zip into `dist/docx-master.zip`.
 *
 * Run AFTER `tsdown` (which produces `dist/docx-master/scripts/`) — the
 * `bun run build:skill` script chains both.
 */
import {
  cpSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs"
import { join, relative } from "node:path"
import JSZip from "jszip"

const ROOT = import.meta.dirname
const SKILL_NAME = "docx-master"
const SKILL_SRC = join(ROOT, "skill")
const STAGE_DIR = join(ROOT, "dist", SKILL_NAME)
const ZIP_PATH = join(ROOT, "dist", `${SKILL_NAME}.zip`)
const SCRIPTS_DIR = join(STAGE_DIR, "scripts")

if (!existsSync(SCRIPTS_DIR)) {
  console.error(
    `${relative(ROOT, SCRIPTS_DIR)}/ not found — run \`bun run build\` (tsdown) before build-skill.`,
  )
  process.exit(1)
}

// 1. Copy SKILL.md
const skillMdSrc = join(SKILL_SRC, "SKILL.md")
if (!existsSync(skillMdSrc)) {
  throw new Error(`missing SKILL.md at ${relative(ROOT, skillMdSrc)}`)
}
cpSync(skillMdSrc, join(STAGE_DIR, "SKILL.md"))

// 2. Copy references/ (replacing any prior staged copy)
const refsSrc = join(SKILL_SRC, "references")
const refsDst = join(STAGE_DIR, "references")
rmSync(refsDst, { recursive: true, force: true })
if (existsSync(refsSrc)) {
  cpSync(refsSrc, refsDst, { recursive: true })
}

// 2b. Copy OOXML XSDs into scripts/_shared/schemas. xmllint-wasm needs the
// full ECMA-376 + OPC + MCE schema set at runtime to validate; loading from
// disk under the bundle keeps the skill a single self-contained drop-in.
const schemasSrc = join(ROOT, "vendor", "ooxml-schemas")
const schemasDst = join(SCRIPTS_DIR, "_shared", "schemas")
rmSync(schemasDst, { recursive: true, force: true })
if (existsSync(schemasSrc)) {
  cpSync(schemasSrc, schemasDst, { recursive: true })
}

// 2b'. Drop the NOTICE.md from the staged copy — it's source-tree
// provenance documentation, not a runtime artifact. The original lives at
// vendor/ooxml-schemas/NOTICE.md in the source repo.
const stagedNotice = join(schemasDst, "NOTICE.md")
if (existsSync(stagedNotice)) rmSync(stagedNotice)

// 2c. Copy xmllint-wasm runtime files (index-node.js + xmllint-node.js +
// xmllint.wasm) into scripts/_shared/xmllint-wasm. The package can't be
// bundled by tsdown because its index does `require("./xmllint-node.js")`
// and references the .wasm by relative path; bundling rewrites those paths
// out of existence. Copying the runtime tree preserves the require chain.
const xmllintSrc = join(ROOT, "node_modules", "xmllint-wasm")
const xmllintDst = join(SCRIPTS_DIR, "_shared", "xmllint-wasm")
rmSync(xmllintDst, { recursive: true, force: true })
if (!existsSync(xmllintSrc)) {
  console.error(`xmllint-wasm not installed; run \`bun install\` first.`)
  process.exit(1)
}
cpSync(xmllintSrc, xmllintDst, {
  recursive: true,
  filter: (src) => {
    const base = src.split("/").pop() ?? ""
    return !["README.md", "COPYING", ".npmignore"].includes(base)
  },
})

// 2d. Copy temml runtime. Marked external in tsdown.config because the
// bundler corrupts the surrogate-range string literals in temml's
// tokenizer. We only need temml.cjs + package.json (so Node's module
// resolution finds the CJS entry); CSS, web fonts, and the unbuilt source
// tree are not needed at runtime.
const temmlSrc = join(ROOT, "node_modules", "temml")
const temmlDst = join(SCRIPTS_DIR, "_shared", "temml")
rmSync(temmlDst, { recursive: true, force: true })
if (!existsSync(temmlSrc)) {
  console.error(`temml not installed; run \`bun install\` first.`)
  process.exit(1)
}
const temmlDistDst = join(temmlDst, "dist")
mkdirSync(temmlDistDst, { recursive: true })
cpSync(join(temmlSrc, "dist", "temml.cjs"), join(temmlDistDst, "temml.cjs"))
cpSync(join(temmlSrc, "package.json"), join(temmlDst, "package.json"))

// 3. Zip the staged dir
await zipDir(STAGE_DIR, ZIP_PATH, SKILL_NAME)

// 4. Fan out to per-harness directories.
//
// docx-master's SKILL.md frontmatter (name + description) is universal across
// every harness that loads Markdown skills. The only thing that differs is
// where each harness expects skills to live. We stage one identical copy of
// the bundle into each provider's config directory so users can `cp -r` the
// matching tree into their project without learning the layout themselves.
//
// Layout produced under dist/:
//   claude-code/.claude/skills/docx-master/
//   cursor/.cursor/skills/docx-master/
//   codex/.agents/skills/docx-master/      (Codex repo skills, also used user-wide)
//   gemini/.gemini/skills/docx-master/
//   opencode/.opencode/skills/docx-master/
//   github/.github/skills/docx-master/
//
// `plugin/skills/docx-master/` at the repo root mirrors the bundle for Claude
// Code's marketplace install path (see .claude-plugin/marketplace.json).
const PROVIDERS = [
  { dir: "claude-code", config: ".claude", display: "Claude Code" },
  { dir: "cursor", config: ".cursor", display: "Cursor" },
  { dir: "codex", config: ".agents", display: "Codex CLI" },
  { dir: "gemini", config: ".gemini", display: "Gemini CLI" },
  { dir: "opencode", config: ".opencode", display: "OpenCode" },
  { dir: "github", config: ".github", display: "GitHub Copilot" },
]

for (const p of PROVIDERS) {
  const dest = join(ROOT, "dist", p.dir, p.config, "skills", SKILL_NAME)
  rmSync(dest, { recursive: true, force: true })
  mkdirSync(dest, { recursive: true })
  cpSync(STAGE_DIR, dest, { recursive: true })
}

// dist/plugin/ mirror, used by .claude-plugin/marketplace.json's
// `source: "./dist/plugin"` reference. Lives under dist/ alongside every
// other build artifact so a single gitignore line covers them all.
const PLUGIN_DIR = join(ROOT, "dist", "plugin", "skills", SKILL_NAME)
rmSync(PLUGIN_DIR, { recursive: true, force: true })
mkdirSync(PLUGIN_DIR, { recursive: true })
cpSync(STAGE_DIR, PLUGIN_DIR, { recursive: true })

// 5. Report
const zipSize = statSync(ZIP_PATH).size
console.log(`✓ Skill bundle:  ${relative(ROOT, STAGE_DIR)}/`)
console.log(`✓ Skill archive: ${relative(ROOT, ZIP_PATH)} (${formatBytes(zipSize)})`)
for (const p of PROVIDERS) {
  const dest = join("dist", p.dir, p.config, "skills", SKILL_NAME)
  console.log(`✓ ${p.display.padEnd(16)} ${dest}/`)
}
console.log(`✓ Plugin mirror   ${relative(ROOT, PLUGIN_DIR)}/`)
console.log("")
console.log("Contents:")
listTree(STAGE_DIR, "  ")
console.log("")

/* ---------- helpers ---------- */

async function zipDir(srcDir: string, outPath: string, archiveRoot: string) {
  const zip = new JSZip()
  walk(srcDir, (filePath) => {
    const rel = relative(srcDir, filePath)
    const archivePath = join(archiveRoot, rel).replaceAll("\\", "/")
    zip.file(archivePath, readFileSync(filePath))
  })
  const buf = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  })
  writeFileSync(outPath, buf)
}

function walk(dir: string, fn: (file: string) => void) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name)
    if (entry.isDirectory()) walk(p, fn)
    else if (entry.isFile()) fn(p)
  }
}

function listTree(dir: string, indent: string) {
  const entries = readdirSync(dir, { withFileTypes: true }).sort((a, b) =>
    a.name.localeCompare(b.name),
  )
  for (const e of entries) {
    const p = join(dir, e.name)
    if (e.isDirectory()) {
      console.log(`${indent}${e.name}/`)
      listTree(p, indent + "  ")
    } else {
      const size = statSync(p).size
      console.log(`${indent}${e.name}  (${formatBytes(size)})`)
    }
  }
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`
  return `${(n / 1024 / 1024).toFixed(2)} MB`
}
