#!/usr/bin/env bun
/**
 * Stage SKILL.md + references/ from top-level `skill/` alongside tsdown's
 * scripts/ output, then zip into `dist/docx-master.zip`.
 *
 * The staged bundle lives at `dist/plugin/skills/docx-master/`, which is
 * also the canonical location consumed by Claude Code's plugin marketplace
 * (`.claude-plugin/marketplace.json` → `source: "./dist/plugin"`) AND the
 * exact structure pushed to the publish repo `hawa130/docx-master` by the
 * release workflow. One staged copy serves every consumer; no fan-out.
 *
 * Run AFTER `tsdown` (which writes `<STAGE_DIR>/scripts/`) — the
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
import { dirname, join, relative } from "node:path"
import JSZip from "jszip"

const ROOT = import.meta.dirname
const SKILL_NAME = "docx-master"
const SKILL_SRC = join(ROOT, "skill")
const PUBLISH_DIR = join(ROOT, "dist", "plugin")
const STAGE_DIR = join(PUBLISH_DIR, "skills", SKILL_NAME)
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

// 2b''. Stage the bundled blank.docx template alongside the bundled
// scripts. `lib/apply/blank-source.ts`'s resolver looks one of two ways:
// `<MODULE_DIR>/_assets/blank.docx` (when the function inlines into the
// apply.js entry → MODULE_DIR is SCRIPTS_DIR) and `<MODULE_DIR>/../_assets`
// (when it lands in a `_shared/` chunk). Copying to SCRIPTS_DIR/_assets/
// satisfies both candidates.
const blankSrc = join(ROOT, "lib", "apply", "_assets", "blank.docx")
const blankDst = join(SCRIPTS_DIR, "_assets", "blank.docx")
if (!existsSync(blankSrc)) {
  console.error(`blank.docx template missing at ${relative(ROOT, blankSrc)} — corrupted checkout?`)
  process.exit(1)
}
mkdirSync(dirname(blankDst), { recursive: true })
cpSync(blankSrc, blankDst)

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

// 4. Render .claude-plugin/{marketplace,plugin}.json templates into the
// publish image (dist/plugin/.claude-plugin/). The top-level source copies
// point into dist/ so local `claude plugin install ./` works from this dev
// repo; the rendered copies use the flat layout the publish repo will have
// (.claude-plugin/ and skills/ as siblings at the repo root). Only path
// fields are rewritten — version, description, homepage, repository remain
// canonical in the source files.
const PLUGIN_MANIFEST_SRC = join(ROOT, ".claude-plugin")
const PLUGIN_MANIFEST_DST = join(PUBLISH_DIR, ".claude-plugin")
mkdirSync(PLUGIN_MANIFEST_DST, { recursive: true })

const marketplaceTpl = JSON.parse(
  readFileSync(join(PLUGIN_MANIFEST_SRC, "marketplace.json"), "utf-8"),
) as { plugins: Array<{ source: string; version: string }> }
const firstPlugin = marketplaceTpl.plugins[0]
if (!firstPlugin) {
  throw new Error("marketplace.json: plugins[] is empty")
}
firstPlugin.source = "."

const pluginTpl = JSON.parse(
  readFileSync(join(PLUGIN_MANIFEST_SRC, "plugin.json"), "utf-8"),
) as { skills: string; version: string }
pluginTpl.skills = "./skills/"

// CI injects DOCX_MASTER_VERSION from the git tag (e.g. `v0.2.0` → `0.2.0`).
// Local builds leave source manifests' version untouched.
const versionOverride = process.env.DOCX_MASTER_VERSION
if (versionOverride) {
  firstPlugin.version = versionOverride
  pluginTpl.version = versionOverride
}

writeFileSync(
  join(PLUGIN_MANIFEST_DST, "marketplace.json"),
  JSON.stringify(marketplaceTpl, null, 2) + "\n",
)
writeFileSync(
  join(PLUGIN_MANIFEST_DST, "plugin.json"),
  JSON.stringify(pluginTpl, null, 2) + "\n",
)

// 5. Report
const zipSize = statSync(ZIP_PATH).size
console.log(`✓ Skill bundle:   ${relative(ROOT, STAGE_DIR)}/`)
console.log(`✓ Plugin manifest ${relative(ROOT, PLUGIN_MANIFEST_DST)}/`)
console.log(`✓ Publish image:  ${relative(ROOT, PUBLISH_DIR)}/`)
console.log(`✓ Skill archive:  ${relative(ROOT, ZIP_PATH)} (${formatBytes(zipSize)})`)
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
