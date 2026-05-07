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

// 3. Zip the staged dir
await zipDir(STAGE_DIR, ZIP_PATH, SKILL_NAME)

// 4. Report
const zipSize = statSync(ZIP_PATH).size
console.log(`✓ Skill bundle:  ${relative(ROOT, STAGE_DIR)}/`)
console.log(`✓ Skill archive: ${relative(ROOT, ZIP_PATH)} (${formatBytes(zipSize)})`)
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
