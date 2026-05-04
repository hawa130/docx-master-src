#!/usr/bin/env bun
/**
 * Stage SKILL.md + references into dist/<skill>/ alongside tsdown's scripts/,
 * then zip into dist/<skill>.zip.
 *
 * Run AFTER `tsdown` — relies on dist/<skill>/scripts/ being up to date.
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
const SKILL_NAME = parseSkillName(join(ROOT, "SKILL.md")) || "docx-normalize"
const STAGE_DIR = join(ROOT, "dist", SKILL_NAME)
const ZIP_PATH = join(ROOT, "dist", `${SKILL_NAME}.zip`)
const SCRIPTS_DIR = join(STAGE_DIR, "scripts")

if (!existsSync(SCRIPTS_DIR)) {
  console.error(
    `${relative(ROOT, SCRIPTS_DIR)}/ not found — run \`bun run build\` (tsdown) before build-skill.`,
  )
  process.exit(1)
}

// 1. Copy SKILL.md and references/ into the staged dir (scripts/ already there)
copyFile(join(ROOT, "SKILL.md"), join(STAGE_DIR, "SKILL.md"))
const REFS_SRC = join(ROOT, "references")
const REFS_DST = join(STAGE_DIR, "references")
rmSync(REFS_DST, { recursive: true, force: true })
if (existsSync(REFS_SRC)) {
  cpSync(REFS_SRC, REFS_DST, { recursive: true })
}

// 2. Zip the staged dir
await zipDir(STAGE_DIR, ZIP_PATH, SKILL_NAME)

// 3. Report
const zipSize = statSync(ZIP_PATH).size
console.log(`✓ Skill bundle:  ${relative(ROOT, STAGE_DIR)}/`)
console.log(`✓ Skill archive: ${relative(ROOT, ZIP_PATH)} (${formatBytes(zipSize)})`)
console.log("")
console.log("Contents:")
listTree(STAGE_DIR, "  ")

/* ---------- helpers ---------- */

function parseSkillName(skillPath: string): string | null {
  if (!existsSync(skillPath)) return null
  const text = readFileSync(skillPath, "utf8")
  const m = text.match(/^---[\s\S]*?\bname:\s*([^\n]+?)\s*\n[\s\S]*?---/m)
  return m ? m[1]!.trim() : null
}

function copyFile(src: string, dst: string) {
  if (!existsSync(src)) throw new Error(`missing file: ${src}`)
  cpSync(src, dst)
}

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
