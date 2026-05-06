#!/usr/bin/env bun
/**
 * Stage SKILL.md + references/ from src/skills/<name>/ into dist/<name>/
 * alongside tsdown's scripts/, then zip into dist/<name>.zip.
 *
 * Multi-skill aware: pass a skill name as the first arg, or omit to build the
 * single skill present under src/skills/. With multiple skills present,
 * --all builds them in sequence.
 *
 * Run AFTER `tsdown` (which produces dist/<skill>/scripts/) — `bun run
 * build:skill` chains both.
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
const SKILLS_DIR = join(ROOT, "src", "skills")
const DIST_DIR = join(ROOT, "dist")

const args = process.argv.slice(2)
const wantAll = args.includes("--all")
const explicitName = args.find((a) => !a.startsWith("--"))

const allSkills = listSkills(SKILLS_DIR)
if (allSkills.length === 0) {
  console.error(`no skills found under ${relative(ROOT, SKILLS_DIR)}/`)
  process.exit(1)
}

let toBuild: string[]
if (wantAll) {
  toBuild = allSkills
} else if (explicitName) {
  if (!allSkills.includes(explicitName)) {
    console.error(
      `skill "${explicitName}" not found under ${relative(ROOT, SKILLS_DIR)}/.\n` +
        `  available: [${allSkills.join(", ")}]`,
    )
    process.exit(1)
  }
  toBuild = [explicitName]
} else if (allSkills.length === 1) {
  toBuild = allSkills
} else {
  console.error(
    `multiple skills present — pick one or pass --all.\n` +
      `  available: [${allSkills.join(", ")}]\n` +
      `  usage: bun run build-skill.ts <name>  |  bun run build-skill.ts --all`,
  )
  process.exit(1)
}

for (const skill of toBuild) {
  await buildOne(skill)
}

async function buildOne(skillName: string) {
  const skillSrc = join(SKILLS_DIR, skillName)
  const stageDir = join(DIST_DIR, skillName)
  const zipPath = join(DIST_DIR, `${skillName}.zip`)
  const scriptsDir = join(stageDir, "scripts")

  if (!existsSync(scriptsDir)) {
    console.error(
      `${relative(ROOT, scriptsDir)}/ not found — run \`bun run build\` (tsdown) before build-skill.`,
    )
    process.exit(1)
  }

  // 1. Copy SKILL.md and references/ from the source skill dir
  const skillMdSrc = join(skillSrc, "SKILL.md")
  if (!existsSync(skillMdSrc)) {
    throw new Error(`missing SKILL.md at ${relative(ROOT, skillMdSrc)}`)
  }
  cpSync(skillMdSrc, join(stageDir, "SKILL.md"))

  const refsSrc = join(skillSrc, "references")
  const refsDst = join(stageDir, "references")
  rmSync(refsDst, { recursive: true, force: true })
  if (existsSync(refsSrc)) {
    cpSync(refsSrc, refsDst, { recursive: true })
  }

  // 2. Zip the staged dir
  await zipDir(stageDir, zipPath, skillName)

  // 3. Report
  const zipSize = statSync(zipPath).size
  console.log(`✓ Skill bundle:  ${relative(ROOT, stageDir)}/`)
  console.log(`✓ Skill archive: ${relative(ROOT, zipPath)} (${formatBytes(zipSize)})`)
  console.log("")
  console.log("Contents:")
  listTree(stageDir, "  ")
  console.log("")
}

/* ---------- helpers ---------- */

function listSkills(dir: string): string[] {
  if (!existsSync(dir)) return []
  return readdirSync(dir, { withFileTypes: true })
    .filter((e) => e.isDirectory())
    .filter((e) => existsSync(join(dir, e.name, "SKILL.md")))
    .map((e) => e.name)
    .sort()
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
