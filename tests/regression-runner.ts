#!/usr/bin/env bun
/**
 * Golden-output regression runner.
 *
 * For each (fixture, tool, args) triple in CASES below:
 *   1. Run the tool's compiled JS in dist/ via Bun.spawnSync (no shell).
 *   2. Compare stdout to tests/expected/<case-id>.txt.
 *   3. Print pass / fail per case, exit code 1 if any failure.
 *
 * Refresh the expected files after an intentional change:
 *   bun run test:regression:update
 *
 * Bun.spawnSync is used (not child_process.execSync) because it accepts
 * argv as a string[] — no shell interpolation, no injection surface,
 * even if fixture paths contain spaces or shell metacharacters.
 */

import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs"
import { join, dirname, resolve } from "node:path"

interface Case {
  id: string
  tool: string
  fixture: string
  args: string[]
}

const FIXTURE_DIR_LOCAL = "_local/fixtures"
const FIXTURE_DIR_REPORT = "/Users/geek-tech/hawa130/report/烟草局/fill-project"

const CASES: Case[] = [
  { id: "overview-table-showcase", tool: "overview", fixture: `${FIXTURE_DIR_LOCAL}/_filled_table_showcase.docx`, args: [] },
  { id: "overview-academic-demo", tool: "overview", fixture: `${FIXTURE_DIR_LOCAL}/academic-demo.docx`, args: [] },
  { id: "overview-filled-proposal", tool: "overview", fixture: `${FIXTURE_DIR_LOCAL}/_filled_proposal.docx`, args: [] },
  { id: "overview-table-padding", tool: "overview", fixture: `${FIXTURE_DIR_LOCAL}/table-padding-showcase.docx`, args: [] },
  { id: "overview-empty-thesis-proposal", tool: "overview", fixture: `${FIXTURE_DIR_LOCAL}/空开题报告表.docx`, args: [] },
  { id: "overview-tobacco-template", tool: "overview", fixture: `${FIXTURE_DIR_REPORT}/申报书模板.docx`, args: [] },
  { id: "overview-tobacco-filled", tool: "overview", fixture: `${FIXTURE_DIR_REPORT}/基于参数高效微调的烟草领域知识持续学习模型研究与应用.docx`, args: [] },
]

const EXPECTED_DIR = "tests/expected"
const UPDATE = process.argv.includes("--update")
const SCRIPT_DIR = "dist/plugin/skills/docx-master/scripts"

async function runCase(c: Case): Promise<string> {
  const scriptPath = resolve(`${SCRIPT_DIR}/${c.tool}.js`)
  if (!existsSync(scriptPath)) {
    return `(missing ${scriptPath} — run \`bun run build:skill\` first)`
  }
  const proc = Bun.spawnSync({
    cmd: ["node", scriptPath, c.fixture, ...c.args],
    stdout: "pipe",
    stderr: "pipe",
  })
  const stdout = new TextDecoder().decode(proc.stdout)
  const stderr = new TextDecoder().decode(proc.stderr)
  return stdout + (stderr ? `\n[stderr]\n${stderr}` : "")
}

function compare(c: Case, actual: string): boolean {
  const expectedPath = join(EXPECTED_DIR, `${c.id}.txt`)
  if (UPDATE) {
    mkdirSync(dirname(expectedPath), { recursive: true })
    writeFileSync(expectedPath, actual)
    console.log(`  UPDATED ${c.id}`)
    return true
  }
  if (!existsSync(expectedPath)) {
    console.error(`  MISSING ${c.id}: ${expectedPath} not found (run with --update to create)`)
    return false
  }
  const expected = readFileSync(expectedPath, "utf8")
  if (expected === actual) {
    console.log(`  OK  ${c.id}`)
    return true
  }
  console.error(`  FAIL ${c.id}: output differs`)
  const aLines = actual.split("\n")
  const eLines = expected.split("\n")
  const max = Math.max(aLines.length, eLines.length)
  let printed = 0
  for (let i = 0; i < max && printed < 20; i++) {
    if (aLines[i] !== eLines[i]) {
      console.error(`    line ${i + 1}: expected: ${eLines[i] ?? "<EOF>"}`)
      console.error(`    line ${i + 1}: actual:   ${aLines[i] ?? "<EOF>"}`)
      printed++
    }
  }
  return false
}

console.log(`Running ${CASES.length} regression case(s)${UPDATE ? " (UPDATE MODE)" : ""}...`)
let failed = 0
for (const c of CASES) {
  const stdout = await runCase(c)
  if (!compare(c, stdout)) failed++
}
console.log(`\n${CASES.length - failed} passed, ${failed} failed.`)
process.exit(failed > 0 ? 1 : 0)
