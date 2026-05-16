#!/usr/bin/env bun
/**
 * Regression harness for the in-tree MathML → OMML converter. For each
 * `.tex` in `test/fixtures/math/cases/`:
 *
 *   1. Render to MathML via temml.
 *   2. Convert MathML → OMML via the in-tree converter.
 *   3. Validate the OMML against `shared-math.xsd` (libxml2 / xmllint-wasm).
 *   4. If a sibling `<case>.expected.omml` exists, diff against it
 *      (whitespace-normalized).
 *
 * Exits non-zero on any schema error or expected-diff mismatch.
 * Word visual verification stays a manual step — see
 * `test/fixtures/math/README.md`.
 *
 * Why this is written as top-level await rather than `async function main()`:
 * Bun's worker_threads ↔ async-function-await interaction silently exits
 * the process when awaiting xmllint-wasm's worker promise from inside a
 * nested async function. Top-level await sidesteps it.
 */

import { readdir, readFile, writeFile } from "node:fs/promises"
import { existsSync } from "node:fs"
import { join, basename, dirname } from "node:path"
import { fileURLToPath } from "node:url"
import temml from "temml"
import { convertMathMLToOMML } from "@lib/edit/math/mml-to-omml/index.ts"
import { validateOMath } from "@lib/shared/docx-validate.ts"

const UPDATE = process.argv.includes("--update")

const HERE = dirname(fileURLToPath(import.meta.url))
const CASES_DIR = join(HERE, "test", "fixtures", "math", "cases")

const files = (await readdir(CASES_DIR)).filter((f) => f.endsWith(".tex"))
files.sort()
console.log(`Found ${files.length} cases in ${CASES_DIR}\n`)

interface Result {
  name: string
  ok: boolean
  notes: string[]
}

const results: Result[] = []
for (const f of files) {
  const name = basename(f, ".tex")
  const notes: string[] = []
  let ok = true
  try {
    const latex = (await readFile(join(CASES_DIR, f), "utf8")).trim()
    const mathml = temml.renderToString(latex, {
      xml: true,
      displayMode: true,
      throwOnError: true,
    })
    const omml = convertMathMLToOMML(mathml)
    const schemaErrors = await validateOMath(omml)
    if (schemaErrors.length > 0) {
      ok = false
      notes.push(
        `schema: ${schemaErrors[0]}${schemaErrors.length > 1 ? ` (+${schemaErrors.length - 1})` : ""}`,
      )
    }
    const expectedPath = join(CASES_DIR, `${name}.expected.omml`)
    if (UPDATE) {
      await writeFile(expectedPath, omml + "\n")
      notes.push("updated expected.omml")
    } else if (existsSync(expectedPath)) {
      const expected = (await readFile(expectedPath, "utf8")).trim()
      if (normalize(omml) !== normalize(expected)) {
        ok = false
        notes.push(`diff vs expected.omml`)
      } else {
        notes.push("matches expected")
      }
    } else {
      notes.push("schema-only check (no expected.omml)")
    }
  } catch (err) {
    ok = false
    notes.push(`threw: ${(err as Error).message.slice(0, 160)}`)
  }
  results.push({ name, ok, notes })
}

for (const r of results) {
  const tag = r.ok ? "PASS" : "FAIL"
  console.log(`  ${tag}  ${r.name.padEnd(22)}  ${r.notes.join("; ")}`)
}
const failed = results.filter((r) => !r.ok).length
console.log(`\n${results.length - failed}/${results.length} passed`)
if (failed > 0) process.exit(1)

function normalize(xml: string): string {
  return xml.replace(/>\s+</g, "><").trim()
}
