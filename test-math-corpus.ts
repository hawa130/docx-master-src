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
const ERRORS_DIR = join(HERE, "test", "fixtures", "math", "errors")

const files = (await readdir(CASES_DIR)).filter((f) => f.endsWith(".tex"))
files.sort()
const errorFiles = (await readdir(ERRORS_DIR)).filter((f) => f.endsWith(".mml"))
errorFiles.sort()
console.log(`Found ${files.length} cases + ${errorFiles.length} error cases\n`)

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

// Negative tests — each .mml in errors/ feeds the converter directly
// (not via temml, since malformed input would never come from temml
// in practice — these guard against MathML from other producers and
// unsupported elements). The error must mention the substring in
// <case>.expected-error.txt.
for (const f of errorFiles) {
  const name = "error:" + basename(f, ".mml")
  const notes: string[] = []
  let ok = false
  const mathml = (await readFile(join(ERRORS_DIR, f), "utf8")).trim()
  const expectedErrPath = join(ERRORS_DIR, `${basename(f, ".mml")}.expected-error.txt`)
  const expectedSubstring = existsSync(expectedErrPath)
    ? (await readFile(expectedErrPath, "utf8")).trim()
    : ""
  try {
    convertMathMLToOMML(mathml)
    notes.push(`did not throw (expected substring "${expectedSubstring.slice(0, 50)}")`)
  } catch (err) {
    const message = (err as Error).message
    if (expectedSubstring === "" || message.includes(expectedSubstring)) {
      ok = true
      notes.push(`threw with expected message`)
    } else {
      notes.push(
        `threw but message lacks "${expectedSubstring.slice(0, 50)}"; got: ${message.slice(0, 120)}`,
      )
    }
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
