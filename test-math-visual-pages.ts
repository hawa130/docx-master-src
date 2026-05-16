#!/usr/bin/env bun
/**
 * Per-case visual rendering: one .docx → one .pdf → one .png per
 * fixture. Slower than the batched harness but gives a separate image
 * for each case so they can be inspected one at a time.
 */

import { readdir, readFile, writeFile, mkdir } from "node:fs/promises"
import { join, basename, dirname } from "node:path"
import { fileURLToPath } from "node:url"
import { spawnSync } from "node:child_process"
import JSZip from "jszip"

const HERE = dirname(fileURLToPath(import.meta.url))
const CASES_DIR = join(HERE, "test", "fixtures", "math", "cases")
const OUT_DIR = join(HERE, "test", "fixtures", "math", "_visual", "pages")

await mkdir(OUT_DIR, { recursive: true })

const onlyCase = process.argv[2] // optional: render single case
const files = (await readdir(CASES_DIR))
  .filter((f) => f.endsWith(".tex"))
  .filter((f) => !onlyCase || basename(f, ".tex") === onlyCase)
files.sort()

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

for (const f of files) {
  const name = basename(f, ".tex")
  const latex = (await readFile(join(CASES_DIR, f), "utf8")).trim()
  const omml = (await readFile(join(CASES_DIR, `${name}.expected.omml`), "utf8")).trim()
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:rFonts w:ascii="Menlo" w:hAnsi="Menlo"/><w:sz w:val="20"/><w:color w:val="666666"/></w:rPr><w:t xml:space="preserve">${escapeXml(name)}: ${escapeXml(latex)}</w:t></w:r>
    </w:p>
    <w:p>
      <m:oMathPara>${omml}</m:oMathPara>
    </w:p>
    <w:sectPr><w:pgSz w:w="8500" w:h="3500"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`
  const zip = new JSZip()
  zip.file("[Content_Types].xml", contentTypes)
  zip.file("_rels/.rels", rels)
  zip.file("word/document.xml", documentXml)
  const buf = await zip.generateAsync({ type: "nodebuffer" })
  const docxPath = join(OUT_DIR, `${name}.docx`)
  await writeFile(docxPath, buf)

  const r = spawnSync(
    "soffice",
    ["--headless", "--convert-to", "pdf", "--outdir", OUT_DIR, docxPath],
    { encoding: "utf8" },
  )
  if (r.status !== 0) {
    console.error(`soffice failed for ${name}:`, r.stderr)
    continue
  }
  const pdfPath = join(OUT_DIR, `${name}.pdf`)
  const pngPath = join(OUT_DIR, `${name}.png`)
  const s = spawnSync("sips", ["-s", "format", "png", pdfPath, "--out", pngPath], {
    encoding: "utf8",
  })
  if (s.status !== 0) {
    console.error(`sips failed for ${name}:`, s.stderr)
    continue
  }
  console.log(`✓ ${name}`)
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}
