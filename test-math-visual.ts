#!/usr/bin/env bun
/**
 * Visual smoke test: pack every fixture's OMML into one .docx, render
 * to PDF via headless LibreOffice, then to PNG via sips (macOS).
 *
 * Output: test/fixtures/math/_visual/all.{docx,pdf,png}.
 * Manual step: open the PNG and confirm each equation renders the way
 * the LaTeX intends.
 */

import { readdir, readFile, writeFile, mkdir } from "node:fs/promises"
import { join, basename, dirname } from "node:path"
import { fileURLToPath } from "node:url"
import { spawnSync } from "node:child_process"
import JSZip from "jszip"

const HERE = dirname(fileURLToPath(import.meta.url))
const CASES_DIR = join(HERE, "test", "fixtures", "math", "cases")
const OUT_DIR = join(HERE, "test", "fixtures", "math", "_visual")

await mkdir(OUT_DIR, { recursive: true })

const files = (await readdir(CASES_DIR)).filter((f) => f.endsWith(".tex"))
files.sort()

const bodyParts: string[] = []
for (const f of files) {
  const name = basename(f, ".tex")
  const latex = (await readFile(join(CASES_DIR, f), "utf8")).trim()
  const ommlPath = join(CASES_DIR, `${name}.expected.omml`)
  const omml = (await readFile(ommlPath, "utf8")).trim()
  // Label + LaTeX source + the OMML wrapped in oMathPara for display.
  bodyParts.push(`<w:p>
  <w:pPr><w:pStyle w:val="Heading4"/></w:pPr>
  <w:r><w:t xml:space="preserve">${escapeXml(name)}</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:rPr><w:rFonts w:ascii="Menlo" w:hAnsi="Menlo"/><w:sz w:val="18"/><w:color w:val="666666"/></w:rPr><w:t xml:space="preserve">${escapeXml(latex)}</w:t></w:r>
</w:p>
<w:p>
  <m:oMathPara>${omml}</m:oMathPara>
</w:p>`)
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
${bodyParts.join("\n")}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`

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

const zip = new JSZip()
zip.file("[Content_Types].xml", contentTypes)
zip.file("_rels/.rels", rels)
zip.file("word/document.xml", documentXml)
const buf = await zip.generateAsync({ type: "nodebuffer" })
const docxPath = join(OUT_DIR, "all.docx")
await writeFile(docxPath, buf)
console.log(`wrote ${docxPath} (${buf.length}B, ${files.length} cases)`)

console.log("running soffice (this takes ~5s)...")
const soffice = spawnSync(
  "soffice",
  ["--headless", "--convert-to", "pdf", "--outdir", OUT_DIR, docxPath],
  { encoding: "utf8" },
)
if (soffice.status !== 0) {
  console.error("soffice failed:", soffice.stderr)
  process.exit(1)
}
console.log(soffice.stdout.trim())

const pdfPath = join(OUT_DIR, "all.pdf")
const pngPath = join(OUT_DIR, "all.png")
console.log(`rendering ${pdfPath} → ${pngPath}...`)
const sips = spawnSync("sips", ["-s", "format", "png", pdfPath, "--out", pngPath], {
  encoding: "utf8",
})
if (sips.status !== 0) {
  console.error("sips failed:", sips.stderr)
  process.exit(1)
}
console.log(`done — open ${pngPath} to review`)

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}
