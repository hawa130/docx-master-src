/**
 * Builds the bundled blank.docx template used when an `apply` config omits
 * `source`. One-shot — run once, commit the output, re-run only when the
 * baseline needs adjustment (e.g. font default change, A4→Letter shift,
 * new fundamental style addition).
 *
 * Source theme1.xml is borrowed from the pristine Chinese-Word fixture at
 * `test/fixtures/空白Word文档.docx` because Word's default theme is a stable
 * artifact and hand-authoring all the colorScheme / fontScheme / fmtScheme
 * XML is no win. Every other part is hand-authored to keep it minimal:
 * no rsid noise, no mc:Ignorable for absent features, Latin built-in
 * styleIds (Normal / DefaultParagraphFont / TableNormal / NoList) so HF
 * style injection doesn't collide with locale aliases the way it would
 * against zh-CN Word's `a` / `a4` / `a5` IDs.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { dirname } from "node:path"
import JSZip from "jszip"

const SOURCE_FIXTURE = "test/fixtures/空白Word文档.docx"
const OUTPUT = "lib/apply/_assets/blank.docx"

const CONTENT_TYPES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
  `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
  `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
  `</Types>`

const PKG_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<w:body>` +
  `<w:p/>` +
  // A4 portrait, 2.54cm vertical / 3.17cm horizontal margins (Word's "Normal"
  // preset). pgMar units are twips; pgSz 11906×16838 = A4 portrait. Header /
  // footer offsets carried so HF mutation works without rewriting pgMar.
  `<w:sectPr>` +
  `<w:pgSz w:w="11906" w:h="16838"/>` +
  `<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>` +
  `<w:cols w:space="425"/>` +
  `<w:docGrid w:type="lines" w:linePitch="312"/>` +
  `</w:sectPr>` +
  `</w:body>` +
  `</w:document>`

const DOCUMENT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
  `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
  `</Relationships>`

// docDefaults match Word's modern default: 11pt Calibri (Latin) / 宋体 (CJK),
// single line spacing, no paragraph spacing. Agents declaring `theme.fonts`
// retarget the ASCII / hAnsi / EastAsia slots; agents declaring per-style
// fontLatin / fontCJK override at the style level. Either way the cascade
// works because every rFont here is a direct value (not a theme reference),
// which keeps the baseline self-contained.
const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:docDefaults>` +
  `<w:rPrDefault>` +
  `<w:rPr>` +
  `<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>` +
  `<w:kern w:val="2"/>` +
  `<w:sz w:val="22"/>` +
  `<w:szCs w:val="22"/>` +
  `<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>` +
  `</w:rPr>` +
  `</w:rPrDefault>` +
  `<w:pPrDefault/>` +
  `</w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
  `<w:name w:val="Normal"/>` +
  `<w:qFormat/>` +
  `</w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">` +
  `<w:name w:val="Default Paragraph Font"/>` +
  `<w:uiPriority w:val="1"/>` +
  `<w:semiHidden/>` +
  `<w:unhideWhenUsed/>` +
  `</w:style>` +
  `<w:style w:type="table" w:default="1" w:styleId="TableNormal">` +
  `<w:name w:val="Normal Table"/>` +
  `<w:uiPriority w:val="99"/>` +
  `<w:semiHidden/>` +
  `<w:unhideWhenUsed/>` +
  `<w:tblPr>` +
  `<w:tblInd w:w="0" w:type="dxa"/>` +
  `<w:tblCellMar>` +
  `<w:top w:w="0" w:type="dxa"/>` +
  `<w:left w:w="108" w:type="dxa"/>` +
  `<w:bottom w:w="0" w:type="dxa"/>` +
  `<w:right w:w="108" w:type="dxa"/>` +
  `</w:tblCellMar>` +
  `</w:tblPr>` +
  `</w:style>` +
  `<w:style w:type="numbering" w:default="1" w:styleId="NoList">` +
  `<w:name w:val="No List"/>` +
  `<w:uiPriority w:val="99"/>` +
  `<w:semiHidden/>` +
  `<w:unhideWhenUsed/>` +
  `</w:style>` +
  `</w:styles>`

// Empty settings — engine injects <w:updateFields> / <w:evenAndOddHeaders>
// on demand. The fabricate path in settings-mutation handles this case
// (the post-merge MINIMAL_SETTINGS_XML stub matches what we ship here).
const SETTINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`

async function main() {
  if (!existsSync(SOURCE_FIXTURE)) {
    throw new Error(`source fixture missing: ${SOURCE_FIXTURE}`)
  }
  // Borrow Word's default theme from the source fixture — stable artifact,
  // hand-authoring not worth it.
  const sourceZip = await JSZip.loadAsync(readFileSync(SOURCE_FIXTURE))
  const themeEntry = sourceZip.file("word/theme/theme1.xml")
  if (!themeEntry) throw new Error("source fixture lacks word/theme/theme1.xml")
  const themeXml = await themeEntry.async("string")

  const out = new JSZip()
  out.file("[Content_Types].xml", CONTENT_TYPES)
  out.file("_rels/.rels", PKG_RELS)
  out.file("word/document.xml", DOCUMENT_XML)
  out.file("word/_rels/document.xml.rels", DOCUMENT_RELS)
  out.file("word/styles.xml", STYLES_XML)
  out.file("word/settings.xml", SETTINGS_XML)
  out.file("word/theme/theme1.xml", themeXml)

  const buffer = await out.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  })

  mkdirSync(dirname(OUTPUT), { recursive: true })
  writeFileSync(OUTPUT, buffer)
  console.log(`wrote ${OUTPUT} (${buffer.length} bytes)`)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
