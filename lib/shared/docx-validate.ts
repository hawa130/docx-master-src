/**
 * OOXML validator backed by xmllint-wasm (libxml2 compiled to WebAssembly,
 * the same engine the Anthropic docx skill uses via Python lxml). Provides
 * full ECMA-376 XSD validation — element ordering, types, enumerations,
 * minOccurs/maxOccurs, attribute constraints — without Python or native
 * deps. Schemas are bundled under `<scripts>/_shared/schemas/` at build
 * time.
 *
 * Two layers:
 *   - **XSD validation** (`validateXsd`): per-part schema check via
 *     xmllint-wasm. Catches everything the schema covers.
 *   - **Cross-part / package-level checks** (`checkCrossPartRefs`,
 *     `checkContentTypes`, `checkRelationships`, `checkMultiLevelType`):
 *     reference integrity and Word-specific quirks XSD doesn't enforce.
 *     Examples: a `<w:numId>` referenced from document.xml must exist in
 *     numbering.xml; a multi-level abstractNum without `<w:multiLevelType>`
 *     silently fails to render in Word even though XSD permits its absence.
 *
 * Public API:
 *   - `validateDocxFile(path)` — async, opens a docx and validates every
 *     part. Used by `apply` post-write and the standalone `validate` CLI.
 */

import { existsSync, readFileSync, readdirSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"
import type { XMLFileInfo, validateXML as ValidateXMLFn } from "xmllint-wasm"
import { DocxReader, parseXml } from "@lib/xml/reader.ts"
import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, getChildrenNS, wAttr } from "@lib/xml/xml-utils.ts"

export type ValidationError = { part: string; message: string }

const W = NS.w
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

/* ============================================================
 * SCHEMA LOADING
 * ========================================================== */

const MODULE_DIR = dirname(fileURLToPath(import.meta.url))

function findSchemasDir(): string {
  // Built bundle: `<scripts>/_shared/schemas` (sibling of cli-helpers.js).
  // Dev (running from lib/shared/): `<repo>/vendor/ooxml-schemas`.
  const candidates = [
    join(MODULE_DIR, "schemas"),
    join(MODULE_DIR, "..", "..", "vendor", "ooxml-schemas"),
  ]
  for (const p of candidates) {
    if (existsSync(join(p, "ISO-IEC29500-4_2016", "wml.xsd"))) return p
  }
  throw new Error(
    `OOXML XSD schemas not found. Looked in: ${candidates.join(", ")}. ` +
      `Run \`bun run build:skill\` to stage them under dist/.`,
  )
}

/** xmllint-wasm can't be statically bundled by tsdown — its package entry
 * does runtime `require("./xmllint-node.js")` and references `xmllint.wasm`
 * by relative path. We mark it `external` in tsdown.config and copy the
 * package's runtime files into the bundled output under `_shared/
 * xmllint-wasm/`. At runtime we try the bare specifier first (works in dev
 * via node_modules) and fall back to the colocated path (built bundle). */
let cachedValidateXML: typeof ValidateXMLFn | null = null
async function getValidateXML(): Promise<typeof ValidateXMLFn> {
  if (cachedValidateXML) return cachedValidateXML
  // Bare specifier resolves through node_modules in dev. The colocated path
  // resolves the runtime files copied alongside cli-helpers.js by build-skill.
  const candidates = [
    "xmllint-wasm",
    pathToFileURL(join(MODULE_DIR, "xmllint-wasm", "index-node.js")).href,
  ]
  let lastErr: unknown = null
  for (const spec of candidates) {
    try {
      const mod = (await import(spec)) as { validateXML: typeof ValidateXMLFn }
      cachedValidateXML = mod.validateXML
      return cachedValidateXML
    } catch (err) {
      lastErr = err
    }
  }
  throw new Error(
    `xmllint-wasm runtime not found. Looked for: ${candidates.join(", ")}. ` +
      `Last error: ${(lastErr as Error)?.message ?? "(unknown)"}.`,
  )
}

let cachedSchemas: XMLFileInfo[] | null = null
/**
 * Loaded XSDs are placed at flat paths (basename only) in xmllint-wasm's
 * in-memory FS, because Emscripten's MEMFS doesn't auto-create intermediate
 * directories on file write — nested paths like `ecma/fouth-edition/x.xsd`
 * fail with `ErrnoError 44 (ENOENT)`. Each schema's `xs:import` /
 * `xs:include` `schemaLocation` is rewritten to its basename to keep the
 * cross-schema reference graph intact under flattening.
 *
 * All 39 schemas have unique basenames so flattening is collision-free.
 */
function getSchemas(): XMLFileInfo[] {
  if (cachedSchemas) return cachedSchemas
  const root = findSchemasDir()
  const out: XMLFileInfo[] = []
  const walk = (dir: string) => {
    for (const ent of readdirSync(dir, { withFileTypes: true })) {
      const p = join(dir, ent.name)
      if (ent.isDirectory()) walk(p)
      else if (ent.name.endsWith(".xsd")) {
        let contents = readFileSync(p, "utf-8")
        // Flatten xs:import schemaLocation paths to basename (matches MEMFS).
        contents = contents.replace(/schemaLocation="([^"]+)"/g, (_, loc) => {
          const base = loc.split("/").pop() ?? loc
          return `schemaLocation="${base}"`
        })
        // Some ECMA schemas (e.g. opc-coreProperties.xsd) import the XML
        // namespace without a schemaLocation, expecting the validator to
        // resolve it from a built-in. lxml/Python does; xmllint-wasm
        // doesn't. Inject xml.xsd as the explicit location.
        contents = contents.replace(
          /(<xs(?:d)?:import\b[^>]*?\bnamespace="http:\/\/www\.w3\.org\/XML\/1998\/namespace")(?![^>]*\bschemaLocation=)([^>]*?\/>)/g,
          `$1 schemaLocation="xml.xsd"$2`,
        )
        out.push({ fileName: ent.name, contents })
      }
    }
  }
  walk(root)
  cachedSchemas = out
  return out
}

/**
 * Pick the main XSD to validate a given part against. Schemas are flattened
 * to basename in MEMFS, so the picker returns basenames too.
 */
function pickSchema(partPath: string): string | null {
  if (partPath === "[Content_Types].xml") return "opc-contentTypes.xsd"
  if (partPath.endsWith(".rels")) return "opc-relationships.xsd"
  if (partPath.startsWith("docProps/")) {
    if (partPath.endsWith("/app.xml")) return "shared-documentPropertiesExtended.xsd"
    if (partPath.endsWith("/core.xml")) return "opc-coreProperties.xsd"
    if (partPath.endsWith("/custom.xml")) return "shared-documentPropertiesCustom.xsd"
    return null
  }
  if (partPath.startsWith("word/")) {
    if (partPath.includes("/theme/")) return "dml-main.xsd"
    // Microsoft-extension parts (word/people.xml, commentsExtended.xml,
    // commentsIds.xml, commentsExtensible.xml, etc.) live in MS namespaces
    // whose XSDs we don't bundle (they're outside ECMA-376). Skip — these
    // parts aren't anything the engine mutates and their well-formedness
    // is implicitly checked when xmllint compiles its inputs.
    if (
      partPath.endsWith("/people.xml") ||
      /\/comments(Extended|Ids|Extensible)\.xml$/.test(partPath)
    )
      return null
    return "wml.xsd"
  }
  return null
}

/* ============================================================
 * XSD VALIDATION
 * ========================================================== */

/**
 * Errors libxml2 emits on every OOXML doc because the schemas don't cover
 * a few late-added Microsoft extensions and Dublin Core terms. The official
 * Anthropic validator filters the same list — see scripts/office/validators/
 * base.py:IGNORED_VALIDATION_ERRORS.
 */
const IGNORED_XSD_ERROR_PATTERNS = [/hyphenationZone/i, /purl\.org\/dc\//i]

/**
 * Namespaces the ECMA-376 core schemas know about. Attributes / elements in
 * any other namespace (Microsoft extensions w14 / w15 / w16, mc:Ignorable,
 * etc.) are stripped before validation — same approach as the official
 * Anthropic validator's `_clean_ignorable_namespaces`. Without this, a
 * vanilla Word doc emits hundreds of "attribute not allowed" errors against
 * the core schemas because Word adds the extension attrs by default.
 */
const OOXML_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/math",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
  "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://schemas.openxmlformats.org/presentationml/2006/main",
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
  "http://www.w3.org/XML/1998/namespace",
])

/**
 * Strip MC-extension attributes and elements (anything whose namespace
 * isn't in OOXML_NAMESPACES) from the XML before validating. Returns the
 * cleaned string. Operates via regex on the serialized XML — full DOM
 * round-trip is too slow when called per validation across many parts.
 */
function stripMcExtensions(xml: string): string {
  // Drop attributes whose qualified name uses a non-OOXML prefix. Identify
  // the prefixes used in the doc by scanning xmlns: declarations on the
  // root element, then strip attributes / elements with those prefixes.
  // This is approximate — a perfect implementation would parse + walk —
  // but matches what the official validator does in spirit.
  const xmlnsDecls = [...xml.matchAll(/xmlns:([\w-]+)="([^"]+)"/g)]
  const extPrefixes = xmlnsDecls
    .filter(([, , ns]) => !OOXML_NAMESPACES.has(ns ?? ""))
    .map(([, prefix]) => prefix)
  if (extPrefixes.length === 0) return xml
  let out = xml
  for (const prefix of extPrefixes) {
    // Strip attributes:  prefix:name="..." or prefix:name='...'
    out = out.replace(new RegExp(`\\s${prefix}:[\\w-]+="[^"]*"`, "g"), "")
    out = out.replace(new RegExp(`\\s${prefix}:[\\w-]+='[^']*'`, "g"), "")
    // Strip empty elements:  <prefix:name ... />
    out = out.replace(new RegExp(`<${prefix}:[\\w-]+(\\s[^>]*)?/>`, "g"), "")
    // Strip element pairs:  <prefix:name ...>...</prefix:name>
    out = out.replace(
      new RegExp(`<${prefix}:[\\w-]+(\\s[^>]*)?>[\\s\\S]*?</${prefix}:[\\w-]+>`, "g"),
      "",
    )
  }
  // Drop mc:Ignorable on any element (it lives in the mc namespace and
  // some XSDs don't tolerate it even if the prefix matches OOXML).
  out = out.replace(/\s\w+:Ignorable="[^"]*"/g, "")
  return out
}

async function validateXsd(parts: Map<string, string>): Promise<ValidationError[]> {
  const errors: ValidationError[] = []
  const schemas = getSchemas()
  const schemaMap = new Map(schemas.map((s) => [s.fileName, s]))

  // Group parts by main schema to amortize WASM startup across files.
  const groups = new Map<string, Array<{ partPath: string; contents: string }>>()
  for (const [partPath, contents] of parts) {
    const schemaPath = pickSchema(partPath)
    if (!schemaPath) continue
    if (!schemaMap.has(schemaPath)) continue
    if (!groups.has(schemaPath)) groups.set(schemaPath, [])
    groups.get(schemaPath)!.push({ partPath, contents })
  }

  for (const [schemaPath, files] of groups) {
    const mainSchema = schemaMap.get(schemaPath)!
    const preload = schemas.filter((s) => s.fileName !== schemaPath)
    // Flatten input fileNames into unique tokens for MEMFS (which can't
    // create intermediate directories). Map back to the original path when
    // reporting errors so the agent sees "word/styles.xml" not "input_0".
    const inputMap = new Map<string, string>()
    const xmlInputs = files.map((f, i) => {
      const token = `input_${i}.xml`
      inputMap.set(token, f.partPath)
      // MC stripping only applies to main-content parts (word/, ppt/, xl/);
      // docProps/.rels/[Content_Types].xml use cp:/dc:/vt: prefixes that
      // are valid in their own schemas. Same scope as the official
      // Anthropic validator's `_clean_ignorable_namespaces`.
      const inMainContent = /^(word|ppt|xl)\//.test(f.partPath)
      const contents = inMainContent ? stripMcExtensions(f.contents) : f.contents
      return { fileName: token, contents }
    })
    let result
    try {
      const validateXML = await getValidateXML()
      result = await validateXML({ xml: xmlInputs, schema: [mainSchema], preload })
    } catch (err) {
      const detail =
        err instanceof Error
          ? `${err.message}${err.stack ? `\n${err.stack.split("\n").slice(0, 3).join("\n")}` : ""}`
          : `(non-Error throw: ${typeof err}) ${JSON.stringify(err)}`
      errors.push({
        part: `(xmllint)`,
        message: `validator crashed for schema ${schemaPath}: ${detail}`,
      })
      continue
    }
    if (result.valid) continue
    for (const e of result.errors) {
      const text = e.message
      if (IGNORED_XSD_ERROR_PATTERNS.some((re) => re.test(text))) continue
      const reportedFile = e.loc?.fileName ?? ""
      const part = inputMap.get(reportedFile) ?? reportedFile ?? schemaPath
      errors.push({ part, message: text })
    }
  }
  return errors
}

/* ============================================================
 * CROSS-PART / PACKAGE-LEVEL CHECKS (XSD doesn't cover these)
 * ========================================================== */

/** Word built-in styleIds that resolve via docDefaults / latentStyles even
 * without an explicit `<w:style>` entry. */
const WORD_BUILTIN_STYLE_IDS = new Set(["Normal", "NoList", "TableNormal", "DefaultParagraphFont"])

/**
 * `<w:multiLevelType>` is optional per CT_AbstractNum schema, but Word
 * silently fails to render numbering on multi-level schemes that omit it.
 * XSD won't flag this — we have to.
 */
function checkMultiLevelType(numberingXml: string | null): ValidationError[] {
  if (!numberingXml) return []
  const doc = parseXmlOrNull(numberingXml)
  if (!doc?.documentElement) return []
  const errors: ValidationError[] = []
  for (const abs of getChildrenNS(doc.documentElement, W, "abstractNum")) {
    const aid = wAttr(abs, "abstractNumId") ?? "(unknown)"
    const levels = getChildrenNS(abs, W, "lvl")
    const mlt = firstChildNS(abs, W, "multiLevelType")
    if (!mlt && levels.length > 1) {
      errors.push({
        part: "word/numbering.xml",
        message: `abstractNum ${aid}: missing <w:multiLevelType> on a ${levels.length}-level scheme — Word silently fails to render numbering without it.`,
      })
    }
  }
  return errors
}

/**
 * Every `<w:numId>` referenced from document.xml must exist in
 * numbering.xml; every `<w:pStyle>` that numbering.xml binds must exist in
 * styles.xml (excepting Word built-ins). pStyle references in document.xml
 * itself are not flagged — Word falls back to Normal gracefully.
 */
function checkCrossPartRefs(parts: Map<string, string>): ValidationError[] {
  const errors: ValidationError[] = []
  const stylesDoc = parseXmlOrNull(parts.get("word/styles.xml") ?? null)
  const numberingDoc = parseXmlOrNull(parts.get("word/numbering.xml") ?? null)
  const documentDoc = parseXmlOrNull(parts.get("word/document.xml") ?? null)

  const knownStyleIds = new Set<string>()
  if (stylesDoc?.documentElement) {
    for (const s of getChildrenNS(stylesDoc.documentElement, W, "style")) {
      const sid = wAttr(s, "styleId")
      if (sid) knownStyleIds.add(sid)
    }
  }
  const knownNumIds = new Set<string>()
  if (numberingDoc?.documentElement) {
    for (const n of getChildrenNS(numberingDoc.documentElement, W, "num")) {
      const nid = wAttr(n, "numId")
      if (nid) knownNumIds.add(nid)
    }
  }

  if (documentDoc?.documentElement && knownNumIds.size > 0) {
    let pIdx = 0
    const visit = (el: Element) => {
      for (const child of getChildren(el)) {
        if (child.namespaceURI === W && child.localName === "p") pIdx += 1
        if (child.namespaceURI === W && child.localName === "numId") {
          const nid = wAttr(child, "val")
          if (nid && !knownNumIds.has(nid)) {
            errors.push({
              part: "word/document.xml",
              message: `paragraph #${pIdx}: numId="${nid}" doesn't exist in numbering.xml.`,
            })
          }
        }
        visit(child)
      }
    }
    visit(documentDoc.documentElement)
  }

  if (numberingDoc?.documentElement && knownStyleIds.size > 0) {
    for (const abs of getChildrenNS(numberingDoc.documentElement, W, "abstractNum")) {
      const aid = wAttr(abs, "abstractNumId")
      for (const lvl of getChildrenNS(abs, W, "lvl")) {
        const pStyle = firstChildNS(lvl, W, "pStyle")
        if (!pStyle) continue
        const sid = wAttr(pStyle, "val")
        if (sid && !knownStyleIds.has(sid) && !WORD_BUILTIN_STYLE_IDS.has(sid)) {
          errors.push({
            part: "word/numbering.xml",
            message: `abstractNum ${aid} lvl ${wAttr(lvl, "ilvl") ?? "?"}: pStyle references "${sid}", which doesn't exist in styles.xml — the style→numbering binding will silently fail.`,
          })
        }
      }
    }
  }
  return errors
}

/** Every part in the package must have either an Override or a Default
 * content type entry; every Override must point at an existing part. */
function checkContentTypes(
  parts: Map<string, string>,
  packageEntries: ReadonlyArray<string>,
): ValidationError[] {
  const errors: ValidationError[] = []
  const ct = parts.get("[Content_Types].xml")
  if (!ct) {
    return [{ part: "[Content_Types].xml", message: "missing" }]
  }
  const ctDoc = parseXmlOrNull(ct)
  if (!ctDoc?.documentElement) return [{ part: "[Content_Types].xml", message: "unparseable" }]

  const overrides = new Set<string>()
  for (const ov of getChildrenNS(ctDoc.documentElement, CT_NS, "Override")) {
    const partName = ov.getAttribute("PartName")
    if (partName) overrides.add(partName.replace(/^\//, ""))
  }
  const defaults = new Set<string>()
  for (const def of getChildrenNS(ctDoc.documentElement, CT_NS, "Default")) {
    const ext = def.getAttribute("Extension")
    if (ext) defaults.add(ext.toLowerCase())
  }

  const declared = (path: string) => {
    if (overrides.has(path)) return true
    const ext = path.split(".").pop()?.toLowerCase()
    return ext ? defaults.has(ext) : false
  }
  for (const path of packageEntries) {
    if (path.endsWith("/")) continue
    if (path === "[Content_Types].xml") continue
    if (!declared(path)) {
      errors.push({
        part: "[Content_Types].xml",
        message: `part "${path}" has neither Override nor Default content type — Word will reject the package.`,
      })
    }
  }
  for (const ov of overrides) {
    if (!packageEntries.includes(ov)) {
      errors.push({
        part: "[Content_Types].xml",
        message: `Override targets "/${ov}" but no such part exists in the package.`,
      })
    }
  }
  return errors
}

/** Every `r:id`/`r:embed`/`r:link` attribute used in an XML part must
 * resolve to a Relationship in the corresponding .rels file. */
function checkRelationships(
  parts: Map<string, string>,
  packageEntries: ReadonlyArray<string>,
): ValidationError[] {
  const errors: ValidationError[] = []
  const entrySet = new Set(packageEntries)

  for (const relsPath of packageEntries) {
    if (!relsPath.endsWith(".rels")) continue
    const text = parts.get(relsPath)
    if (!text) continue
    const relsDoc = parseXmlOrNull(text)
    if (!relsDoc?.documentElement) continue
    const ridSeen = new Set<string>()
    const ridToTargetMode = new Map<string, string>()
    const ridToTarget = new Map<string, string>()
    for (const rel of getChildrenNS(relsDoc.documentElement, PKG_REL_NS, "Relationship")) {
      const rid = rel.getAttribute("Id")
      const target = rel.getAttribute("Target") ?? ""
      const mode = rel.getAttribute("TargetMode") ?? "Internal"
      if (!rid) continue
      if (ridSeen.has(rid)) {
        errors.push({ part: relsPath, message: `duplicate relationship Id "${rid}".` })
      }
      ridSeen.add(rid)
      ridToTarget.set(rid, target)
      ridToTargetMode.set(rid, mode)
    }
    const baseDir =
      relsPath === "_rels/.rels"
        ? ""
        : relsPath.replace(/_rels\/[^/]+\.rels$/, "").replace(/\/$/, "")
    for (const [rid, target] of ridToTarget) {
      if (ridToTargetMode.get(rid) === "External") continue
      const resolved = resolveRelativePath(baseDir, target)
      if (!entrySet.has(resolved)) {
        errors.push({
          part: relsPath,
          message: `relationship "${rid}" Target "${target}" → "${resolved}" doesn't exist in the package.`,
        })
      }
    }
  }

  for (const xmlPath of packageEntries) {
    if (!xmlPath.endsWith(".xml")) continue
    const relsPath = relsForPart(xmlPath)
    if (!relsPath || !entrySet.has(relsPath)) continue
    const relsDoc = parseXmlOrNull(parts.get(relsPath) ?? null)
    if (!relsDoc?.documentElement) continue
    const knownRids = new Set<string>()
    for (const rel of getChildrenNS(relsDoc.documentElement, PKG_REL_NS, "Relationship")) {
      const rid = rel.getAttribute("Id")
      if (rid) knownRids.add(rid)
    }
    const xmlDoc = parseXmlOrNull(parts.get(xmlPath) ?? null)
    if (!xmlDoc?.documentElement) continue
    const visit = (el: Element) => {
      for (const attrName of ["id", "embed", "link"]) {
        const ridAttr = el.getAttributeNS?.(R_NS, attrName) ?? null
        if (ridAttr && !knownRids.has(ridAttr)) {
          errors.push({
            part: xmlPath,
            message: `<${el.localName}> r:${attrName}="${ridAttr}" references a non-existent relationship in ${relsPath}.`,
          })
        }
      }
      for (const child of getChildren(el)) visit(child)
    }
    visit(xmlDoc.documentElement)
  }
  return errors
}

/* ============================================================
 * ENTRY POINTS
 * ========================================================== */

async function validateOoxmlParts(
  parts: Map<string, string>,
  packageEntries?: ReadonlyArray<string>,
): Promise<ValidationError[]> {
  const errors: ValidationError[] = []
  errors.push(...(await validateXsd(parts)))
  errors.push(...checkMultiLevelType(parts.get("word/numbering.xml") ?? null))
  errors.push(...checkCrossPartRefs(parts))
  if (packageEntries) {
    errors.push(...checkContentTypes(parts, packageEntries))
    errors.push(...checkRelationships(parts, packageEntries))
  }
  return errors
}

export async function validateDocxFile(filePath: string): Promise<ValidationError[]> {
  const reader = await DocxReader.open(filePath)
  const entries = reader.listEntries()
  const parts = new Map<string, string>()
  for (const e of entries) {
    if (!e.endsWith(".xml") && !e.endsWith(".rels")) continue
    const text = await reader.readText(e)
    if (text !== null) parts.set(e, text)
  }
  return validateOoxmlParts(parts, entries)
}

/** Validate an `<m:oMath>...</m:oMath>` fragment against shared-math.xsd.
 *  Catches mml2omml's schema-invalid output (e.g. stray `<m:rPr>`) before
 *  it reaches document.xml. Returns one error per validator complaint;
 *  empty array = valid. */
export async function validateOMath(omml: string): Promise<string[]> {
  const schemas = getSchemas()
  const main = schemas.find((s) => s.fileName === "shared-math.xsd")
  if (!main) return [`shared-math.xsd not found in bundled schemas`]
  const preload = schemas.filter((s) => s.fileName !== "shared-math.xsd")
  let result
  try {
    const validateXML = await getValidateXML()
    result = await validateXML({
      xml: [{ fileName: "math.xml", contents: omml }],
      schema: [main],
      preload,
    })
  } catch (err) {
    return [`validator crashed: ${(err as Error)?.message ?? String(err)}`]
  }
  if (result.valid) return []
  const out: string[] = []
  for (const e of result.errors) {
    if (IGNORED_XSD_ERROR_PATTERNS.some((re) => re.test(e.message))) continue
    out.push(e.message)
  }
  return out
}

/* ============================================================
 * helpers
 * ========================================================== */

function relsForPart(partPath: string): string | null {
  if (partPath.endsWith(".rels")) return null
  const slash = partPath.lastIndexOf("/")
  if (slash < 0) return `_rels/${partPath}.rels`
  const dir = partPath.slice(0, slash)
  const file = partPath.slice(slash + 1)
  return `${dir}/_rels/${file}.rels`
}

function resolveRelativePath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1)
  const segs = (baseDir ? `${baseDir}/${target}` : target).split("/")
  const out: string[] = []
  for (const s of segs) {
    if (s === "." || s === "") continue
    if (s === "..") out.pop()
    else out.push(s)
  }
  return out.join("/")
}

function parseXmlOrNull(text: string | null): Document | null {
  if (!text) return null
  try {
    return parseXml(text)
  } catch {
    return null
  }
}
