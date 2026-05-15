/**
 * Word built-in style names that Word translates between locales.
 *
 * `<w:name w:val="heading 1"/>` is the *English canonical* form. zh-CN
 * Word UI renders the same style as "标题 1"; German as "Überschrift 1";
 * etc. Word resolves field codes (STYLEREF, REF, etc.) against the
 * *locale-translated* form — so `STYLEREF "heading 1"` fails to match
 * in any non-EN Word, even though the underlying styles.xml carries
 * the English name.
 *
 * Two consumers in the apply pipeline care about this list:
 *   - style-name collision detection (apply-styles.ts preflight) treats
 *     locale-aliased pairs as the same built-in identity to avoid the
 *     "two style entries with the same effective name" failure
 *   - STYLEREF emit (`emitInlineStyleRef`) routes built-in heading
 *     targets through the locale-neutral `STYLEREF N` form, and refuses
 *     to emit `STYLEREF "<name>"` for built-in *non-heading* styles
 *     (Title / Caption / Subtitle / etc.) — that form silently fails
 *     in non-EN Word at render time
 *
 * Maintenance: the list of built-in names is fixed by Word's template
 * (`normal.dotm` ships ~250 styles, ~30 of which are routinely
 * referenced and locale-translated). zh-CN is the only locale we
 * currently track since it's the only one in production use; new
 * locale pairs are additive and don't require code changes.
 */

/** Built-in style English canonical name → zh-CN localized form. */
const ENGLISH_TO_ZH_CN: ReadonlyArray<readonly [string, string]> = [
  ["Normal", "正文"],
  ["heading 1", "标题 1"],
  ["heading 2", "标题 2"],
  ["heading 3", "标题 3"],
  ["heading 4", "标题 4"],
  ["heading 5", "标题 5"],
  ["heading 6", "标题 6"],
  ["heading 7", "标题 7"],
  ["heading 8", "标题 8"],
  ["heading 9", "标题 9"],
  ["Title", "标题"],
  ["Subtitle", "副标题"],
  ["Body Text", "正文文本"],
  ["Caption", "题注"],
  ["Quote", "引用"],
  ["List Bullet", "列表项目符号"],
  ["List Number", "列表编号"],
  ["Header", "页眉"],
  ["Footer", "页脚"],
  ["TOC 1", "目录 1"],
  ["TOC 2", "目录 2"],
  ["TOC 3", "目录 3"],
  ["Default Paragraph Font", "默认段落字体"],
  ["Normal Table", "普通表格"],
  ["No List", "无列表"],
]

/** All names (English canonical + every locale alias) that resolve to
 *  a Word built-in style. Used to detect when an agent-emitted STYLEREF
 *  argument would silently fail in a non-EN Word locale. */
const BUILT_IN_NAMES: ReadonlySet<string> = new Set(
  ENGLISH_TO_ZH_CN.flatMap(([eng, zh]) => [eng, zh]),
)

/** Canonical-name map: any locale form → English canonical form. Used
 *  by collision detection to treat "Normal" and "正文" as the same
 *  built-in identity. */
const TO_CANONICAL: ReadonlyMap<string, string> = new Map(
  ENGLISH_TO_ZH_CN.flatMap(([eng, zh]) => [
    [eng, eng] as const,
    [zh, eng] as const,
  ]),
)

/** Map a style name to its English canonical form when it's a known
 *  built-in identity; otherwise return the input unchanged. Used for
 *  comparing two names to see if they'd collide on Word's built-in
 *  identity. */
export function canonicalStyleName(name: string): string {
  return TO_CANONICAL.get(name) ?? name
}

/** Returns true when the name matches a built-in Word style whose
 *  display form differs across Word UI languages. Locale-aliased pairs
 *  match by either side (English "Title" or Chinese "标题"). */
export function isBuiltInLocalizable(name: string): boolean {
  return BUILT_IN_NAMES.has(name)
}
