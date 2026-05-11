# Standardize-shape `apply` config

Standardize adds **structure** (semantic styleId, outlineLevel, auto-numbering, sectioning). Typography stays where the template put it — touch font / size / alignment only when (a) the user prompt explicitly names them, or (b) chrome is inconsistent enough that no exemplar can be extracted. Reading "标准化 / 排版" as "reset typography to canonical values" is the most common agent failure on this skill.

How to compose an `apply` config that reshapes a Word document's style
system, numbering, and role assignments. Operates by **pattern**: you
describe categories of paragraphs (regex match / fingerprint match /
specific style), the engine applies uniformly to every member of the
category. Per-paragraph enumeration is the **last resort**, not the default.

## When to use this shape

- The user wants the doc brought into a consistent, standardized form (whole-doc reshape).
- The user wants a focused style-system change (add a Heading3, change Heading2 size, install numbering).
- A template needs structure work before content can be filled (typed chrome → auto-numbering, missing Heading levels, missing list-bound style).

For surgical changes at specific paragraph indices or table cells, see the
edits-shape config in [edit.md](edit.md). Both shapes are blocks of the same
`apply` config — combine them in one call when a task needs both.

## Iterating with --dry-run

The supported workflow:

1. `apply --dry-run <config.json>` — applies in-memory, prints the change report, doesn't touch the output file.
2. Read the report. Refine the config based on what hit / what didn't.
3. Repeat until correct, then `apply` (no flag) to write.

Each cycle is seconds. `styles[]` is sparse — declare only what you're touching; untouched styles stay as they are.

---

## Default workflow: pattern-driven config

For most whole-doc reshape tasks, design **one** `apply` config that does the entire job. Start with this skeleton and fill it in based on what the doc + content needs:

```jsonc
{
  "source": "input.docx",
  "output": "output.docx",

  // 1. styles[]: install one entry per semantic role the doc + content need.
  //    Mode A (preferred) extracts the definition from a representative
  //    paragraph already styled as this role — no fields invented. Mode B
  //    (explicit fields) is the fallback when no exemplar exists or the
  //    user prompt explicitly names typography. Locale defaults belong on
  //    fresh styles, not piled onto Mode A extractions.
  "styles": [
    // Mode A (preferred): replace 33/47/... with paragraph indices from
    // your overview that already look like this role. The engine pulls
    // font / size / weight / indent / spacing from the actual paragraph —
    // nothing gets hallucinated. outlineLevel is structural (binds the
    // style to Word's outline view), so it stays at top-level.
    { "id": "Heading1", "name": "heading 1", "fromParagraph": 33, "outlineLevel": 0 },
    { "id": "Heading2", "name": "heading 2", "fromParagraph": 47, "outlineLevel": 1 },
    { "id": "Heading3", "name": "heading 3", "fromParagraph": 58, "outlineLevel": 2 },
    { "id": "Heading4", "name": "heading 4", "fromParagraph": 71, "outlineLevel": 3 },

    // Mode B (fallback): explicit fields. Use when source has no
    // representative paragraph (empty template), OR the user prompt
    // explicitly names typography. Declare ONLY what the user spec asks
    // or what the empty slot needs (locale defaults).
    { "id": "ListNumber", "name": "List Number", "fontCJK": "宋体", "size": 12, "firstLineIndent": "2char" }
    // Add Caption / Quote / Code / etc. as the doc + content require.
  ],

  // 2. Numbering as ARRAY: one multi-level scheme for headings, separate
  //    single-level scheme for body lists. Install both in one pass.
  "numbering": [
    {
      "levels": [
        { "level": 0, "numFmt": "chineseCounting",  "lvlText": "%1、",     "suff": "nothing", "styleId": "Heading1" },
        { "level": 1, "numFmt": "chineseCounting",  "lvlText": "（%2）",    "suff": "nothing", "styleId": "Heading2" },
        { "level": 2, "numFmt": "decimal",          "lvlText": "%3.",      "suff": "space",   "styleId": "Heading3" },
        { "level": 3, "numFmt": "decimal",          "lvlText": "（%4）",    "suff": "nothing", "styleId": "Heading4" }
      ]
    },
    {
      "levels": [
        { "level": 0, "numFmt": "decimal", "lvlText": "%1.", "suff": "space", "styleId": "ListNumber" }
      ]
    }
  ],

  // 3. Pattern rules describe categories, engine applies uniformly.
  //    One rule per chrome shape the doc uses. Engine catches every match.
  //    `stripMatch: true` removes the manual prefix during restyle so
  //    auto-numbering takes over.
  "pattern_rules": [
    { "regex": "^[一二三四五六七八九十百千]+、",     "style": "Heading1", "stripMatch": true },
    { "regex": "^（[一二三四五六七八九十百千]+）",   "style": "Heading2", "stripMatch": true },
    { "regex": "^\\d+\\.\\d+\\s",                "style": "Heading3", "stripMatch": true }
    // Add patterns per your survey of the doc's actual chrome shapes.
  ],

  // 4. Exclude false positives the regex caught wrongly (rare).
  "exclude": [],

  // 5. Outliers — paragraphs that don't fit any pattern but need restyling.
  //    Use sparingly; if you find yourself listing > 5, your pattern is wrong.
  "assignments": []
}
```

**Targeting precedence (first match wins):** `exclude` > `assignments` > `pattern_rules` > `bulk_rules` > implicit-keep.

This is one config, one `apply` call, dry-run + apply. **Don't** enumerate chrome paragraphs into `assignments` — that's where round-by-round selective skipping happens. Patterns describe; assignments correct.

### Typical recipe shapes

The pattern_rules above are tuned to one common Chinese academic-form pattern. Other recipes:

- **Decimal-only hierarchy** (`1`, `1.1`, `1.1.1`): `pattern_rules` with `^\\d+\\.\\s` / `^\\d+\\.\\d+\\s` / `^\\d+\\.\\d+\\.\\d+\\s` for Heading1/2/3.
- **Chapter-sentinel + decimal**: `^第[一二三四...]+章` for Heading1, `^\\d+\\.\\d+\\s` for Heading2, etc.
- **English standard**: `^Chapter \\d+` / `^\\d+\\.\\s` / `^\\d+\\.\\d+\\s` etc.
- **Pre-styled chrome**: when chrome doesn't share a text pattern but does share a styleId/font, use `bulk_rules: [{ fingerprint, style }]` instead of `pattern_rules`. The fingerprint comes from the overview's visual summary.

`numbering-formats.md` has full templates for academic / technical / governmental / legal multi-level shapes.

### When the recipe doesn't fit

Some real cases the recipe needs adapting for:

- **Inconsistent chrome**: section 1 uses bare colon-labels (`Name:`), sections 2+ use enumerator chrome (`1.`, `2.`). Convert what's structural via `pattern_rules`; leave inline labels alone (they're not section headings).
- **Unstable manual prefixes**: source has `1.1` in some chapters, `1.` in others (per-chapter restart). Use `stripPrefixPatterns: ["%1.%2", "%1."]` on the relevant level — longer pattern first.
- **Chrome that shares a text shape with body content**: e.g., body paragraphs cite `第三章` as a reference. The pattern matches both; use `exclude` or refine the regex (anchor at paragraph start AND require trailing context).
- **Mixed content + chrome**: in a fill task, the markdown content may also have hand-typed numbering. `pattern_rules` catches both; `stripMatch` cleans both. Same mechanism — no separate path for content vs chrome.

---

## Designing the style system

**Source of values** (priority order): (1) user requirements, (2) template document via `template`, (3) values extracted from a representative paragraph via `fromParagraph`, (4) sensible defaults. Don't invent values when a representative paragraph exists.

**Two modes per `styles` entry:**

1. **`fromParagraph`** (preferred when extracting from the doc): pick the first occurrence of the dominant fingerprint for the role and set `fromParagraph: <index>`. The tool extracts the full computed rPr + pPr from that paragraph's *dominant text run* (longest non-numbering-prefix run, so `"1.1 研究方法"` extracts the title formatting, not the prefix's). Use `overrides` to add fields the source lacks (e.g. `outlineLevel`) or apply user-requested specifics.
2. **Manual mode**: specify fields directly — when no representative paragraph exists, when synthesizing a role, or when the user fully specified the style.

Modes can mix within one `styles` array.

**`fontLatin` vs `fontCJK`:** `fontLatin` is the Latin / Western text slot (writes to OOXML `w:ascii` and `w:hAnsi`); `fontCJK` is the East-Asian slot (`w:eastAsia`). When the user names only a CJK font ("正文宋体" / "标题黑体"), set `fontCJK` and leave `fontLatin` unset so the source's Latin font is preserved. Set both only when the user explicitly says the same font should apply to Latin too.

**Indent unit preservation:** when the source used Word's character-based indent (`w:firstLineChars` / `w:hangingChars`, what Word writes for "首行缩进 N 字符"), extraction gives `"Nchar"` so font-size auto-scaling round-trips. Fixed twips give `"Npt"`. Don't manually convert "char" values to pt — that locks the indent to one font size.

**Override existing styles before creating new ones.** Run `inspect_style_def` to discover what the source already has — POI / WPS / school templates often play the role of Normal / Heading 1 / etc. under short auto-generated styleIds (`a`, `a1`, `2`, `10`, ...). Override by their exact styleId: upsertStyle mutates in place, preserving everything you didn't specify (basedOn, default="1", numPr, link, etc.). Verify the style is actually used for its intended role first — overriding `Heading1` while it's misused as body text would corrupt those paragraphs; reassign the paragraphs first.

**Override sparsely.** Declare only the fields the user's spec explicitly requires. Mode A `fromParagraph` and locale-default backfills (CJK 2-char indent, etc.) belong on *fresh* styles — piled onto an existing source definition they silently rewrite fields the user didn't ask to change. The template author's existing values stay where the user spec doesn't override them.

**Exception: chaotic source.** When source styleIds don't separate roles — headings, body, captions all bound to the same `a` — override can't help; changing `a` would shift all of them at once. Install fresh semantic styles (`Heading1` / `BodyText` / ...) and route paragraphs to them via `pattern_rules` / `assignments`.

**`name` must not alias any existing style's identity.** Word treats `<w:name>` as the built-in style identity marker, including locale aliases ("Normal" ≡ "正文" ≡ "標準"; "Heading 1" ≡ "标题 1"; "Body Text" ≡ "正文文本"). When two different styleIds claim the same identity, Word silently drops the second style's `rPr` at render time. Three safe approaches:

- **Override existing by its styleId** so no new name enters the doc.
- **When creating new with a built-in styleId** (`Heading1`, `BodyText`, `Caption`, ...), use the canonical English built-in name matching the styleId — `name: "Body Text"` for `id: "BodyText"`, `name: "heading 1"` for `id: "Heading1"`. Word applies its own UI localization for display.
- **When creating new with a custom styleId** (e.g. `BodyEmphasis`), the simplest safe rule is `name = id`.

The engine catches direct string-equal collisions plus the major en/zh-CN aliases at preflight.

**Three layers for setting fonts**, in increasing breadth:

- **Per-role override** (`styles[]` entries with `fontCJK` / `fontLatin` / `overrides`): targeted, only the styled paragraphs change. Use when the user is talking about specific roles ("标题黑体, 正文宋体", "Heading2 字号改小一号").
- **Whole-doc default** (declare a `Normal` entry; other styles' `basedOn: "Normal"` chain inherits): wide, covers most pStyle-bound paragraphs via cascade. Use when the user expresses uniform default ("整篇统一用 Times New Roman").
- **Document-design font scheme** (`theme.fonts` block): widest, modifies theme1.xml so any docDefaults / styles / runs that reference theme fonts auto-resolve to the new values. Use when the user is talking about the document's design layer ("把这份文档的主题字体改成 X / Y").

Bias toward the narrowest layer that captures the intent — wider layers risk surprising the user with effects on chrome they didn't intend to touch.

**Chinese font size names** (初号/一号/.../小六): see `references/chinese-font-sizes.md` for the pt mapping when the user specifies sizes in Chinese terms.

## Designing numbering schemes

When the document has typed heading prefixes (chrome or content prefixes), migrate to automatic numbering. Skip only when the user explicitly opts out, or no numbered headings exist.

If the manual scheme itself is inconsistent across the document — e.g. H1 has numbers in chapter 1 but not chapter 2, or H2 uses chapter-prefixed `"1.1"` in some chapters and per-chapter-restart `"1."` in others — auto-migration is a normalization decision that may change author-intended semantics. Ask the user before applying rather than picking one scheme silently.

Each level binds to a heading style via `styleId`; higher levels reset lower-level counters automatically. Field names mirror OOXML: `numFmt` (e.g. `decimal` / `chineseCounting` / `bullet`), `lvlText` (the rendered prefix pattern, e.g. `"%1."` / `"%1.%2"` / `"第%1章"`), and `suff` controls the gap between the marker and the paragraph text — `"space"` when the marker ends in a digit or character (`1. Title`, `第一章 研究方法`, `1.1 概述`), `"nothing"` when the trailing punctuation already separates them (`一、研究方法`, `（一）背景`), `"tab"` only for wide-list layouts.

See `references/numbering-formats.md` for full value tables and ready-made templates.

**`outlineLevel` is independent of numbering level.** They co-occur for headings but are different OOXML concepts: `outlineLevel` (set on the paragraph style's pPr) controls TOC inclusion, navigation pane, outline view. Numbering level (the `level` field above) controls auto-number display. Set `outlineLevel` explicitly on heading styles via `overrides`; don't expect numbering to imply it. List styles (`ListBullet` / `ListNumber`) take numbering but **must not** carry `outlineLevel` — otherwise list items pollute the TOC.

**`stripPrefixPatterns` defaults to `[lvlText]`** — for the simple case where the doc's manual prefix matches the new auto-numbering pattern, you don't need to write it. Specify it only when the source mixed prefix styles within one role. Patterns are tried in order, first match wins; **longer pattern must come first** or shorter ones swallow prefixes the longer ones wanted.

The `%N` placeholder in a pattern matches Arabic digits or the common Chinese numerals (一二三...百千). It does NOT match Roman numerals, Latin letters, less-common Chinese forms (壹貳叄), circled digits (①), or other locale-specific shapes. The dry-run report's "Numbered-style paragraphs not matched by any stripPrefixPattern" section surfaces unstripped paragraphs with leading-text samples — read those samples and add a `pattern_rules` entry with `stripMatch: true` and an explicit regex covering the unrecognized shape.

**Heading numbering vs list numbering — different schemes:**

- **Heading-class** (outline-bearing prefixes that define document structure): convert to `Heading1` / `Heading2` / ... bound to **one unified multi-level scheme**.
- **List-class** (local enumerations inside body sections): convert to `ListNumber` / `ListBullet` bound to a **separate single-level scheme** (restart per group).

Detection signals:

- Position: heading sits at section start, followed by short title text; list item sits inside a section among other list items or body prose.
- Surrounding format: heading already styled distinctly (bold, larger size); list items sit in body-style.
- Pattern shape: headings tend to decimal hierarchy or chapter sentinels; lists tend to parenthesized or single-level digits.
- Depth: headings nest 2–4 levels; lists usually 1.

Both numbering schemes go in the same config — `numbering: [{...heading...}, {...list...}]`.

## Targeted Restyle (when scope is narrow)

The user is expressing focused changes with the rest of the document expected to stay untouched. Decisive signal: *scope narrowing* — they name a specific change or ask that something be preserved. *Illustrative phrasings: "加个 X 样式 / 其他不动", "Heading2 字号改小一号", "保留手动编号，只调字体".*

**Mindset:** small, focused, additive. Locate target → minimal config → dry-run → verify → apply.

### Workflow

1. **Locate target paragraphs.**
   - Content-pattern targets: `find_paragraphs --regex` returns matching paragraphs with index, fingerprint, and text preview, without the overhead of a full overview.
   - Visual targets ("just the H2s"): `inspect_style <fingerprint>` if you can identify the fingerprint, or run a quick `overview` to spot it.
   - Single-paragraph targets: `inspect_range <para> <para>` for full computed style.

2. **Decide the target style.**
   - Reuse an existing styleId (check via `inspect_style_def`) when possible — preserves whatever's already wired.
   - Define a new style only when no existing one fits.

3. **Pick the narrowest tool.**
   - `restyle` for paragraph style assignment only — same config as `apply` minus `template` / `numbering` / `edits`. Most common.
   - `migrate_numbering` for numbering-only changes.
   - `import_template` for template-only imports.
   - `apply` when the change spans multiple operations.

4. **Sparse config.** Don't redeclare untouched styles. Prefer `pattern_rules` (content-based) or targeted `assignments`. Avoid broad `bulk_rules` unless the fingerprint cleanly captures only your target.

5. **Dry-run + apply.**

### Notes

- **No fingerprint coverage requirement.** Untargeted paragraphs simply stay as they are.
- **Run-level direct formatting on untouched paragraphs is preserved.** The uniform-strip rule fires only on paragraphs the script restyles.
- **When a request grows beyond "targeted":** if you find yourself adding 5+ styles, declaring numbering, or reaching for a template — switch to the default workflow.

## Escape Hatch (manual XML)

For requests that can't be expressed via the config — custom watermarks, embedded objects, raw OOXML constructs — fall back to manual XML editing of the unzipped docx.

**This bypasses every safety net** the skill provides: validation, run-level formatting preservation, numId migration, original-file protection. Use only when no other path fits, and tell the user explicitly.

Outline:
1. Unzip the docx (`unzip docx -d /tmp/docx-unpacked/`).
2. Edit the relevant XML file(s) — typically `word/document.xml`, `word/styles.xml`, `word/numbering.xml`, `word/header*.xml`. Preserve namespaces, element ordering, and `xml:space="preserve"` on whitespace.
3. Re-zip preserving the directory structure (`cd /tmp/docx-unpacked && zip -r ../output.docx .`).
4. Open in Word; if it errors or silently drops content, your XML edit broke something — do not deliver.

If a request *can* be expressed via `apply` config, do that instead.

---

## Reading the Overview Output

`overview` prints visual style summary + document skeleton inline. Four conventions worth knowing in advance:

- **Letter vs hash labels.** `[A]`, `[B]` are sorted by frequency in this run (volatile across edits); the summary also shows a 6-char content hash next to each letter (`A [c4f9]: ...`). `bulk_rules.fingerprint` accepts either — use letters for in-session iteration, hashes in configs you intend to keep across doc revisions.
- **Numbered ≠ unnumbered fingerprints.** The fingerprint hash includes whether a paragraph carries a numbering reference, so visually identical paragraphs split into separate fingerprints (suffix "List") when one is auto-numbered and the other isn't.
- **Layout vs data tables.** Layout tables are inlined into the skeleton between `--- LAYOUT TABLE ---` markers; data and form tables are summarized as a single non-paragraph block.
- **Empty paragraphs and truncation.** Consecutive empty paragraphs are compressed (`--- empty ×N ---`); paragraph text is truncated to ~40 chars in the skeleton — use `inspect_range` for full text.

## Edge cases

- **Empty paragraphs as spacing**: preserve them. Removing is structural, not stylistic, and risks breaking cover-page layout.
- **Table caption vs figure caption**: table captions go ABOVE the table, figure captions BELOW. Use `inspect_neighbors` to confirm.
- **Table footnotes**: text right after a table starting with "注：" / "Note:" is a footnote, not body text.
- **Unnumbered special headings** (摘要 / Abstract / 目录 / 参考文献 / 致谢 / 附录): share the visual style of Heading1 but have no chapter number. Use `HeadingNoNum` or suppress numbering on the same style via `exclude`.
- **Appendix numbering** often restarts with a different scheme (附录A / A.1 / A.2) — may need a second `numbering` entry.
- **Pre-printed chrome (forms, templates with printed labels and instructions)**: when the visual summary shows a long tail of low-occurrence fingerprints with short average text length (`avg ≤20ch`), those are usually printed labels / cover chrome, not author content.
- **Source's base style violates the document's stated specs**: e.g. doc's `Normal` sets `bold: true` while the spec says 正文不加粗. Override the base style in `styles[]` with the corrected fields.

## Compose with other shapes

- After installing the style system, add an `edits[]` block to the same `apply` config (or follow up with another call) for content insertion or surgical touch-ups on specific paragraphs the rules missed. See [edit.md](edit.md).
- Read-only check before reshape: `audit`. The audit's violation list often translates directly into the standardize-shape blocks (`styles[]` / `pattern_rules` / `numbering`) of an `apply` config.
