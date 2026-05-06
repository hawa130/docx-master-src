---
name: docx-normalize
description: "Automate Word (.docx) formatting work: classify paragraph roles, inject named styles, migrate manual heading numbers (1./1.1/第N章) to multi-level auto-numbering, import template styles, apply targeted single-style edits, audit a document against a typography spec. Use whenever the user wants to format, restyle, normalize, standardize, or audit a Word document — including phrasings like '排版', '格式', '套模板', '按学校 / 期刊格式', '标题字号不对', '统一一下样式', '加个图注样式', '其他不动只改 X', applying thesis/paper specs, aligning to a template, fixing inconsistent formatting. Do NOT use for: PDFs, spreadsheets, or plain-text / Markdown source files (unless the task is specifically to *output* a docx)."
---

# Document Formatting Normalizer

Automate the mechanical version of Word typesetting: classify paragraph roles, inject named styles, migrate manual numbering to automatic numbering, apply targeted edits, and audit conformance — all by mutating the docx's OOXML directly. Output is a new docx plus a change report; the original file is never touched.

## Core Principle

**You are the analyst. The tools are your instruments.**

Tools only present facts — computed styles, element positions, document structure. They never classify or judge. All semantic reasoning is yours: deciding what role a paragraph plays, what style to name it, whether two similar formats should merge or stay separate, how to handle edge cases.

## What Are You Doing?

Pick the path that matches the user's request. Each path's workflow is below.

| User intent | Path |
|---|---|
| 完整排版 / 套学校格式 / 按这个模板做 / "排版乱了帮我整理" | → **Full Standardization** |
| 加 / 改 / 修单一处样式（"加个 Reference 样式" / "Heading2 字号改小一点" / "其他不动") | → **Targeted Edit** |
| 检查文档合不合规范，但**不**修改 | → **Audit** |
| 上面没覆盖（自定义水印、特殊 OOXML 操作等） | → **Escape Hatch** |

When in doubt — user just hands over a docx without specifics, or says broad things like "排版一下" / "按这个模板" — default to **Full Standardization**. But when scope is explicitly narrowed ("只改..." / "保留..." / "加一个 X，其他不动" / "Heading2 改小一号"), don't fall back to full; **Targeted Edit** is the right tool.

## Iterating is Normal

You don't have to plan the entire transformation upfront and submit one giant config. Apply changes incrementally:

- `apply_styles --dry-run` between edits — seconds per cycle, no file written
- Re-call `inspect_*` between edits to stay grounded in actual document state, not in your earlier mental model
- `styles[]` is **sparse by design** — only declare styles you're touching; untouched styles stay as they are
- Untargeted paragraphs go to **implicit-keep**; that's correct, not a missed decision (unless you're on the Full Standardization path, where coverage is expected)

Small step → dry-run → verify → next step is a fully supported workflow, not a fallback. Use it whenever the change is bounded.

## Tool Reference

All tools invoked via `node <script> <args>`, output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | First call on Full Standardization or Audit. Returns metadata, page setup (mm), theme, style definitions, numbering schemes (clustered by pattern), visual style statistics, and document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | Full text and computed styles for a paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | Per-run rPr dump when a paragraph has run-level mixed formatting (bold lead phrase, colored numbering prefix, inline emphasis). Tells you which properties are uniform vs. mixed across runs — critical for predicting what `apply_styles` will preserve vs. strip. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | What surrounds a paragraph: nearby images, tables, equations, page breaks, sibling paragraphs. Default radius 4. **First choice for figure-caption / table-caption / first-after-heading classification.** |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | When you see a fingerprint in the overview and need to understand what role it plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | Pre-defined styles in `styles.xml` — what they currently are and what `basedOn` chain produces them. Use before reusing or overriding an existing styleId. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | Page setup differences between sections (headers, footers, page numbering). |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. Use to discover content-defined roles (figure/table captions, references, keywords) and to validate `pattern_rules` regex coverage before applying. **First choice on the Targeted Edit path** — it locates exactly the paragraphs you want to change, without requiring a full overview. |
| `apply_styles --dry-run` | `node scripts/apply_styles.js --dry-run <config.json>` | Iterate on a config without writing the output file. Returns the full change report including sample affected paragraphs and the Style Resolution annotation block. **Use between every config edit** — it's seconds per cycle. |
| `apply_styles` | `node scripts/apply_styles.js <config.json>` | Final step. Outputs the reformatted document and a change report. |

---

## Path: Full Standardization

When the user wants the document brought into a consistent, standardized form — typical phrasings: "帮我排版", "按学校格式", "套这个模板", or just handing over a messy docx without specifics. The default path when scope is unclear.

### Step 1: Understand the Goal

Read the user's request and any attached files. The user may give you:

- **Just a docx** — apply full default standardization.
- **Explicit text guidelines** ("一级标题三号黑体加粗, 正文小四宋体1.5倍行距") — translate the natural language into structured `styles[i]` fields yourself. The script does NOT parse it. Pass the user's wording *verbatim* into `requirements: { <styleId>: "..." }`; the change report prints it side-by-side with your resolved fields for visual verification. See Step 4 for 字号/字体/颜色 mappings.
- **Template / reference document** — pass via `template: { source, styles: [...] }`. The script clones the named styles' full definitions (with basedOn ancestors) into source's styles.xml and migrates referenced numIds. Run `overview` / `inspect_style_def` on the template first to choose which styleIds to import. The template's *style system* transfers; its *document structure* (chapter count, content, page setup) does not.

User-supplied typography preferences ("正文宋体小四") and "附加 / 另外 / 顺便 X" phrasings *add to* full standardization — they don't replace it. When intent is genuinely ambiguous (role assignments, spec-vs-doc conflict), ask one focused question rather than guess.

### Step 2: Inspect the Document

Start by calling `overview`. This gives you the full picture in one call.

Read the overview carefully. Form hypotheses about:
- What type of document is this? (thesis, report, contract, manual, letter...)
- Where does the structural information live — in `styles.xml` definitions, in the content itself, or both?
- How many distinct visual styles exist? Do they map cleanly to semantic roles?

Then drill in with `inspect_*` tools only as needed. For a simple document the overview alone may be sufficient; for a complex one you might inspect 3-5 areas.

### Step 3: Classify by Visual Fingerprint

Map each visual fingerprint to a semantic role. The overview provides deduplicated fingerprints (A, B, C, ...) with occurrence counts — work at this level, not paragraph-by-paragraph.

For most fingerprints, the mapping is obvious from format + context:
- The largest bold centered text appearing once → `Title`
- Bold text matching a numbering pattern → `Heading1` / `Heading2` / `Heading3`
- The most frequent plain text style → `BodyText`
- Small centered text near images/tables → `FigureCaption` / `TableCaption`
- Monospace font → `Code`

Only drill into individual paragraphs (via `inspect_range` or `inspect_style`) when a fingerprint is ambiguous — e.g. the same visual format is used for both table captions and figure captions. Common roles:

**Structural roles:**
- `Title` — document title (usually appears once, cover page)
- `Heading1` / `Heading2` / `Heading3` — section headings at different levels
- `HeadingNoNum` — unnumbered section headings (摘要, Abstract, 参考文献, 致谢, 附录)
- `BodyText` — main body paragraphs
- `FirstParagraph` — first paragraph after a heading (when styled differently from the rest of body)

**Caption/label roles:**
- `FigureCaption` — caption below an image (图 X-X ...)
- `TableCaption` — caption above a table (表 X-X ...)
- `TableFootnote` — notes below a table (注：..., 数据来源：...)
- `EquationCaption` — equation number, usually right-aligned (X-X)

**List roles:**
- `ListBullet` — unordered list items
- `ListNumber` — ordered list items
- `ListContinue` — continuation paragraph within a list item

**Special roles:**
- `Code` — code blocks (monospace font, possibly shaded background)
- `Quote` — block quotations (indented, possibly different font)
- `Reference` — bibliography entries ([1] ..., with hanging indent)
- `Keywords` — keyword line (关键词：... / Keywords: ...)
- `Abstract` — abstract body text (may differ from normal body)
- `BodyEmphasis` — uniformly-bold body paragraphs acting as in-paragraph sub-titles or labels (no outline level, no auto-numbering; often short phrases but longer fully-bold paragraphs fit too). The style **must explicitly set `bold: true`** — when bold is uniform across the paragraph's runs, restyle strips it as redundant direct formatting, so the new style needs to carry it.

**Fixed content (do not restyle — preserve as-is):**
- Cover page elements (school name, field labels, date)
- Header/footer content
- Table of contents (auto-generated)

Only create roles that actually exist in the document. If you discover roles not listed here (e.g. `Theorem` in math papers), create them.

### Step 4: Define the Style System

**Source of values** (priority order): (1) user requirements, (2) template document, (3) values extracted from a representative paragraph in the document, (4) sensible defaults (last resort). Don't invent values when a representative paragraph exists.

**Two modes per `styles` entry:**

1. **`fromParagraph`** (preferred when extracting from the doc): pick the first occurrence of the dominant fingerprint for the role and set `fromParagraph: <index>`. The tool extracts the full computed rPr + pPr from that paragraph's *dominant text run* (longest non-numbering-prefix run, so `"1.1 研究方法"` extracts the title formatting, not the prefix's). Use `overrides` to add fields the source lacks (e.g. `outlineLevel`) or apply user-requested specifics.
2. **Manual mode**: specify fields directly — when no representative paragraph exists, when synthesizing a role, or when the user fully specified the style.

Modes can mix within one `styles` array:

```json
{
  "styles": [
    { "id": "BodyText", "name": "正文", "fromParagraph": 60 },
    { "id": "Heading2", "name": "二级标题", "fromParagraph": 19,
      "overrides": { "outlineLevel": 1 } },
    { "id": "Caption",  "name": "图表注", "font": "宋体", "size": 10.5,
      "alignment": "center", "lineSpacing": 1.5 }
  ]
}
```

**Normalization rule:** for outliers (e.g. Heading1 appears 5 times, 4 of one pattern + 1 different), source from the majority. The same applies when two fingerprints play the same role — take the majority's values, route both fingerprints' paragraphs to one style. "Normalize" means routing inconsistent paragraphs to one consistent style, NOT replacing the author's choices with values you think look better.

**What `fromParagraph` extracts:** font, fontEastAsia (only if different from font), size, bold/italic (only if true), color (only if not auto), alignment, spaceBefore, spaceAfter, lineSpacing (with original lineRule preserved), firstLineIndent, hangingIndent, outlineLevel (only when the source has it set — add via `overrides` if you need it on a heading whose source paragraph lacks it).

**Does NOT extract:** `numId` / `numLevel` — numbering is bound through `numbering.levels[].styleId`, not hardcoded per paragraph.

**`font` vs `fontEastAsia`:** `font` is the source's ASCII/Latin slot (often inherited from theme defaults like Arial / Times); `fontEastAsia` holds the CJK font. A paragraph rendering as Chinese can still extract `font: "Arial"` — that's the Latin font that would render Latin characters in the same paragraph. When the user names only a Chinese font ("正文宋体" / "标题黑体"), set `fontEastAsia` only and leave `font` (ASCII) unset so the source's Latin font is preserved. Set both fields only when the user explicitly says the same font should apply to Latin too.

**Indent unit preservation:** when the source used Word's character-based indent (`w:firstLineChars` / `w:hangingChars`, what Word writes for "首行缩进 N 字符"), extraction gives `"Nchar"` so font-size auto-scaling round-trips. Fixed twips give `"Npt"`. Don't manually convert "char" values to pt — that locks the indent to one font size.

**When the document already defines the style ID you want:** if its parameters match your target, reuse as-is. If they differ, override (the script updates the existing definition rather than creating a duplicate). But first verify the style is actually used for its intended role — overriding `Heading1` while it's misused as body text would corrupt those paragraphs; reassign the paragraphs first. Use Word built-in IDs (`Heading1` / `Heading2` / `BodyText` / `Caption`) when the role matches, so TOC / nav / outline view work; never create parallel styles like `MyHeading1`.

**Chinese font size names** (初号/一号/.../小六): see `references/chinese-font-sizes.md` for the pt mapping when the user specifies sizes in Chinese terms.

### Step 5: Define the Numbering Scheme

When the document has typed heading prefixes (`"1. 引言"` / `"1.1 研究方法"` / `"第N章 ..."`), migrate to automatic numbering — this is part of standardization. Skip only when the user explicitly opts out, the source already has real `numId` references you want to preserve (verify with `inspect_range` — typed-text prefixes look identical to auto-numbers but behave totally differently), or no numbered headings exist.

If the manual scheme itself is inconsistent across the document — e.g. H1 has numbers in chapter 1 but not chapter 2, or H2 uses chapter-prefixed `"1.1"` in some chapters and per-chapter-restart `"1."` in others — auto-migration is also a normalization decision that may change author-intended semantics. Ask the user before applying rather than picking one scheme silently.

Each level binds to a heading style via `styleId`; higher levels reset lower-level counters automatically. Field names mirror OOXML: `numFmt` (e.g. `decimal` / `chineseCounting` / `bullet`) and `lvlText` (the rendered prefix pattern, e.g. `"%1."` / `"%1.%2"` / `"第%1章"`); see `references/numbering-formats.md` for value tables. Minimal example for three-level decimal headings:

```jsonc
"numbering": {
  "levels": [
    { "level": 0, "numFmt": "decimal", "lvlText": "%1.",     "styleId": "Heading1",
      "stripPrefixPatterns": ["%1."] },
    { "level": 1, "numFmt": "decimal", "lvlText": "%1.%2",   "styleId": "Heading2",
      "stripPrefixPatterns": ["%1.%2", "%1."] },
    { "level": 2, "numFmt": "decimal", "lvlText": "%1.%2.%3", "styleId": "Heading3",
      "stripPrefixPatterns": ["%1.%2.%3", "%1.%2", "%1."] }
  ]
}
```

**Mixed manual prefix styles within one role:** authors often mix patterns at the same heading level — e.g. chapter 1's H2s are "1.1 ..." while chapter 2's are "1. ..." (restart per chapter). One regex can't normalize both. Use `stripPrefixPatterns: ["%1.%2", "%1."]` — patterns tried in order, first match wins; longer pattern must come first or `"%1."` will strip just "1." from "1.1 ..." leaving ".1 ...".

**Preserving design colors on numbers:** if the source styles numbers in a different color/weight than title text (e.g. blue numbers + black bold titles), set `numRPr` on the level. The marker is rendered with this rPr; the title uses the paragraph style.

### Step 6: Review Plan Before Execution

Before calling `apply_styles`, self-check:

1. **Style values have sources** — every parameter came from user spec, template, or `inspect_style` extraction. None were invented.
2. **Heading styles have `outlineLevel`** — required for TOC / nav / outline view.
3. **Numbering migrated** when source has typed heading prefixes (per Step 5). `stripPrefixPatterns` covers mixed variants within a role.
4. **Every fingerprint has a decision** — restyle / keep / exclude / flag. No fingerprint left unaccounted for. *(This coverage rule applies to Full Standardization only — Targeted Edit and Audit paths do not require it.)*

Fix any issue before proceeding.

### Step 7: Execute

Call `apply_styles` with your decision in a JSON config.

**Top-level fields:**

```jsonc
{
  source, output,                          // REQUIRED. Input/output paths (must differ).
  dryRun,                                  // optional. Preview without writing the file.

  styles: [ ... ],                         // REQUIRED. Paragraph styles to inject —
                                           //   either via fromParagraph extraction or
                                           //   manual fields, with optional overrides.

  numbering: { levels: [ ... ] },          // Multi-level auto-numbering bound to heading
                                           //   styles. See Step 5 for when to include / skip.

  template: { source, styles: [ ... ] },   // optional. Import named styles from another
                                           //   docx; basedOn ancestors auto-pulled,
                                           //   numId references migrated.

  requirements: { Heading1: "原话...", BodyText: "..." },  // optional. ANNOTATION ONLY — script records
                                           //   the user's natural-language spec next to
                                           //   the agent-resolved fields in the report
                                           //   for visual verification. Not parsed.

  // Paragraph-to-style mapping, in resolution order:
  exclude:       [ idx, ... ],
  assignments:   [ { para, action, style?, reason? }, ... ],
  pattern_rules: [ { regex, style, stripMatch? }, ... ],
  bulk_rules:    [ { fingerprint, style }, ... ],
}
```

Full schema in `references/apply-styles-config.md` — read once before composing your first config.

Key invariants:
- Paragraph indexing is 1-based, matching `#NNN` labels in the skeleton. Layout-table paragraphs are indexed; data/form tables are not.
- Paragraph mapping order (first match wins): `exclude > assignments > pattern_rules > bulk_rules > implicit-keep`.
- Style-field priority (later wins): defaults → template-imported → fromParagraph → direct fields → overrides.
- Restyle behavior: when a paragraph is restyled, run-level direct formatting that is *uniform across all runs* gets stripped (so the new style's defaults take effect). Direct formatting that *differs between runs* is preserved as intentional inline emphasis — bold lead phrase + non-bold body, colored numbering prefix + black title, etc. survive automatically.

### Step 8: Validate and Report

Iterate with `apply_styles --dry-run` first. The change report includes a per-style sample of the first affected paragraphs and a Style Resolution block listing every injected style — styles with a `requirements` entry show user spec next to resolved fields for side-by-side verification, styles without one show the resolved fields alone (still auditable). Read both before committing.

**Safety guarantees:**

- The original file is never modified — `apply_styles` writes a fresh copy.
- The output is validated before being kept; if validation fails the output is discarded, the original returned unchanged, and the error reported. Don't silently retry on validation errors — surface them.
- Section properties (page size, margins, headers, footers, columns) are never modified.

**When to `flag` vs. apply:** flag when the *role assignment* is genuinely uncertain (could be heading or emphasized body, prefix doesn't match any known pattern, ambiguous between two roles). Don't flag formatting variance within a clear role (one heading is 15pt while the rest are 16pt — just normalize it).

**Hand-off:** present a concise summary of the change report to the user, then deliver the output. If the document contains a TOC, remind the user to right-click → "Update Field" in Word after opening.

---

## Path: Targeted Edit

When the user wants a focused change to an already-formatted document while leaving everything else untouched. Triggers: "加个 X 样式 / 其他不动", "Heading2 字号改小一号", "把所有 [N] 开头的段落统一缩进", "保留手动编号，只调字体".

**Mindset:** small, focused, additive. Locate target → minimal config → dry-run → verify → apply. Don't try to classify the whole document.

### Workflow

1. **Locate target paragraphs.**
   - Content-pattern targets: `find_paragraphs --regex` is the right tool — it returns matching paragraphs with index, fingerprint, and text preview, without the overhead of a full overview.
   - Visual targets ("just the H2s"): `inspect_style <fingerprint>` if you can identify the fingerprint, or run a quick `overview` to spot it.
   - Single-paragraph targets: `inspect_range <para> <para>` for full computed style.

2. **Decide the target style.**
   - **Reuse an existing styleId** (check via `inspect_style_def`) when possible — preserves whatever's already wired (TOC links, outline level, basedOn chain). Just override the fields the user wants changed.
   - **Define a new style** only when no existing one fits: `fromParagraph` extracts from a representative paragraph; or specify fields directly. See § Step 4 (Define the Style System) for extraction details.

3. **Write a sparse `apply_styles` config.**
   - `styles[]`: only the styles you're touching. **Don't redeclare untouched styles** — that's noise that risks accidental overrides.
   - Mapping: prefer `pattern_rules` (content-based) or targeted `assignments` (specific paragraphs by index). Avoid broad `bulk_rules` unless a fingerprint cleanly captures your target — fingerprint changes feel cheap but can spread.
   - Skip `numbering` and `template` blocks unless those are the change.

4. **Dry-run.** The change report should touch only paragraphs you intended. Untouched fingerprints appear under "implicit-keep" — that's correct behavior on this path, not a missing decision.

5. **Apply.**

### Notes specific to this path

- **No fingerprint coverage requirement.** Step 6's coverage check is a Full Standardization rule. Here, untargeted paragraphs simply stay as they are.
- **Run-level direct formatting on untouched paragraphs is preserved.** The uniform-strip rule fires only on paragraphs the script restyles; everything else keeps its inline formatting.
- **Iterate.** Apply one change, dry-run, verify, then layer the next change on the result. Two small configs are easier to debug than one large one.
- **When a request grows beyond "targeted":** if you find yourself adding 5+ styles, declaring numbering, or reaching for the template — switch to Full Standardization. Targeted Edit's discipline is keeping the change small.

---

## Path: Audit

When the user wants to check whether a document conforms to a spec **without modifying it**. Read-only.

### Workflow

1. **Read the spec.** What rules need verifying? Body font / size / line spacing? Heading hierarchy and outline levels? Image-caption format? Page margins? Reference style?

2. **Gather facts.** `overview` for the global picture; `inspect_style`, `inspect_style_def`, `inspect_section`, `find_paragraphs` for specifics. The same tools as the other paths — just no `apply_styles`.

3. **Compare and report.** Produce a structured violation list:
   - What the spec required
   - What the document actually has
   - Which paragraphs / sections are affected (use letter labels and `#NNN` indices so the user can navigate to issues directly)
   
   If the user wants to fix the violations after seeing the audit, that becomes a separate Full Standardization or Targeted Edit request — don't auto-fix without permission.

### Tip

If the spec is long (a 20-page school formatting standard), don't try to audit every clause in one pass. Group violations by area (cover page / headings / body / figures / references) and let the user prioritize.

---

## Path: Escape Hatch

For requests the above paths can't express — custom watermarks, embedded objects, raw OOXML constructs — fall back to manual XML editing of the unzipped docx.

**This bypasses every safety net** the skill provides: validation, run-level formatting preservation, numId migration, original-file protection. Use only when no other path fits, and tell the user explicitly that you're doing it.

Outline:
1. Unzip the docx (`unzip docx -d /tmp/docx-unpacked/`).
2. Edit the relevant XML file(s) — typically `word/document.xml`, `word/styles.xml`, `word/numbering.xml`, `word/header*.xml`. Preserve namespaces, element ordering (e.g. `<w:pPr>` child order), and `xml:space="preserve"` on whitespace.
3. Re-zip preserving the directory structure (`cd /tmp/docx-unpacked && zip -r ../output.docx .`).
4. Open in Word; if it errors or silently drops content, your XML edit broke something — do not deliver.

If a request *can* be expressed via `apply_styles` config, do that instead. The escape hatch is a last resort, not a quick path.

---

## Important Guidelines

### What This Skill Does
- Inject named style definitions into `styles.xml`
- Replace direct formatting with style references on paragraphs
- Convert manual numbering to automatic numbering
- Normalize inconsistent formatting to the majority pattern
- Audit a document against a typography spec (read-only)

### What This Skill Does NOT Do
- Edit content (text, grammar, structure) — except stripping manual numbering / bullet prefixes ("第1章 " / "1.1 " / "• ") when you configure `stripPrefixPatterns` on the matching numbering level. The script removes only what your patterns match; nothing is auto-stripped without configuration.
- Restyle paragraphs inside data tables, or modify any table structure (cell sizes, borders, cell-level formatting).
- Update the TOC itself — heading `outlineLevel` is set correctly, but the user must right-click → "Update Field" in Word after opening.
- Modify field codes (`STYLEREF`, `TOC`, `REF`, `DATE`, ...) — preserved as-is.
- Trigger Word / LibreOffice rendering (recompute TOC content, refresh cross-references, rerun field calculations). Boundary: scripts manipulate OOXML; rendering is the user's job.

### Edge Cases to Watch For

- **Empty paragraphs as spacing**: preserve them. Removing is structural, not stylistic, and risks breaking cover-page layout.
- **Table caption vs figure caption**: table captions go ABOVE the table, figure captions BELOW the figure. Use `inspect_neighbors` to confirm which side the image/table is on.
- **Table footnotes**: text right after a table starting with "注：" / "来源：" / "Note:" is a footnote, not body text.
- **Unnumbered special headings** (摘要 / Abstract / 目录 / 参考文献 / 致谢 / 附录): share the visual style of Heading1 but have no chapter number. Use `HeadingNoNum` or suppress numbering on the same style.
- **Appendix numbering** often restarts with a different scheme (附录A / A.1 / A.2) — may need a second `numbering` entry.
- **Layout vs data tables**: the overview tool classifies these — layout tables (single-cell content containers) are inlined into the skeleton; data/form tables are summarized. Verify the classification when inspecting unfamiliar table-heavy documents.

## Document Skeleton Format

`overview` outputs metadata, page setup, theme, style definitions, numbering schemes, visual style summary, and a document skeleton:

```
=== Visual Style Summary (deduplicated) ===
A: 宋体 22pt Bold Center      ×1
B: 宋体 16pt Bold              ×6
C: 宋体 14pt Bold              ×18
D: 宋体 12pt 1stIndent          ×89
...

=== Document Skeleton ===
--- Section 1 (para #1-#11) ---
Header: (none)
Footer: (none)

  #001 [A]  "某某大学"
  #002 [A]  "本科毕业设计（论文）"
  ...

--- Section 2 (para #12-#28) ---
Header: "某某大学本科毕业论文"
Footer: Roman numeral page number

  #012 [B]  "摘  要"
  #013 [D]  "（摘要内容）"

--- Section 3 (para #29-#68) ---
Footer: Arabic page number (restart from 1)

  #029 [B]  "第1章  绪论"
  #030 [C]  "1.1  研究背景及意义"
  #031 [D]  "本文针对某某问题展开研究..."
  ...
  #049 [E]  "表 3-1 不同方法的性能对比"
  --- TABLE (5×4) headers:["方法","Precision","Recall","F1"] ---
  #050 [F]  "注：加粗数据表示最优结果"
  --- IMAGE (14cm × 8cm) ---
  #062 [E]  "图 3-2 实验结果对比"
  #081 --- empty ×3 ---
```

**Conventions:**
- `[A]`, `[B]` are letter labels (sorted by frequency in this run — volatile across edits). The summary also shows a 6-char content hash next to each letter (e.g. `A [c4f9]: ...`) — `bulk_rules.fingerprint` accepts either form. Use the letter for in-session iteration; use the hash in configs you intend to keep across doc revisions, since the hash stays stable when paragraphs are added/removed and frequency-rank shifts.
- The fingerprint hash includes font, size, weight/italic, color, alignment, first-line-indent, AND whether the paragraph carries a numbering reference — so visually identical paragraphs split into different fingerprints (with "List" suffix) when one is auto-numbered and the other is plain body. `bulk_rules` can target list items independently.
- Non-paragraph elements appear as `--- TYPE (details) ---`. Consecutive empty paragraphs are compressed: `--- empty ×N ---`.
- Layout tables (single-cell content containers) are expanded inline under `--- LAYOUT TABLE ---` / `--- END LAYOUT TABLE ---` markers; data tables and form tables are summarized.
- Text is truncated to ~40 chars; use `inspect_range` for full text.

## File Structure

```
docx-normalize/
├── SKILL.md
├── scripts/
│   ├── overview.js                   ← Document overview and skeleton
│   ├── inspect_range.js              ← Detailed paragraph range view
│   ├── inspect_runs.js               ← Per-run rPr dump + run-level diversity summary
│   ├── inspect_neighbors.js          ← Adjacent elements (image/table/para/break) ±radius
│   ├── inspect_style.js              ← Visual fingerprint occurrences
│   ├── inspect_style_def.js          ← Named style definition details
│   ├── inspect_section.js            ← Section page setup details
│   ├── find_paragraphs.js            ← Regex search across paragraph text
│   └── apply_styles.js               ← Execute formatting changes; supports --dry-run
└── references/
    ├── apply-styles-config.md        ← Full apply_styles config schema (read before composing your first config)
    ├── numbering-formats.md          ← Numbering format reference (read when handling numbered headings)
    └── chinese-font-sizes.md         ← 初号/一号/.../小六 → pt mapping (read when user specifies Chinese font sizes)
```
