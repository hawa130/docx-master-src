---
name: docx-normalize
description: "Normalize Word (.docx) formatting: classify paragraphs, inject named styles, automate numbering, import template styles. Use when reformatting or standardizing a docx, applying thesis/paper typography specs, or aligning to a template."
---

# Document Formatting Normalizer

Read a user's Word document, identify the semantic role of every element (headings, body text, captions, lists, code blocks, etc.), inject standardized named styles, replace direct formatting with style references, and convert manual numbering into automatic numbering. Output a reformatted copy of the original document plus a change report.

## Core Principle

**You are the analyst. The tools are your instruments.**

Tools only present facts — computed styles, element positions, document structure. They never classify or judge. All semantic reasoning is yours: deciding what role a paragraph plays, what style to name it, whether two similar formats should merge or stay separate, and how to handle edge cases.

## Workflow

### Step 1: Understand the Goal

Check what the user has provided. Only ask for clarification if the goal is genuinely ambiguous — otherwise proceed with defaults and state your assumptions in the report.

- **Explicit text guidelines** ("一级标题三号黑体加粗, 正文小四宋体1.5倍行距"): translate the user's natural language into structured `styles[i]` fields yourself. The script does NOT parse natural language — that would be brittle (it can't tell "不要加粗" from "加粗", can't expand "比一级小一号", and silently fails on synonyms / Chinese numerals / sentence grammar). You're an LLM; translation is exactly what you're for. Use the 字号 / 字体 / 颜色 reference tables in this document (Step 4) to map keywords. Pass the user's original text *verbatim* into `requirements: { styleId: "..." }` — the script will display "user specified X / agent resolved {Y}" side-by-side in the change report so any reviewer can verify your translation by eye.
- **Template / reference document**: pass it via `apply_styles`'s `template: { source, styles: [...] }` field. The script copies the named styles' full pPr/rPr definitions (including basedOn ancestors) into the source's styles.xml, and migrates any referenced `numId` to fresh IDs in the source's numbering.xml. You don't need to run a separate analysis pass on the template by hand — but DO run `overview` / `inspect_style_def` on the template first to know which styleIds are worth importing.
- **No guidelines**: infer the intended style system from the document itself, identify inconsistencies, and normalize to the majority pattern.
- **Default scope:** Reformat in place, preserving all content.

**When the user provides a template/reference document + a target document:**

Don't blindly copy the template's document structure onto the target. A template may have 3 chapters as examples, but the target has 8 — the style system transfers via `template.styles`, the structure does not. Also note: page setup (margins, paper size, headers/footers) is not transferred by `template` — report differences to the user but don't auto-apply.

**Combine layers freely.** A common pattern is: import a school's template for the heading hierarchy + body text, supply user requirements that override one or two specific fields ("摘要部分用楷体"), and use `pattern_rules` for figure/table/reference classification. All three coexist; resolution priority is documented under Step 7.

### Step 2: Inspect the Document

Start by calling `overview`. This gives you the full picture in one call.

Read the overview carefully. Form hypotheses about:
- What type of document is this? (thesis, report, contract, manual, letter...)
- Where does the structural information live — in `styles.xml` definitions, in the content itself, or both?
- How many distinct visual styles exist? Do they map cleanly to semantic roles?

Then use `inspect_*` tools **only as needed** to resolve uncertainties. You don't need to inspect everything. For a simple document, the overview alone may be sufficient. For a complex one, you might drill into 3-5 areas.

**When to reach for `inspect_runs`:** if a paragraph's role is unclear, *before* assuming the dominant-run fingerprint covers everything, dump its runs. The "Run-level diversity" section at the bottom tells you which character properties are uniform vs. mixed — this is what `apply_styles` will preserve vs. strip when you restyle. If you see e.g. `b: on / —` mixed across runs, that's a bold lead phrase that will survive the smart-strip; if you see `b=on` uniform, the bold is style-controllable. Skipping this and assuming uniform formatting is the most common source of post-apply surprises.

**When to reach for `inspect_neighbors`:** classifying figure captions, table captions, abstracts, the first paragraph after a heading — anything that's defined by *what's adjacent* rather than by the paragraph's own format. Pass the candidate's index; check whether image/table appears at distance 1 (or 2 if there's an empty paragraph between). Default radius 4 covers most real layouts; bump with `--radius` if you suspect longer separations.

**When to reach for `find_paragraphs`:** before writing `pattern_rules`, validate the regex matches what you think it matches by running it through this tool first. Also useful for "are there any paragraphs starting with `[N]`?" type discovery questions — much faster than scrolling overview.

**Tool Reference:**

All tools are invoked via `node <script> <args>` and write structured output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | Always. Call this first. Returns metadata, page setup (mm), theme, style definitions, numbering schemes (clustered by pattern), visual style statistics, and document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | When you need full text and computed styles for a specific paragraph range. |
| `inspect_runs` | `node scripts/inspect_runs.js <file> <para>` | When a paragraph has run-level mixed formatting (bold lead phrase, colored numbering prefix, inline emphasis) and you need to see each run's rPr separately. Output also tells you which properties are uniform vs. mixed across runs — critical for predicting what `apply_styles` will preserve vs. strip. |
| `inspect_neighbors` | `node scripts/inspect_neighbors.js <file> <para> [--radius N]` | When you need to know what surrounds a paragraph: nearby images, tables, equations, page breaks, sibling paragraphs. Returns ordered before/after windows with type-specific structured fields. Default radius 4. **First choice for figure-caption / table-caption / first-after-heading classification.** |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | When you see a fingerprint in the overview and need to understand what role it plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | When the document has pre-defined styles in `styles.xml` and you want to understand or preserve them. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | When you need to understand page setup differences between sections (headers, footers, page numbering). |
| `find_paragraphs` | `node scripts/find_paragraphs.js <file> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]` | Cross-document text search. Returns matching paragraphs with index, fingerprint, and text preview. Use to discover content-defined roles (figure/table captions, references, keywords) and to validate `pattern_rules` regex coverage before applying. |
| `apply_styles --dry-run` | `node scripts/apply_styles.js --dry-run <config.json>` | Iterate on a config without writing the output file. Returns the full change report including sample affected paragraphs and the Style Resolution annotation block. Use this between config edits — it's seconds-per-cycle instead of full write+validate. |
| `apply_styles` | `node scripts/apply_styles.js <config.json>` | Final step. Reads a JSON config file containing your complete decision. Outputs the reformatted document and a change report. |

### Step 3: Classify by Visual Fingerprint

Map each visual fingerprint to a semantic role. The overview provides deduplicated fingerprints (A, B, C, ...) with occurrence counts — work at this level, not paragraph-by-paragraph.

For most fingerprints, the mapping is obvious from format + context:
- The largest bold centered text appearing once → `Title`
- Bold text matching a numbering pattern → `Heading1` / `Heading2` / `Heading3`
- The most frequent plain text style → `BodyText`
- Small centered text near images/tables → `FigureCaption` / `TableCaption`
- Monospace font → `Code`

Only drill into individual paragraphs (via `inspect_range` or `inspect_style`) when a fingerprint is ambiguous — e.g. the same visual format is used for both table captions and figure captions, or a format appears in both headings and emphasis text. Common roles:

**Structural roles:**
- `Title` — document title (usually appears once, cover page)
- `Heading1` / `Heading2` / `Heading3` — section headings at different levels
- `HeadingNoNum` — unnumbered section headings (摘要, Abstract, 参考文献, 致谢, 附录)
- `BodyText` — main body paragraphs
- `FirstParagraph` — first paragraph after a heading (if styled differently, e.g. no indent)

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

**Fixed content (do not restyle — preserve as-is):**
- Cover page elements (school name, field labels, date)
- Header/footer content
- Table of contents (auto-generated)

Only create roles that actually exist in the document. If you discover roles not listed here (e.g. `Theorem`, `Proof`, `Definition` in an academic paper), create them.

### Step 4: Define the Style System

For each role, determine formatting parameters. See the `styles` field in the `apply_styles` contract (Step 7) for the full list of available fields.

**Where do parameter values come from?**

Priority order:
1. User's explicit formatting requirements (highest authority)
2. A reference/template document provided by the user
3. The document's existing `styles.xml` definitions (if well-defined)
4. The actual computed values from the document content (extract, don't invent)
5. Reasonable defaults for the document type (last resort only)

**Critical rule: when no user guidelines are provided, extract parameter values directly from the document.** Don't transcribe values by hand — point at a representative paragraph and let the tool extract them.

The `apply_styles` `styles` array supports two modes:

1. **`fromParagraph` mode (preferred when extracting from the document):** Pick one paragraph that exemplifies the role (typically the first occurrence of the dominant fingerprint), set `fromParagraph: <index>`, and the tool will extract the full computed rPr + pPr and use them as the style definition. Use `overrides` to add structural fields the source paragraph lacks (e.g. `outlineLevel` for headings) or to override a specific value the user requested.

2. **Manual mode (when there is no representative paragraph):** Specify each field directly — used when the user provides explicit formatting requirements, when extracting from a template document where you don't have it loaded as the source, or when synthesizing a style for a role with no existing instance.

You can mix modes within the same `styles` array. Example:

```json
{
  "styles": [
    { "id": "BodyText", "name": "正文", "fromParagraph": 60 },
    { "id": "Heading2", "name": "二级标题", "fromParagraph": 19,
      "overrides": { "outlineLevel": 1, "alignment": "left" } },
    { "id": "Caption",  "name": "图表注", "font": "宋体", "size": 10.5,
      "alignment": "center", "lineSpacing": 1.5 }
  ]
}
```

For example:
- Document uses 1.2× line spacing everywhere → use `fromParagraph` on a body paragraph; the tool extracts 1.2× faithfully (not 1.5×)
- Title is left-aligned → `fromParagraph` preserves it; don't override `alignment` to center
- Heading1 appears 5 times, 4 are blue non-bold, 1 is bold non-blue → pass one of the four as `fromParagraph` (the majority), not the outlier

"Normalize" means making inconsistent formatting consistent (by routing all five paragraphs to the same style), NOT replacing the author's chosen formatting with values you think look better.

**What `fromParagraph` extracts:** font, fontEastAsia (only if different from font), size, bold/italic (only if true), color (only if not auto), alignment, spaceBefore, spaceAfter, lineSpacing (with original lineRule preserved), firstLineIndent, hangingIndent, outlineLevel.

**Indent unit preservation:** when the source paragraph used Word's character-based indent (`w:firstLineChars` / `w:hangingChars` — what Word writes for "首行缩进 N 字符"), `fromParagraph` extracts it as `"Nchar"` (e.g. `"2char"`) so the round-trip preserves font-size auto-scaling. When the source used fixed twips (`w:firstLine` / `w:hanging`), it extracts as `"Npt"`. Don't manually convert "char" values to pt — that locks the indent to one font size and breaks downstream font changes.

**Source-run selection within `fromParagraph`:** the tool picks the *dominant text run* — the run carrying the most non-numbering text. Numbering-prefix-only runs (pure digits/dots/parens/bullets/whitespace) are excluded. So `"1.1 数据集导入"` (prefix run + bold title run) extracts the title run's "DengXian 15pt Bold", not the prefix run's "DengXian 15pt blue". You don't need to hand-pick a no-prefix paragraph as the source.

**What `fromParagraph` does NOT extract** (so you must add via `overrides` if needed): `outlineLevel` when the source paragraph has none, `numId`/`numLevel` (always omitted — numbering is bound through the `numbering.levels[].styleId` field, not by hardcoding the source's numId).

**Validation:** `fromParagraph` must reference an indexed paragraph (1-based). Paragraphs inside data tables and form tables are not indexed and cannot be referenced. The tool errors out at the start of execution if the index is invalid.

**Handling inconsistencies between fingerprints of the same role:**

When two fingerprints should map to the same role but differ slightly (e.g. H=18pt blue and F=18pt bold for the same heading level), take the majority pattern's values. Report the normalization in the change report.

**Handling existing styles in the document:**

The document may already define named styles like `Heading1`, `BodyText`, etc. Rules:
- **Style exists and matches your target** → reuse it as-is. Do not re-inject.
- **Style exists but parameters differ from your target** → override it with your target parameters. This updates the existing definition rather than creating a duplicate. But first confirm the style is actually used for its intended role — if `Heading1` is misused as body text throughout the document, overriding it would break those paragraphs. In that case, reassign paragraphs to the correct styles before overriding.
- **Style does not exist** → create it. Use Word built-in IDs (`Heading1`, `Heading2`, `BodyText`, `Caption`) when the role matches, so the document works with Word's TOC, navigation pane, and outline view.
- **Never create a parallel style** like `MyHeading1` when `Heading1` would be correct. This fragments the style system.

When the user provides text requirements, parse Chinese font size names using this mapping:

```
初号 = 42pt    小初 = 36pt
一号 = 26pt    小一 = 24pt
二号 = 22pt    小二 = 18pt
三号 = 16pt    小三 = 15pt
四号 = 14pt    小四 = 12pt
五号 = 10.5pt  小五 = 9pt
六号 = 7.5pt   小六 = 6.5pt
```

### Step 5: Define the Numbering Scheme

If the document has numbered headings (manual or automatic), define a multi-level numbering scheme. Refer to `references/numbering-formats.md` for the full pattern table and `numFmt`/`lvlText` syntax.

Each level binds to a heading style via `pStyle`. When a higher-level heading appears, lower levels reset automatically.

If the document uses manual numbering (just typed text, no `numPr`), convert it to automatic numbering so users no longer need to renumber manually when inserting new sections.

**Mixed manual prefix styles within one role:** authors often mix patterns within the same heading level — e.g. chapter 1's H2 paragraphs are written as "1.1 …", but chapter 2's H2 paragraphs are written as "1. …" (restart per chapter, no chapter prefix). One regex can't normalize both. Use `stripPrefixPatterns: ["%1.%2", "%1."]` on the level — patterns are tried in order, first match wins. Always list the longer pattern first or you'll strip a partial prefix (e.g. `"%1."` would strip just "1." from "1.1 …" leaving ".1 …").

**Preserving design colors on numbers:** if the source document deliberately styles auto/manual numbers in a different color/weight than the title text (very common in design documents — blue numbers + black bold titles), set `numRPr` on the level. The number marker is rendered with this rPr; the rest of the heading uses the paragraph style.

### Step 6: Review Plan Before Execution

Before calling `apply_styles`, present your plan to yourself as a self-check:

1. **Style parameter source** — for each style, are the parameter values taken from `inspect_style` output (or user requirements)? If any value was made up without a source, call `inspect_style` for that fingerprint first.
2. **Style definitions** — list each style with its parameters. Does any style have missing critical fields (e.g. a heading without `outlineLevel`)?
3. **Numbering scheme** — does the `lvlText` pattern match what appears in the document? Are all levels bound to the correct styles?
4. **Fingerprint coverage** — does every fingerprint from the overview have a decision? Each should map to one of: a style (`restyle`), fixed content (`keep`/`exclude`), or uncertain (`flag`). No fingerprint should be left unaccounted for.
5. **Exclude list** — are cover page paragraphs, TOC entries, and other fixed content excluded?
6. **High-risk paragraphs** — are there paragraphs where the role is ambiguous? Flag them rather than guess.

If any of these checks reveal an issue, go back and inspect further before proceeding.

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

  numbering: { levels: [ ... ] },          // optional. Multi-level auto-numbering bound
                                           //   to heading styles, with stripPrefixPatterns
                                           //   (handles mixed manual prefixes) and numRPr
                                           //   (independent rPr for the marker).

  template: { source, styles: [ ... ] },   // optional. Import named styles from another
                                           //   docx; basedOn ancestors auto-pulled,
                                           //   numId references migrated.

  requirements: { id: "原话..." },         // optional. ANNOTATION ONLY — script records
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

**For the full schema** — every field, every option, every comment, plus the complete style-field merge priority and paragraph-mapping resolution order — read `references/apply-styles-config.md` once before composing your first config. The reference is structured by section (Style entries / Numbering / Template import / Paragraph mapping) so you can jump to what you need.

**Quick anchors:**
- Paragraph indexing is 1-based, matching `#001`, `#002` labels in the overview skeleton. Paragraphs inside layout tables are indexed; data/form tables are not.
- Resolution order: `exclude > assignments > pattern_rules > bulk_rules > implicit-keep`. First match wins.
- Style-field priority (later wins): defaults → template-imported → fromParagraph → direct styles[i] fields → overrides. requirements doesn't participate (annotation only).

### Step 8: Validate and Report

**Safety rules (non-negotiable):**

1. The original file is NEVER modified. `apply_styles` copies it first, then modifies the copy.
2. After modification, the tool validates the output file.
3. If validation fails: the output file is discarded, the original is returned unchanged, and the error is reported. Do NOT attempt to manually fix validation errors and retry silently — report the failure to the user.
4. **When to `flag` vs. when to apply:**
   - **Flag:** Structural ambiguity (is this a heading or emphasized text?), numbering pattern unclear (manual prefix doesn't match any known pattern), paragraph could belong to two different roles.
   - **Don't flag:** Minor formatting inconsistency within a clear role (e.g. one heading is 15pt while the rest are 16pt — just normalize it). Repetitive patterns that match an already-classified fingerprint — apply the bulk rule confidently.
   - Principle: flag when the *role assignment* is uncertain, not when the *formatting parameters* have minor variance.
5. Never modify section properties (page size, margins, headers, footers, columns). These are preserved as-is from the original.

**Iterate with `--dry-run` first.** Before committing, run `apply_styles --dry-run <config>` to see the full change report without writing the output. The report includes a per-style sample of the first ~5 affected paragraphs (with text preview, route used, and any prefix stripping) — use this to verify your `bulk_rules` / `pattern_rules` / `requirements` actually targeted the right segments. Adjust the config and re-run until the dry-run looks right, then drop the flag to commit.

**Change report:** `apply_styles` writes the report to stdout as structured text. Present a concise summary to the user in chat, then deliver the output `.docx` file. The report includes:

- Template import summary (if `template` was used): which style IDs imported, basedOn ancestors auto-pulled, numId remappings.
- Styles injected (count and list).
- Paragraphs restyled (count, grouped by style).
- Manual numbering prefixes converted (per pattern).
- Pattern rules matched (per regex, with strip count).
- **Style Resolution** (if `requirements` was provided): per style, "user specified X / agent resolved {Y}" side-by-side. The script doesn't grade — you / the user / a reviewer reads and confirms the translation captures intent. There's no algorithmic check here on purpose; a regex check would silently miss negation, hierarchical refs, and synonyms.
- Sample affected paragraphs per style — use to verify routing.
- Inconsistencies fixed (e.g. "3 headings normalized from 15pt to 16pt").
- Flagged paragraphs (with reasons — the user should review these).
- Validation result (pass/fail).
- If the document contains a TOC: remind the user to right-click the TOC in Word and select "Update Field" after opening.

## Important Guidelines

### What This Skill Does
- Inject named style definitions into `styles.xml`
- Replace direct formatting with style references on paragraphs
- Convert manual numbering to automatic numbering
- Normalize inconsistent formatting (e.g. 15pt → 16pt when majority is 16pt)

### What This Skill Does NOT Do
- Rewrite or rephrase content
- Fix cross-references or field codes
- Correct grammar, punctuation, or spelling
- Rearrange document structure
- Generate new content or sections
- Modify or regenerate table of contents (this skill ensures heading styles have correct `outlineLevel`, but does not update the TOC itself — remind the user to manually update the TOC in Word after opening the output file)
- Restyle paragraphs inside data tables (multi-row/multi-column tables with headers that present structured data — preserve their formatting as-is)
- Modify table structure, cell sizes, borders, or cell-level formatting
- Modify or interpret field codes (`STYLEREF`, `TOC`, `REF`, `DATE`, etc.) — these are preserved as-is

**One exception to "does not edit content":** When converting manual numbering to automatic numbering, the manually typed prefix (e.g. "第1章 " in "第1章 绪论") must be removed from the paragraph text, because the numbering system will now generate it automatically. This is not a content edit — it is migrating formatting information from inline text into the style system. The same applies to manual bullet characters ("• ", "- ") when converting to proper list styles.

### Edge Cases to Watch For

**Mixed-format paragraphs:** A paragraph may contain runs with different roles — e.g. "关键词：" (bold) followed by keyword text (normal), or a list item with a bold lead phrase ("幻觉与安全性 与传统软件不同…"), or a heading with a colored numbering run before a bold title ("1. 开发背景"). This is a single paragraph with character-level style differences, not two roles. Assign the paragraph its role (`Keywords` / `ListNumber` / `Heading1`) and the tool will preserve the cross-run differences automatically: when restyling, only run-level direct formatting that is *uniform across all runs* (i.e. redundant overrides) is stripped — properties that differ between runs are kept as intentional inline emphasis.

**Empty paragraphs as spacing:** Many documents use blank paragraphs for vertical spacing. Always preserve them — removing empty paragraphs is a structural change, not a formatting change, and risks breaking intentional layout (especially on cover pages).

**Table caption position:** Table captions go ABOVE the table. Figure captions go BELOW the figure. Run `inspect_neighbors <para>` and check which side the image/table is on — distance 1 with image before = figure caption, distance 1 with table after = table caption.

**Table footnotes:** Text immediately after a table with smaller font or starting with "注：", "来源：", "Note:" is a table footnote, not body text.

**Unnumbered special headings:** Sections like 摘要, Abstract, 目录, 参考文献, 致谢, 附录 share the visual style of Heading1 but have no chapter number. Create a separate `HeadingNoNum` style or use the same style with numbering suppressed.

**Appendix numbering:** Appendices often restart with a different numbering scheme (附录A, A.1, A.2). This may require a second numbering definition.

**Layout tables vs data tables:** Some documents (e.g. 开题报告表, 申请表) use single-cell tables as bordered content containers — the entire document body lives inside table cells. These are layout tables, not data tables.
- **Layout table** (typically 1 column, or a single merged cell): the paragraphs inside are regular content (headings, body text, lists). Treat them exactly like top-level paragraphs — classify, restyle, include in the skeleton.
- **Data table** (multiple rows × multiple columns, has a header row): presents structured data. Do not restyle content inside. Show only a summary in the skeleton.
- **Form table** (label-value grid, e.g. "姓名：___"): fixed layout for user input. Treat as fixed content (`keep`).
- **Mixed tables:** A single large table may contain layout regions, form fields, and data grids (e.g. 开题报告表 with content sections + evaluation rubric). Classify by region, not by table — the tool may expand some rows and summarize others within the same table.
- The overview tool detects layout tables and expands their content in the skeleton. The Agent does not need to handle this distinction manually — but should verify the tool's classification when inspecting.

## Document Skeleton Format

The `overview` tool outputs document metadata, page setup, theme, style definitions, numbering definitions, a visual style summary, and a document skeleton. The skeleton format:

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
  ...

--- Section 3 (para #29-#68) ---
Footer: Arabic page number (restart from 1)

  #029 [B]  "第1章  绪论"
  #030 [C]  "1.1  研究背景及意义"
  #031 [D]  "本文针对某某问题展开研究..."
  ...
  #049 [E]  "表 3-1 不同方法的性能对比"
  --- TABLE (5×4) headers:["方法","Precision","Recall","F1"] ---
  #050 [F]  "注：加粗数据表示最优结果"
  ...
  --- IMAGE (14cm × 8cm) ---
  #062 [E]  "图 3-2 实验结果对比"
  ...
  #081 --- empty ×3 ---
  ...
```

**Skeleton conventions:**
- `[A]`, `[B]`, etc. are visual fingerprint labels from the summary. The fingerprint hash includes font, size, weight/italic, color, alignment, first-line-indent, AND whether the paragraph carries a numbering reference. So two paragraphs that look visually identical but one is in an auto-numbered list and the other is plain body text will get *different* fingerprints (the listed one will be marked "List" in the summary). This lets `bulk_rules` map list items to a `ListNumber` style without sweeping plain body paragraphs along with them.
- Non-paragraph elements appear as `--- TYPE (details) ---`
- Consecutive empty paragraphs are compressed: `--- empty ×N ---`
- Section breaks show header/footer changes
- Text is truncated to ~40 chars with full text available via `inspect_range`
- **Layout tables** (single-cell content containers) are expanded — their internal paragraphs appear with sequential `#NNN` numbering, indented under `--- LAYOUT TABLE ---` / `--- END LAYOUT TABLE ---` markers
- **Data tables** (multi-column with headers) show only a summary line: `--- TABLE (rows×cols) headers:[...] ---`
- **Form tables** (label-value grids) show a summary: `--- FORM TABLE (rows×cols) ---`

**Example: form-style document (e.g. 学位论文开题报告表)**

```
  #001 [A]  "某某大学"
  #002 [A]  "学位论文开题报告表"
  ...
  #011 [B]  "表格填写要求：正文字体宋体，字号小四，行间距固定值20磅。"
  #012 --- empty ×20 ---
  #032 [C]  "论文概况"
  --- LAYOUT TABLE ---
    #033 [D]  "选题来源：自主命题"
    #034 [D]  "中文摘要："
    #035 [D]  "本研究针对某某问题展开探索..."
  --- END LAYOUT TABLE ---
  #046 [C]  "选题依据"
  --- LAYOUT TABLE ---
    #047 [E]  "选题意义"
    #048 [D]  "近年来某某领域发展迅速..."
    #049 [D]  "然而现有研究仍存在若干不足..."
    ...
    #055 [E]  "国内外研究现状"
    #056 [D]  "该领域的发展经历了多个阶段..."
    #057 [F]  "某某方向的研究进展"
    #058 [D]  "在这一方向上，相关工作主要集中在..."
    ...
  --- END LAYOUT TABLE ---
  ...
  #200 [C]  "评审评语及结论"
  --- TABLE (6×3) headers:["一级指标","二级指标","评价意见"] ---
```

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
    └── numbering-formats.md          ← Numbering format reference (read when handling numbered headings)
```

