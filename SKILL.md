---
name: docx-normalize
description: "Normalize Word document formatting: fix inconsistent styles, convert direct formatting to named styles, convert manual numbering to automatic numbering. Use whenever a user wants to standardize, clean up, or reformat a .docx file's layout and styles."
---

# Document Formatting Normalizer

Read a user's Word document, identify the semantic role of every element (headings, body text, captions, lists, code blocks, etc.), inject standardized named styles, replace direct formatting with style references, and convert manual numbering into automatic numbering. Output a reformatted copy of the original document plus a change report.

## Core Principle

**You are the analyst. The tools are your instruments.**

Tools only present facts — computed styles, element positions, document structure. They never classify or judge. All semantic reasoning is yours: deciding what role a paragraph plays, what style to name it, whether two similar formats should merge or stay separate, and how to handle edge cases.

## Workflow

### Step 1: Understand the Goal

Check what the user has provided. Only ask for clarification if the goal is genuinely ambiguous — otherwise proceed with defaults and state your assumptions in the report.

- **Explicit guidelines provided?** The user may provide formatting requirements in text ("一级标题三号黑体加粗, 正文小四宋体1.5倍行距") or as a separate reference document. If so, these are the authoritative style definitions — the document's actual formatting is secondary.
- **No guidelines?** Infer the intended style system from the document itself, identify inconsistencies, and normalize to the majority pattern.
- **Default scope:** Reformat in place, preserving all content.

**When the user provides a template/reference document + a target document:**

This is a two-phase task. Keep the phases separate:
1. **Extract style definitions from the template.** Run `overview` on the template first. Extract the style system (format parameters for each role, numbering scheme). The template defines *what styles should look like*. Note: page setup (margins, paper size, headers/footers) is not transferred — only report differences to the user.
2. **Classify the target document.** Run `overview` on the target. Identify what role each paragraph plays. The target defines *which paragraphs get which styles*.

Do not blindly copy the template's document structure onto the target. A template may have 3 chapters as examples, but the target has 8 — the style system transfers, the structure does not.

### Step 2: Inspect the Document

Start by calling `overview`. This gives you the full picture in one call.

Read the overview carefully. Form hypotheses about:
- What type of document is this? (thesis, report, contract, manual, letter...)
- Where does the structural information live — in `styles.xml` definitions, in the content itself, or both?
- How many distinct visual styles exist? Do they map cleanly to semantic roles?

Then use `inspect_*` tools **only as needed** to resolve uncertainties. You don't need to inspect everything. For a simple document, the overview alone may be sufficient. For a complex one, you might drill into 3-5 areas.

**Tool Reference:**

All tools are invoked via `node <script> <args>` and write structured output to stdout.

| Tool | Invocation | When to Use |
|------|------------|-------------|
| `overview` | `node scripts/overview.js <file>` | Always. Call this first. Returns metadata, page setup, theme, style definitions, numbering schemes, visual style statistics, and document skeleton. |
| `inspect_range` | `node scripts/inspect_range.js <file> <from> <to>` | When you need full text and computed styles for a specific paragraph range. |
| `inspect_style` | `node scripts/inspect_style.js <file> <fingerprint>` | When you see a fingerprint in the overview and need to understand what role it plays across the document. |
| `inspect_style_def` | `node scripts/inspect_style_def.js <file> <styleId>` | When the document has pre-defined styles in `styles.xml` and you want to understand or preserve them. |
| `inspect_section` | `node scripts/inspect_section.js <file> <index>` | When you need to understand page setup differences between sections (headers, footers, page numbering). |
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
4. The majority pattern in the document content (most common format for each role)
5. Reasonable defaults for the document type

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

### Step 6: Review Plan Before Execution

Before calling `apply_styles`, present your plan to yourself as a self-check:

1. **Style definitions** — list each style with its parameters. Does any style have missing critical fields (e.g. a heading without `outlineLevel`)?
2. **Numbering scheme** — does the `lvlText` pattern match what appears in the document? Are all levels bound to the correct styles?
3. **Fingerprint coverage** — does every fingerprint from the overview have a decision? Each should map to one of: a style (`restyle`), fixed content (`keep`/`exclude`), or uncertain (`flag`). No fingerprint should be left unaccounted for.
4. **Exclude list** — are cover page paragraphs, TOC entries, and other fixed content excluded?
5. **High-risk paragraphs** — are there paragraphs where the role is ambiguous? Flag them rather than guess.

If any of these checks reveal an issue, go back and inspect further before proceeding.

### Step 7: Execute

Call `apply_styles` with your complete decision.

**Input contract:**

```
apply_styles({

  source: "/path/to/original.docx",    // REQUIRED. Path to the uploaded file.
                                         // This file is NEVER modified.

  output: "/path/to/output.docx",       // REQUIRED. Path for the new file.
                                         // Must differ from source.

  styles: [                              // REQUIRED. At least one style.
    {
      id:              "Heading1",       // REQUIRED. Style ID for styles.xml.
                                         //   Use Word built-in IDs when possible
                                         //   (Heading1, Heading2, BodyText, Caption, etc.)
      name:            "一级标题",        // REQUIRED. Display name.
      basedOn:         "Normal",         // optional. Parent style ID. Default: "Normal".
      font:            "黑体",           // optional. Latin/ASCII font.
      fontEastAsia:    "黑体",           // optional. CJK font. Default: same as font.
      size:            16,               // optional. Font size in pt (not half-pt).
      bold:            true,             // optional. Default: false.
      italic:          false,            // optional. Default: false.
      color:           "auto",           // optional. Hex ("2E75B6") or "auto". Default: "auto".
      alignment:       "left",           // optional. "left"|"center"|"right"|"both". Default: "left".
      lineSpacing:     1.5,              // optional. Multiple (1.0, 1.5, 2.0) or exact pt value.
      spaceBefore:     12,               // optional. Space before paragraph in pt.
      spaceAfter:      6,                // optional. Space after paragraph in pt.
      firstLineIndent: "2char",          // optional. "Nchar" or pt value.
      hangingIndent:   null,             // optional. Hanging indent in pt.
      outlineLevel:    0,                // optional. 0-8. REQUIRED for heading styles (enables TOC).
    },
    ...
  ],

  numbering: {                           // optional. Omit if document has no numbered headings/lists.
    levels: [
      {
        level:   0,                      // REQUIRED. 0-8.
        format:  "chineseCounting",      // REQUIRED. numFmt value.
        text:    "第%1章",               // REQUIRED. lvlText pattern.
        styleId: "Heading1",             // REQUIRED. Binds this level to a style.
        start:   1,                      // optional. Starting number. Default: 1.
      },
      ...
    ]
  },

  assignments: [                         // optional. Per-paragraph overrides (1-based para index).
    { para: 1,  action: "keep" },
    { para: 33, action: "restyle", style: "Heading1" },
    { para: 47, action: "flag",    reason: "Ambiguous: looks like heading but no numbering" },
    ...
  ],

  bulk_rules: [                          // optional. Apply style by visual fingerprint label.
    { fingerprint: "D", style: "BodyText" },
    { fingerprint: "E", style: "Code" },
    ...
  ],

  exclude: [1, 2, 3, 4, 5]              // optional. Para indices to never touch. Overrides everything.
})
```

**Paragraph indexing:** 1-based, matching `#001`, `#002` labels in the overview skeleton. Paragraphs inside layout tables are included in the numbering. Paragraphs inside data tables and form tables are not indexed.

**Resolution order** when multiple rules match the same paragraph:
1. `exclude` — if listed here, paragraph is untouched. Full stop.
2. `assignments` — if a per-paragraph rule exists, it wins.
3. `bulk_rules` — if no per-paragraph rule, match by fingerprint.
4. No match — paragraph is left unchanged (implicit `keep`).

**Assignment actions:**
- `keep` — preserve original formatting exactly, do not apply any style.
- `restyle` — apply the named style, remove conflicting direct formatting. `style` field is REQUIRED.
- `flag` — do not modify the paragraph. Record in the change report with the `reason` string. `reason` field is REQUIRED.

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

**Change report:** `apply_styles` writes the report to stdout as structured text. Present a concise summary to the user in chat, then deliver the output `.docx` file. The report includes:

- Styles injected (count and list)
- Paragraphs restyled (count, grouped by style)
- Inconsistencies fixed (e.g. "3 headings normalized from 15pt to 16pt")
- Flagged paragraphs (with reasons — the user should review these)
- Validation result (pass/fail)
- If the document contains a TOC: remind the user to right-click the TOC in Word and select "Update Field" after opening

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

**Mixed-format paragraphs:** A paragraph may contain runs with different roles — e.g. "关键词：" (bold) followed by keyword text (normal). This is a single paragraph with a character-level style difference, not two roles. Assign the paragraph its role (`Keywords`) and note the run-level formatting in the style definition.

**Empty paragraphs as spacing:** Many documents use blank paragraphs for vertical spacing. Always preserve them — removing empty paragraphs is a structural change, not a formatting change, and risks breaking intentional layout (especially on cover pages).

**Table caption position:** Table captions go ABOVE the table. Figure captions go BELOW the figure. Use successor/predecessor element types to distinguish them, even if they share identical formatting.

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
- `[A]`, `[B]`, etc. are visual fingerprint labels from the summary
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
│   ├── inspect_style.js              ← Visual fingerprint occurrences
│   ├── inspect_style_def.js          ← Named style definition details
│   ├── inspect_section.js            ← Section page setup details
│   └── apply_styles.js               ← Execute formatting changes, output new file + report
└── references/
    └── numbering-formats.md          ← Numbering format reference (read when handling numbered headings)
```

