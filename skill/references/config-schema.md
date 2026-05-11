# `apply` Config Schema

Full field-by-field reference for the JSON config consumed by
`apply [--dry-run] <config.json>`. Read this before composing your first
config; SKILL.md only carries a top-level summary.

## Top-level shape

```jsonc
{
  source: "/path/to/original.docx",  // REQUIRED. Path to the input file.
                                     // This file is NEVER modified — apply
                                     // copies it first, then writes the modified copy.
  output: "/path/to/output.docx",    // REQUIRED. Path for the new file. Must differ from source.
  dryRun: false,                     // optional. true = in-memory pipeline + report, no file written.
                                     // Equivalent to the --dry-run CLI flag.

  template: { ... },                 // optional. See § "Template import" below.
  theme:    { fonts: {...} },        // optional. See § "Theme" below.
  styles:   [ ... ],                 // REQUIRED. See § "Style entries" below.
  numbering: { levels: [...] }       // optional. Single scheme, OR array of
            | [ {levels:[...]}, ... ], // schemes for parallel installations
                                     // (e.g. multi-level heading + single-level list).
  requirements: { id: "..." },       // optional. Annotation only — see § "Requirements" below.

  // Paragraph-to-style mapping, in resolution order:
  exclude:        [1, 2, 3],         // array of paragraph indices (numbers, not objects);
                                     //   see § "Paragraph mapping"
  assignments:    [ ... ],
  pattern_rules:  [ ... ],
  bulk_rules:     [ ... ],
}
```

## Style entries

`styles[]` is required and holds at least one entry. Each entry follows one of two shapes (you can mix freely across the array):

### Mode A — extract from a representative paragraph (preferred)

```jsonc
{
  id:            "Heading1",   // REQUIRED. Style ID written into styles.xml.
                               // Use Word built-ins (Heading1, BodyText, Caption, ...)
                               // when the role matches — keeps TOC / nav pane / outline
                               // working without extra wiring.
  name:          "一级标题",    // REQUIRED. Display name shown in Word's style panel.
  basedOn:       "Normal",     // optional. Default "Normal".
  fromParagraph: 33,           // 1-based paragraph index. The tool extracts that
                               // paragraph's full computed rPr + pPr (using the
                               // dominant text run, skipping numbering-prefix-only
                               // runs) and uses them as the style definition.
  overrides: {                 // optional. Any field listed under Mode B can appear here.
    outlineLevel: 0,           // typical use: add structural fields the source lacks
    alignment:    "left",      // or override a specific value per user request
  },
}
```

### Mode B — define manually

```jsonc
{
  id:              "Caption",  // REQUIRED.
  name:            "图表注",    // REQUIRED.
  basedOn:         "Normal",   // optional. Default "Normal".
  fontLatin:       "Arial",    // optional. Latin / Western text font (writes
                               //   to OOXML w:ascii AND w:hAnsi).
  fontCJK:         "黑体",     // optional. CJK font (w:eastAsia). Most common
                               //   field in Chinese-academic configs.
  size:            10.5,       // optional. pt (not half-pt).
  bold:            false,      // optional. Default false.
  italic:          false,      // optional. Default false.
  color:           "auto",     // optional. Hex ("2E75B6") or "auto". Default "auto".
  vertAlign:       "superscript", // optional. "superscript" | "subscript" | "baseline".
                               //   Used on character styles (FootnoteReference,
                               //   EndnoteReference, or custom super/sub styles).
                               //   "baseline" is the explicit reset — distinct from
                               //   omitting the field (= inherit cascade).
  alignment:       "center",   // optional. "left" | "center" | "right" | "both".
  lineSpacing:     1.5,        // optional. Number or "Npt" string (e.g. 20 or "20pt").
                               //   Number: <10 → multiplier (auto rule); ≥10 → pt (exact rule).
                               //   "Npt" string: always pt (exact), regardless of magnitude.
  lineRule:        "atLeast",  // optional. "auto" | "exact" | "atLeast". Overrides the
                               //   default rule. Use "atLeast" to faithfully round-trip a
                               //   source's atLeast rule.
  spaceBefore:     12,         // optional. pt before paragraph.
  spaceAfter:      6,          // optional. pt after paragraph.
  firstLineIndent: "2char",    // optional. "Nchar" / "Npt" / pt number / null.
                               //   "Nchar" → emitted as `w:firstLineChars` (1/100 char),
                               //     auto-scales with run font size — required for the
                               //     standard "首行缩进 2 字符" academic convention.
                               //   "Npt" or number → emitted as `w:firstLine` (fixed
                               //     twips), does NOT scale with font.
                               //   `null` (or omitted) → no indent emitted; for
                               //     existing paragraphs the cascade decides.
                               //   `0` / `"0pt"` → explicitly emit zero-indent
                               //     (overrides an inherited indent to nothing).
                               //   Prefer "Nchar" for thesis/paper body text.
  hangingIndent:   null,       // optional. Same units as firstLineIndent. For
                               //   bibliography/reference entries use "2char" (or
                               //   pt to match the leading "[N] " marker width).
  outlineLevel:    0,          // optional. 0–9 per OOXML §17.3.1.20: 0–8 are
                               //   heading levels (0 = H1, 1 = H2, …); 9 = body
                               //   text (no outline). Set on heading styles to
                               //   enable TOC / outline view / nav pane.
}
```

Modes mix in the same array. `overrides` companions `fromParagraph` — layer additions / replacements on top of the extracted shape.

## Numbering

```jsonc
numbering: {
  levels: [
    {
      level:   0,                        // REQUIRED. 0-8.
      numFmt:  "chineseCounting",        // REQUIRED. OOXML w:numFmt value — decimal /
                                         //   chineseCounting / bullet / lowerRoman / etc.
                                         //   See references/numbering-formats.md for table.
      lvlText: "第%1章",                 // REQUIRED. OOXML w:lvlText pattern.
      styleId: "Heading1",               // REQUIRED. Binds this level to a paragraph style.
      start:   1,                        // optional. Starting number. Default 1.
      stripPrefixPatterns: ["%1.%2", "%1."],
                                         // optional. Alternative manual-prefix patterns
                                         //   to strip from paragraph text (tried in
                                         //   order, longest first). Use when the source
                                         //   mixes styles — some H2 written as "1.1 ...",
                                         //   others as "1. ...". Defaults to [lvlText].
      suff: "nothing",                   // "tab" | "space" | "nothing". Inserted between
                                         //   the auto-generated marker and the paragraph
                                         //   text. **Specify per level** — `"space"` when
                                         //   the marker ends in a digit / character
                                         //   (`1. Title`, `第一章 …`), `"nothing"` when
                                         //   trailing punctuation already separates
                                         //   (`一、…`, `（一）…`). Word's spec default is
                                         //   "tab", which renders an ugly gap unsuitable
                                         //   for CJK; we don't fall back to it.
                                         // If omitted, suff is inferred from trailing
                                         //   whitespace in `lvlText` (0 → "nothing",
                                         //   1 → "space", 2+ → "tab") and that whitespace
                                         //   is stripped from the emitted lvlText. The
                                         //   inference is a tolerant fallback for
                                         //   imprecise input — explicit `suff` is the
                                         //   intended path.
      numRPr: {                          // optional. rPr applied to the auto-generated
        color: "3370FF",                 //   number marker only — independent of the
        bold:  false,                    //   title text. Use to keep designs like
                                         //   "blue numbering + black title".
      },
    },
    ...
  ]
}
```

Omit `numbering` entirely if the document has no numbered headings/lists.

## Requirements (annotation only)

```jsonc
requirements: {
  BodyText: "正文请使用宋体小四号字，行距设为1.5倍，首行缩进两个字符",
  Heading1: "标题用黑体三号加粗居中显示",
}
```

The script does NOT parse this. It records each string and displays it side-by-side with the agent-resolved structured fields under "=== Style Resolution ===" in the change report. The agent (you) translates the natural-language spec into the structured `styles[i]` fields above — Chinese typography requires understanding negation ("不要加粗"), hierarchical references ("比一级小一号"), Chinese numerals ("两个字符"), and synonyms (思源黑体 / 苹方 / 方正小标宋), which a fixed regex parser silently fails on.

The Style Resolution display is the verification mechanism — agent / user / reviewer reads both lines and confirms the translation matches intent.

## Template import

```jsonc
template: {
  source: "/path/to/template.docx",  // path to the reference docx
  styles: ["BodyText", "Heading1", "Heading2", "Heading3", "Caption"],
                                     // styleIds to import. basedOn ancestors are
                                     // auto-pulled if they don't exist in source.
  importNumbering: true,             // default true. When an imported style references
                                     // a numId, the corresponding abstractNum + num
                                     // are migrated to fresh IDs in the source's
                                     // numbering.xml. The imported style's numPr is
                                     // rewritten to point at the new numId.
}
```

If a styleId already exists in source, the template's definition wins — the template is treated as the authoritative stylebook for the listed IDs. Other source styles are untouched.

## Theme

```jsonc
theme: {
  fonts: {
    majorLatin:    "Times New Roman",  // optional. Heading/Title Latin font (Office major scheme).
    majorEastAsia: "黑体",              // optional. Heading/Title CJK font.
    minorLatin:    "Times New Roman",  // optional. Body Latin font (Office minor scheme).
    minorEastAsia: "宋体",              // optional. Body CJK font.
  },
}
```

Modifies `word/theme/theme1.xml`. Any `docDefaults` / `styles[]` / direct rPr that references theme fonts (`<w:rFonts w:asciiTheme="majorHAnsi"/>` etc.) auto-resolves to the new values — the document-design layer instead of per-style declarations.

Use when the user wants the doc's underlying font scheme changed ("把这份文档的主题字体改成 X / Y"), not when they want one specific style restyled. Sparse: declare only the slots being changed; omitted slots keep the source's existing theme value.

## Paragraph mapping

Paragraphs are 1-based, matching `#001`, `#002` labels in the overview skeleton. Paragraphs inside layout tables are included in the numbering; paragraphs inside data tables and form tables are not indexed.

```jsonc
exclude: [1, 2, 3],                    // never touch — overrides everything else
assignments: [                          // per-paragraph rules (highest precedence
  { para: 1,  action: "keep" },         //   among the non-exclude branches)
  { para: 33, action: "restyle", style: "Heading1" },
  { para: 47, action: "flag",    reason: "Ambiguous: looks like heading but no numbering" },
],
pattern_rules: [                       // regex-based by paragraph text
  { regex: "^图\\s*\\d+[-.]\\d+", style: "FigureCaption" },
  { regex: "^表\\s*\\d+[-.]\\d+", style: "TableCaption" },
  { regex: "^\\[\\d+\\]\\s+",     style: "Reference" },
  { regex: "^(关键词|Keywords?)\\s*[:：]", style: "Keywords", stripMatch: true },
  // First match wins (rules tried in order). The match must be anchored at the
  // start of paragraph text. `stripMatch: true` removes the matched leading text —
  // useful when the new style replaces the label via numbering or other mechanism.
],
bulk_rules: [                          // by visual fingerprint label from overview
  { fingerprint: "D", style: "BodyText" },
  { fingerprint: "E", style: "Code" },
],
```

### Resolution order

For each paragraph, the first matching branch wins:

1. `exclude` — paragraph is untouched. Full stop.
2. `assignments` — per-paragraph rule.
3. `pattern_rules` — first regex match anchored at start of text.
4. `bulk_rules` — match by fingerprint label.
5. No match — paragraph is left unchanged (implicit `keep`).

### Assignment actions

- `keep` — preserve original formatting exactly. `style` not used.
- `restyle` — apply the named style and strip conflicting direct formatting. `style` is required.
- `flag` — do not modify the paragraph; record it in the change report with the supplied `reason`. Use when the role assignment is genuinely ambiguous.

## Style-field resolution priority

For each entry's final field values, layered low → high (later wins):

1. Defaults.
2. `template` imported style (if same ID).
3. `styles[i].fromParagraph` extracted values.
4. `styles[i]` direct fields (fontLatin / fontCJK / size / bold / ...).
5. `styles[i].overrides` — deliberate per-style escape.

`requirements` is annotation only and does not participate in resolution.
