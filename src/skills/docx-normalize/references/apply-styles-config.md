# `apply_styles` Config Schema

Full field-by-field reference for the JSON config consumed by
`apply_styles [--dry-run] <config.json>`. Read this before composing
your first config; SKILL.md only carries a top-level summary.

## Top-level shape

```jsonc
{
  source: "/path/to/original.docx",  // REQUIRED. Path to the input file.
                                     // This file is NEVER modified — apply_styles
                                     // copies it first, then writes the modified copy.
  output: "/path/to/output.docx",    // REQUIRED. Path for the new file. Must differ from source.
  dryRun: false,                     // optional. If true, runs the entire pipeline
                                     // in memory and prints the report but does NOT
                                     // write the output file or run post-write
                                     // validation. Equivalent to passing --dry-run
                                     // on the CLI. Use during config iteration.

  template: { ... },                 // optional. See § "Template import" below.
  styles:   [ ... ],                 // REQUIRED. See § "Style entries" below.
  numbering:{ levels: [...] },       // optional. See § "Numbering" below.
  requirements: { id: "..." },       // optional. Annotation only — see § "Requirements" below.

  // Paragraph-to-style mapping, in resolution order:
  exclude:        [1, 2, 3],         // see § "Paragraph mapping"
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
  font:            "黑体",     // optional. Latin/ASCII font.
  fontEastAsia:    "黑体",     // optional. CJK font. Default: same as font.
  size:            10.5,       // optional. pt (not half-pt).
  bold:            false,      // optional. Default false.
  italic:          false,      // optional. Default false.
  color:           "auto",     // optional. Hex ("2E75B6") or "auto". Default "auto".
  alignment:       "center",   // optional. "left" | "center" | "right" | "both".
  lineSpacing:     1.5,        // optional. <10 → multiplier (auto rule); ≥10 → pt.
  lineRule:        "atLeast",  // optional. "auto" | "exact" | "atLeast". Overrides
                               //   the <10/≥10 heuristic. Use "atLeast" to faithfully
                               //   round-trip a source's atLeast rule.
  spaceBefore:     12,         // optional. pt before paragraph.
  spaceAfter:      6,          // optional. pt after paragraph.
  firstLineIndent: "2char",    // optional. "Nchar" / "Npt" / pt number.
                               //   "Nchar" → emitted as `w:firstLineChars` (1/100 char),
                               //     auto-scales with run font size — required for the
                               //     standard "首行缩进 2 字符" academic convention.
                               //   "Npt" or number → emitted as `w:firstLine` (fixed
                               //     twips), does NOT scale with font.
                               //   Prefer "Nchar" for thesis/paper body text.
  hangingIndent:   null,       // optional. Same units as firstLineIndent. For
                               //   bibliography/reference entries use "2char" (or
                               //   pt to match the leading "[N] " marker width).
  outlineLevel:    0,          // optional. 0-8. Set on heading styles to enable
                               //   TOC / outline view / nav pane.
}
```

You can mix modes in the same array. Mode A is preferred when a representative paragraph already exists in the source — `fromParagraph` extracts the full computed rPr+pPr faithfully (with the indent-unit and dominant-run rules described above). Use Mode B (or `overrides`) when the source has no clean exemplar or when the user spec calls for values the source doesn't currently have.

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
4. `styles[i]` direct fields (font / size / bold / ...).
5. `styles[i].overrides` — deliberate per-style escape.

`requirements` is annotation only and does not participate in resolution.
