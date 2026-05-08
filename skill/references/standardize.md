# Command: `standardize`

Apply style / numbering / structural changes to a Word document. Operates by **role**: you describe a class of paragraphs (every H2, every figure caption, anything matching this fingerprint) and the engine restyles all matching instances.

The sections below describe **three common config shapes**, not exclusive paths — they're points along a density spectrum. Configs are sparse by design: declare only what you're touching, leave the rest implicit.

- **Full Standardization** — broad reshape covering the whole doc, agent owns most style decisions.
- **Targeted Restyle** — narrow scope, only the styles the user named.
- **Escape Hatch** — request can't be expressed via the config; manual XML editing as a last resort.

When intent is genuinely ambiguous, prefer asking one focused question over guessing. When a task spans `standardize` + `edit` (typical for messy templates with content to fill), see SKILL.md "Composing scopes" — `standardize` first, `edit` after.

## Iterating is normal

You don't have to plan the entire transformation upfront and submit one giant config. Apply changes incrementally:

- `apply_styles --dry-run` between edits — seconds per cycle, no file written
- Re-call `inspect_*` between edits to stay grounded in actual document state, not in your earlier mental model
- `styles[]` is **sparse by design** — only declare styles you're touching; untouched styles stay as they are
- Untargeted paragraphs go to **implicit-keep**; that's correct, not a missed decision (unless you're on the Full Standardization path, where coverage is expected)

Small step → dry-run → verify → next step is a fully supported workflow, not a fallback. Use it whenever the change is bounded.

---

## Path: Full Standardization

When the user wants the document brought into a consistent, standardized form. *Illustrative phrasings: "帮我排一下版", "套学校格式", "按这个模板做", or just receiving a docx without specifics.* Same surface phrasing can fall in Targeted Restyle when the user signals narrowed scope; the examples illustrate the concept, they're not triggers.

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

**Structural:** `Title`, `Heading1` / `Heading2` / `Heading3`, `BodyText`. Two less-obvious ones:
- `HeadingNoNum` — unnumbered top-level sections (摘要, Abstract, 参考文献, 致谢, 附录) that share Heading1's visual style but have no chapter number.
- `FirstParagraph` — first paragraph after a heading when styled differently from body (often no first-line indent).

**Caption / label:** `FigureCaption` (below image), `TableCaption` (above table), `TableFootnote` (注/数据来源 lines below a table), `EquationCaption` (right-aligned number).

**Lists:** `ListBullet`, `ListNumber`, `ListContinue` (continuation paragraph within a list item).

**Special:** `Code`, `Quote`, `Reference` ([1] entries with hanging indent), `Keywords`, `Abstract`. Plus one with a tool-specific gotcha:
- `BodyEmphasis` — uniformly-bold body paragraphs acting as in-paragraph sub-titles or labels (no outline level, no auto-numbering). The style **must explicitly set `bold: true`** — when bold is uniform across the paragraph's runs, restyle strips it as redundant direct formatting, so the new style needs to carry it.

**Fixed content (do not restyle — preserve as-is):** cover page elements (school name, field labels, date), header/footer content, table of contents (auto-generated).

Only create roles that actually exist in the document. If you discover roles not listed here (e.g. `Theorem` in math papers), create them.

### Step 4: Define the Style System

**Source of values** (priority order): (1) user requirements, (2) template document, (3) values extracted from a representative paragraph in the document, (4) sensible defaults (last resort). Don't invent values when a representative paragraph exists.

**Two modes per `styles` entry:**

1. **`fromParagraph`** (preferred when extracting from the doc): pick the first occurrence of the dominant fingerprint for the role and set `fromParagraph: <index>`. The tool extracts the full computed rPr + pPr from that paragraph's *dominant text run* (longest non-numbering-prefix run, so `"1.1 研究方法"` extracts the title formatting, not the prefix's). Use `overrides` to add fields the source lacks (e.g. `outlineLevel`) or apply user-requested specifics.
2. **Manual mode**: specify fields directly — when no representative paragraph exists, when synthesizing a role, or when the user fully specified the style.

Modes can mix within one `styles` array. **Prefer overriding existing styles by their actual styleId** when the source has an equivalent role — see "When the document already defines..." below. Only create new styles when no existing one fits.

```jsonc
{
  "styles": [
    // Override existing — `a` is this source's Normal-equivalent (POI default).
    // Discover via `inspect_style_def` first; styleId varies per source.
    { "id": "a", "name": "Normal", "fontCJK": "宋体", "size": 12 },

    // Reuse Word built-in styleId for new roles. `name` uses the canonical
    // English built-in name (Word applies its own UI localization).
    { "id": "Heading2", "name": "heading 2", "fromParagraph": 19,
      "overrides": { "outlineLevel": 1 } },

    // Custom-named role — `name` doesn't alias any built-in identity.
    { "id": "Caption",  "name": "Caption", "fontCJK": "宋体", "size": 10.5,
      "alignment": "center", "lineSpacing": 1.5 }
  ]
}
```

**Normalization rule:** for outliers (e.g. Heading1 appears 5 times, 4 of one pattern + 1 different), source from the majority. The same applies when two fingerprints play the same role — take the majority's values, route both fingerprints' paragraphs to one style. "Normalize" means routing inconsistent paragraphs to one consistent style, NOT replacing the author's choices with values you think look better.

**What `fromParagraph` extracts:** fontLatin, fontCJK (only if different from fontLatin), size, bold/italic (only if true), color (only if not auto), alignment, spaceBefore, spaceAfter, lineSpacing (with original lineRule preserved), firstLineIndent, hangingIndent, outlineLevel (only when the source has it set — add via `overrides` if you need it on a heading whose source paragraph lacks it).

**Does NOT extract:** `numId` / `numLevel` — numbering is bound through `numbering.levels[].styleId`, not hardcoded per paragraph.

**`fontLatin` vs `fontCJK`:** `fontLatin` is the Latin / Western text slot (writes to OOXML `w:ascii` and `w:hAnsi`); `fontCJK` is the East-Asian slot (`w:eastAsia`). A paragraph rendering as Chinese still has a Latin font that would render any Latin characters in the same paragraph — `fromParagraph` extracts both when they differ. When the user names only a CJK font ("正文宋体" / "标题黑体"), set `fontCJK` and leave `fontLatin` unset so the source's Latin font is preserved. Set both only when the user explicitly says the same font should apply to Latin too. (`inspect_range` and friends use the same field names — `fontCJK` / `fontLatin` — so what you read out of inspect is what you paste into the config.)

**Indent unit preservation:** when the source used Word's character-based indent (`w:firstLineChars` / `w:hangingChars`, what Word writes for "首行缩进 N 字符"), extraction gives `"Nchar"` so font-size auto-scaling round-trips. Fixed twips give `"Npt"`. Don't manually convert "char" values to pt — that locks the indent to one font size.

**Override existing styles before creating new ones.** Run `inspect_style_def` to discover what the source already has — POI / WPS / school templates often play the role of Normal / Heading 1 / etc. under short auto-generated styleIds (`a`, `a1`, `2`, `10`, ...). Override by their exact styleId: upsertStyle mutates in place, preserving everything you didn't specify (basedOn, default="1", numPr, link, etc.). This is the safest path because it avoids the most common rendering trap — name collision (see below). Override is **non-destructive on unmanaged properties**: if you don't specify `basedOn`, the existing `basedOn` is preserved. Only the fields you specify are written. Verify the style is actually used for its intended role first — overriding `Heading1` while it's misused as body text would corrupt those paragraphs; reassign the paragraphs first.

**`name` must not alias any existing style's identity.** Word treats `<w:name>` as the built-in style identity marker, including locale aliases ("Normal" ≡ "正文" ≡ "標準"; "Heading 1" ≡ "标题 1"; "Body Text" ≡ "正文文本"). When two different styleIds claim the same identity, Word silently drops the second style's `rPr` at render time — your fonts vanish, scripts see no error, dry-run reports look correct. Two safe approaches:

- **Override existing by its styleId** so no new name enters the doc.
- **When creating new with a built-in styleId** (`Heading1`, `BodyText`, `Caption`, ...), use the canonical English built-in name matching the styleId — `name: "Body Text"` for `id: "BodyText"`, `name: "heading 1"` for `id: "Heading1"`, `name: "Caption"` for `id: "Caption"`. Word applies its own UI localization for display.
- **When creating new with a custom styleId** (e.g. `BodyEmphasis`, `MyCallout`), the simplest safe rule is `name = id` (e.g. `name: "BodyEmphasis"`). Any name that doesn't match a built-in style's English name or its locale aliases works; matching the styleId is the easiest way to satisfy that without checking the alias table.

Using a localization that doesn't match the styleId (e.g., `id: "BodyText"` with `name: "正文"` — that's Normal's localization, not BodyText's) is wrong AND collision-prone. The engine catches direct string-equal collisions plus the major en/zh-CN aliases at preflight; other locales rely on this rule.

Use Word built-in IDs (`Heading1` / `Heading2` / `BodyText` / `Caption`) for new styles when the role matches, so TOC / nav / outline view work; never create parallel styles like `MyHeading1`.

**Chinese font size names** (初号/一号/.../小六): see `references/chinese-font-sizes.md` for the pt mapping when the user specifies sizes in Chinese terms.

**Three layers for setting fonts. Decide by understanding what the user is expressing — the example phrases below are illustrations, not keyword triggers.**

- **Per-role override** (`styles[]` entries with `fontCJK` / `fontLatin` / `overrides`). Use when the user is talking about *specific roles* and you can name which ones. Effect: targeted, only the styled paragraphs change. Illustrative phrasings: "标题黑体, 正文宋体", "图注小五号", "Heading2 字号改小一号", "参考文献保持原样".

- **Whole-doc default** (declare a `Normal` entry in `styles[]`; other styles' `basedOn: "Normal"` chain inherits from it). Use when the user is expressing a uniform default for everything that doesn't have a more specific role decided — no role distinction in their phrasing. Effect: wide; covers most pStyle-bound paragraphs via the cascade. Illustrative phrasings: "整篇统一用 Times New Roman", "全文宋体小四" without further role specification.

- **Document-design font scheme** (`theme.fonts` block). Use when the user is expressing intent at the document's *design layer* — they expect Word's "+正文" / "+标题" UI entries to reflect new fonts, they're talking about the template's font scheme, or they want defaults to propagate even into chrome / future-edited content. Effect: widest; updates theme1.xml directly so any docDefaults / styles / runs that reference theme fonts auto-resolve to the new values. Illustrative phrasings: "把这份文档的主题字体改成 X/Y", "更新文档的字体方案".

The same surface phrase ("把字体改成宋体") can fall into any of the three depending on what the user actually means; ask yourself what they'd expect to change after the operation, not which words appeared. Bias toward the narrowest layer that captures the intent — wider layers risk surprising the user with effects on chrome they didn't intend to touch.

`theme.fonts` and `styles[]` overrides compose cleanly: theme sets the document baseline, styles override specific roles on top.

### Step 5: Define the Numbering Scheme

When the document has typed heading prefixes (`"1. 引言"` / `"1.1 研究方法"` / `"第N章 ..."`), migrate to automatic numbering — this is part of standardization. Skip only when the user explicitly opts out, the source already has real `numId` references you want to preserve (verify with `inspect_range` — typed-text prefixes look identical to auto-numbers but behave totally differently), or no numbered headings exist.

If the manual scheme itself is inconsistent across the document — e.g. H1 has numbers in chapter 1 but not chapter 2, or H2 uses chapter-prefixed `"1.1"` in some chapters and per-chapter-restart `"1."` in others — auto-migration is also a normalization decision that may change author-intended semantics. Ask the user before applying rather than picking one scheme silently.

Each level binds to a heading style via `styleId`; higher levels reset lower-level counters automatically. Field names mirror OOXML: `numFmt` (e.g. `decimal` / `chineseCounting` / `bullet`), `lvlText` (the rendered prefix pattern, e.g. `"%1."` / `"%1.%2"` / `"第%1章"`), and `suff` controls the gap between the marker and the paragraph text — `"space"` when the marker ends in a digit or character (`1. Title`, `第一章 研究方法`, `1.1 概述`), `"nothing"` when the trailing punctuation already separates them (`一、研究方法`, `（一）背景`), `"tab"` only for wide-list layouts. Specify `suff` explicitly per level; the engine infers it from trailing spaces in `lvlText` when omitted, but that's a tolerant fallback rather than the intended path. See `references/numbering-formats.md` for full value tables and ready-made templates. Minimal example for three-level decimal headings:

```jsonc
"numbering": {
  "levels": [
    { "level": 0, "numFmt": "decimal", "lvlText": "%1.",     "suff": "space", "styleId": "Heading1" },
    { "level": 1, "numFmt": "decimal", "lvlText": "%1.%2",   "suff": "space", "styleId": "Heading2" },
    { "level": 2, "numFmt": "decimal", "lvlText": "%1.%2.%3", "suff": "space", "styleId": "Heading3" }
  ]
}
```

**`outlineLevel` is independent of numbering level.** They co-occur for headings but are different OOXML concepts: `outlineLevel` (set on the paragraph style's pPr) controls TOC inclusion, navigation pane, outline view. Numbering level (the `level` field above) controls auto-number display. List styles (`ListBullet` / `ListNumber`) take numbering but **must not** carry `outlineLevel` — otherwise list items pollute the TOC. Set `outlineLevel` explicitly on heading styles via `overrides`; don't expect numbering to imply it.

**`stripPrefixPatterns` — when to specify, when to skip.** When omitted, `stripPrefixPatterns` defaults to `[lvlText]` — so for the simple case where the doc's manual prefix matches the new auto-numbering pattern (`lvlText: "%1."` strips "1." / "2." / etc.), **you don't need to write it**. Specify it only when the source mixed prefix styles within one role — e.g. chapter 1's H2s are "1.1 ..." while chapter 2's are "1. ..." (restart per chapter). Use `stripPrefixPatterns: ["%1.%2", "%1."]` — patterns tried in order, first match wins; **longer pattern must come first** or `"%1."` will strip just "1." from "1.1 ..." leaving ".1 ...".

The `%N` placeholder in a pattern matches Arabic digits or the common Chinese numerals (一二三...百千). It does NOT match Roman numerals, Latin letters, less-common Chinese forms (壹貳叄, 万, 億), circled / parenthesised digits (①, ⒈), or other locale-specific shapes. The dry-run report's "Numbered-style paragraphs not matched by any stripPrefixPattern" section surfaces unstripped paragraphs with leading-text samples — read those samples: if you see an unrecognised shape (like `"附录 A 实验数据"` or `"叁、研究方法"`), bypass `stripPrefixPatterns` for those cases by adding a `pattern_rules` entry with `stripMatch: true` and an explicit regex covering the shape.

The dry-run report also flags the well-handled mixed-pattern case as "Mixed manual numbering detected" once your `stripPrefixPatterns` cover both shapes, so you can confirm with the user before final write.

**Preserving design colors on numbers:** if the source styles numbers in a different color/weight than title text (e.g. blue numbers + black bold titles), set `numRPr` on the level. The marker is rendered with this rPr; the title uses the paragraph style.

### Step 6: Review Plan Before Execution

Before calling `apply_styles`, self-check:

1. **Style values have sources** — every parameter came from user spec, template, or `inspect_style` extraction. None were invented.
2. **Heading styles have `outlineLevel`** — required for TOC / nav / outline view. Word's built-in `Title` style has no `outlineLevel` by default; only set one if the user's template treats Title as part of the heading hierarchy (e.g., a thesis where Title is H0 above H1 chapters).
3. **Numbering migrated** when source has typed heading prefixes (per Step 5). `stripPrefixPatterns` covers mixed variants within a role.
4. **Every fingerprint has a decision** — restyle / keep / exclude / flag. No fingerprint left unaccounted for. *(This coverage rule applies to Full Standardization only — Targeted Restyle and Audit paths do not require it.)*

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

  theme: { fonts: { majorLatin?, majorEastAsia?, minorLatin?, minorEastAsia? } },
                                           // optional. Document-level font scheme override.
                                           //   Use for whole-doc design intent ("全文宋体",
                                           //   "把这份文档的主题字体改成 X / Y"). Modifies
                                           //   theme1.xml so the "+正文"/"+标题" entries
                                           //   in Word's font dropdown show the new fonts.
                                           //   For role-specific changes prefer styles[]
                                           //   overrides. See Step 4 § "Three layers for setting fonts".

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

Full schema in `references/config-schema.md` — read once before composing your first config.

Config-shape invariants (cross-command invariants like file safety, paragraph indexing, paths-resolve, restyle uniform-strip live in SKILL.md and apply here unchanged):
- Paragraph mapping order (first match wins): `exclude > assignments > pattern_rules > bulk_rules > implicit-keep`.
- Style-field priority (later wins): defaults → template-imported → fromParagraph → direct fields → overrides.

### Step 8: Validate and Report

Iterate with `apply_styles --dry-run` first. The change report has several sections worth scanning before committing:

- **Style Resolution** — every injected style listed with user spec (if `requirements` set) next to agent-resolved fields. Read for translation correctness.
- **Paragraphs untouched** — split into "empty (likely spacers)" and "non-empty (verify coverage)". On the Full Standardization path, an unfamiliar count under non-empty means a fingerprint slipped through; on Targeted Restyle, both are expected.
- **Manual numbering converted / Mixed manual numbering detected** — if the source mixed numbering schemes within one role (a real and common case), the report calls this out. Treat it as a normalization decision worth confirming with the user.
- **Sample Affected Paragraphs** — first N restyled per style, with prefix-stripping notes inline. Use these to spot-check that bulk_rules / pattern_rules hit the right targets.

**Safety guarantees:** see SKILL.md "Cross-command invariants" — original-never-modified, validation-or-discard, section properties untouched all apply unchanged.

**When to `flag` vs. apply:** flag when the *role assignment* is genuinely uncertain (could be heading or emphasized body, prefix doesn't match any known pattern, ambiguous between two roles). Don't flag formatting variance within a clear role (one heading is 15pt while the rest are 16pt — just normalize it).

**Hand-off:** present a concise summary of the change report to the user, then deliver the output. If the document contains a TOC, remind the user to right-click → "Update Field" in Word after opening.

---

## Path: Targeted Restyle

The user is expressing focused changes with the rest of the document expected to stay untouched. The decisive signal is *scope narrowing* — they either name a specific change, or ask that something be preserved — not a particular phrase. *Illustrative phrasings: "加个 X 样式 / 其他不动", "Heading2 字号改小一号", "把所有 [N] 开头的段落统一缩进", "保留手动编号，只调字体". A request that doesn't read as narrow on the surface but expects narrow effect lands here too.*

**Mindset:** small, focused, additive. Locate target → minimal config → dry-run → verify → apply. Don't try to classify the whole document.

### Workflow

1. **Locate target paragraphs.**
   - Content-pattern targets: `find_paragraphs --regex` is the right tool — it returns matching paragraphs with index, fingerprint, and text preview, without the overhead of a full overview.
   - Visual targets ("just the H2s"): `inspect_style <fingerprint>` if you can identify the fingerprint, or run a quick `overview` to spot it.
   - Single-paragraph targets: `inspect_range <para> <para>` for full computed style.

2. **Decide the target style.**
   - **Reuse an existing styleId** (check via `inspect_style_def`) when possible — preserves whatever's already wired (TOC links, outline level, basedOn chain). Just override the fields the user wants changed.
   - **Define a new style** only when no existing one fits: `fromParagraph` extracts from a representative paragraph; or specify fields directly. See § Step 4 (Define the Style System) for extraction details.

3. **Write a sparse config.** Pick the narrowest tool that fits the change:
   - **`restyle`** when the change is paragraph style assignment only — same config as `apply_styles` minus `template` / `numbering`. Most common Targeted Restyle case.
   - **`migrate_numbering`** when only adding / replacing a numbering scheme. `styles[]` can be empty if you're binding to heading styles already in the doc.
   - **`import_template`** when only pulling in template styles. Often chained: `import_template` → `restyle` (apply the imported styles to paragraphs).
   - Use the unified `apply_styles` when the change spans multiple operations.

   In all cases: `styles[]` should be sparse. **Don't redeclare untouched styles** — that's noise that risks accidental overrides. Prefer `pattern_rules` (content-based) or targeted `assignments` (specific paragraphs by index). Avoid broad `bulk_rules` unless a fingerprint cleanly captures your target — fingerprint changes feel cheap but can spread.

4. **Dry-run.** The change report should touch only paragraphs you intended. Untouched fingerprints appear under "implicit-keep" — that's correct behavior on this path, not a missing decision.

5. **Apply.**

### Notes specific to this path

- **No fingerprint coverage requirement.** Step 6's coverage check is a Full Standardization rule. Here, untargeted paragraphs simply stay as they are.
- **Run-level direct formatting on untouched paragraphs is preserved.** The uniform-strip rule fires only on paragraphs the script restyles; everything else keeps its inline formatting.
- **Iterate.** Apply one change, dry-run, verify, then layer the next change on the result. Two small configs are easier to debug than one large one.
- **When a request grows beyond "targeted":** if you find yourself adding 5+ styles, declaring numbering, or reaching for the template — switch to Full Standardization. Targeted Restyle's discipline is keeping the change small.

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

## Path: Extend the style system (prep for content fill)

When the user is going to fill a template with content, but the template's existing styles don't cover what the content needs, run `standardize` first to install the missing pieces. Then hand off to `edit`.

Triggered from the planning survey (SKILL.md Core Principle): content depth exceeds the doc's existing Heading levels, or the doc has typed structural prefixes (chrome or pre-populated content) that should become real auto-numbering, or content shape lacks a list-bound / caption / quote / code style.

What to install:

- **Heading levels** — new `Heading3` / `Heading4` etc. in `styles[]` with `basedOn` chained to the existing Heading2 so font and paragraph defaults inherit. Step down visual emphasis sensibly; don't invent a wholly new scheme.
- **Unified multi-level numbering** — one `numbering` config covering every Heading level the doc needs (existing levels, levels you're adding, and the chrome levels the template typed by hand). Declare `stripPrefixPatterns` matching the manual shapes so the script removes them during restyle. `references/numbering-formats.md` has multi-level templates for academic / technical / governmental / legal shapes.
- **List-bound style + separate single-level numbering** — for body lists, distinct from the heading scheme.
- **Other body styles** — caption, quote, code — as needed.

After install, dry-run to confirm the Style Resolution block looks right, apply, and the doc is ready for `edit` content placement.

### Heading numbering vs list numbering — distinct treatments

Two kinds of manual structural prefixes show up in real docs; convert each correctly:

- **Heading-class** — outline-bearing prefixes that define document structure. Depth typically 1–4 levels. Convert to `Heading1` / `Heading2` / ... bound to **one unified multi-level scheme**.
- **List-class** — local enumerations inside body sections. Usually single-level, restart per group. Convert to `ListNumber` / `ListBullet` bound to a **separate single-level scheme**.

Detection signals (visible facts; agent classifies):

- Position: heading sits at section start, followed by short title text; list item sits inside a section among other list items or body prose.
- Surrounding format: heading already styled distinctly (bold, larger size); list items sit in body-style.
- Pattern shape: headings tend to decimal hierarchy or chapter sentinels; lists tend to parenthesized or single-level digits.
- Depth: headings nest 2–4 levels; lists usually 1.

Rare ambiguous cases (a typed `第N章` inside body prose that's a citation rather than a heading; a numbered fragment that could be enumeration or sub-heading) — these are fallback Ask cases per SKILL.md.

---

## Reading the Overview Output

`overview` prints visual style summary + document skeleton inline; the format is self-explanatory once you see it. Four conventions worth knowing in advance, since they aren't visible from the output alone:

- **Letter vs hash labels.** `[A]`, `[B]` are sorted by frequency in this run (volatile across edits); the summary also shows a 6-char content hash next to each letter (`A [c4f9]: ...`). `bulk_rules.fingerprint` accepts either — use letters for in-session iteration, hashes in configs you intend to keep across doc revisions (hashes are stable when paragraphs are added/removed and frequency-rank shifts).
- **Numbered ≠ unnumbered fingerprints.** The fingerprint hash includes whether a paragraph carries a numbering reference, so visually identical paragraphs split into separate fingerprints (suffix "List") when one is auto-numbered and the other isn't. `bulk_rules` can target them independently.
- **Layout vs data tables.** Layout tables (single-cell content containers) are inlined into the skeleton between `--- LAYOUT TABLE ---` markers; data and form tables are summarized as a single non-paragraph block. Verify the classification when inspecting unfamiliar table-heavy documents.
- **Empty paragraphs and truncation.** Consecutive empty paragraphs are compressed (`--- empty ×N ---`); paragraph text is truncated to ~40 chars in the skeleton — use `inspect_range` for the full text.

---

## Edge Cases to Watch For

- **Empty paragraphs as spacing**: preserve them. Removing is structural, not stylistic, and risks breaking cover-page layout.
- **Table caption vs figure caption**: table captions go ABOVE the table, figure captions BELOW the figure. Use `inspect_neighbors` to confirm which side the image/table is on.
- **Table footnotes**: text right after a table starting with "注：" / "来源：" / "Note:" is a footnote, not body text.
- **Unnumbered special headings** (摘要 / Abstract / 目录 / 参考文献 / 致谢 / 附录): share the visual style of Heading1 but have no chapter number. Use `HeadingNoNum` or suppress numbering on the same style.
- **Appendix numbering** often restarts with a different scheme (附录A / A.1 / A.2) — may need a second `numbering` entry.
- **Layout vs data tables**: the overview tool classifies these — layout tables (single-cell content containers) are inlined into the skeleton; data/form tables are summarized. Verify the classification when inspecting unfamiliar table-heavy documents.
- **Pre-printed chrome (forms, templates with printed labels and instructions)**: when the visual summary shows a long tail of low-occurrence fingerprints with short average text length (`avg ≤20ch`), those are usually printed labels / cover chrome, not author content. Step 6's coverage rule applies to *author-content* fingerprints (high count + meaningful `avg ch`); chrome is fixed content by default. Distinguish via the avg length and the `via "name"/id` annotation on the visual summary line — content fingerprints typically share one explicit pStyle, chrome scatters across the default.
- **Source's base style violates the document's stated specs**: e.g. the doc's `Normal` (or a custom base like `a1` "段落") sets `bold: true` while the printed instructions say 正文不加粗. Three options:
  1. **Override the base style** (declare it in `styles[]` with the corrected fields) — fixes everything that inherits, including untouched chrome. Use when chrome correctness matters and inherited side-effects on chrome are acceptable.
  2. **Define a new role-specific style and bulk_rule content fingerprints to it** — leaves the broken base alone, only your targeted paragraphs get the right style. Chrome stays broken. Use when chrome layout is fragile and shouldn't be touched.
  3. **Override the base AND assign chrome paragraphs to a separate fixed style** — most thorough, most config. Use for final-print-ready output.

  Verify the choice via the dry-run Style Resolution block before committing.

## Compose with other commands

- After installing the style system here, `edit` for surgical touch-ups on specific paragraphs the rules missed.
- Mixed input (messy template + content to fill): `standardize` first → `edit` second. Filling directly with `edit` propagates the template's bad styles via Match-Destination-Formatting.
- Read-only check before reshape: `audit`. The audit's violation list often translates directly into a `standardize` config.
