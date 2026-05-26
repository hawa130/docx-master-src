# Archetype: Chinese government / institutional form templates

## Pattern signature

These templates (申报书 / 申请表 / 报告表 / 备案表) share a distinctive structure:

- A cover page using 1×1 or 2×2 layout tables (project title, classification, dates) — title text often holds the most informative single piece of metadata
- An 8×2 / 3×2 form table for cover-level fields (单位 / 联系人 / 日期 / 编号)
- Personnel / budget data tables (multi-column, fixed row count)
- **One big content table** (typically 16×2 or 8×2 with cells merged across cols) — this is where 90% of the substantive text lives. Each row is a section ("研究内容" / "技术路线" / "进度安排" / etc.). Identified by:
  - Total para count in one cell ≥ 5
  - Or contains heading-bound paragraphs (outlineLvl)
  - Or rows have one merged-`gridSpan` cell each (single tc per row regardless of grid columns)

The classifier (see `references/tables.md`) recognizes these big content tables as `layout`, so their paragraphs participate in global #NNN indexing — surgical edits work as on body paragraphs.

## Typical workflow

1. `overview` — confirm classifier picked up the big content table as `layout`. If the table renders as `--- TABLE ... ---` instead of `--- LAYOUT TABLE ---`, paragraphs inside are unindexed and unreachable for paragraph/range locators.
2. `inspect_table <T>` — survey data/form tables to know what cells exist. Multi-paragraph cells show `paras:N` with `K1: ... K2: ...` per-paragraph snippets.
3. `audit` — typically reveals:
   - Direct pPr `outlineLvl` on heading paragraphs that should cascade from a styleId
   - Manual typed prefixes (`一、` / `（一）` / `1.1`) inside heading text
   - Inconsistent font use across personnel cells
4. `apply` config: standardize-shape with:
   - Strip manual prefixes (`stripPrefixPatterns`)
   - Install ProposalH1/H2/H3 styles bound to a `chineseCountingThousand` scheme (set explicit `numId` to make block-level refs deterministic)
   - Normalize body paragraphs to a single ProposalBody style
   - Fill personnel rosters via `cell` locator (use `paragraph: K` narrowing if rows have multi-paragraph cells)
   - Fill the big content table sections via `paragraph` locator (now indexed) or `cell` locator with paragraph: K

## Common pitfalls

- **Cover 1×1 table containing project title**: classifier treats 1×1 tables as layout by default, so the title becomes indexed. Don't try to address it via `cell` locator; use the global paragraph index.
- **MDF dropping outline level**: when replacing a paragraph with `styleId: ProposalH1`, the engine strips anchor's direct `outlineLvl` / `numPr` / `pageBreakBefore` / `widowControl` / `spacing` so the styleId cascade reaches them. No more "Heading 1 styled paragraph rendered as Heading 2".
- **Caption iteration blocked by existing SEQ fields**: after a first apply emits captions, subsequent `replace` is blocked. Use `replace.overwriteFields: true` to regenerate.
- **Personnel roster cells**: `cell` locator with optional `paragraph: K` for per-paragraph reach. `RunLocator` accepts cell coords for run-level edits inside data cells.
- **Numbering restart across chapters**: use scheme-level `restart: "byHeading"` or `{ atStyleChange: "ProposalH2" }` rather than per-instance forks across multiple styles.
- **Validation errors carried through from source**: templates from Word / WPS / LibreOffice / POI often have pre-existing OOXML validation warnings (VML in footers, `numId="0"` sentinels, mc:AlternateContent extensions). These are non-fatal by default (baseline-diff). Use `--allow-validation-warnings` only if apply introduces NEW errors you want to keep around for debugging.

## Reference fixtures

- `/Users/geek-tech/hawa130/report/烟草局/fill-project/申报书模板.docx` — empty template
- `/Users/geek-tech/hawa130/report/烟草局/fill-project/基于参数高效微调的烟草领域知识持续学习模型研究与应用.docx` — filled

## Source of these recommendations

`/Users/geek-tech/hawa130/report/烟草局/fill-project/docx-master-feedback.md` — feedback from a 20-iteration apply workflow on this exact archetype.
