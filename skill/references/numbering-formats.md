# Numbering Format Reference

## OOXML numFmt Values

| `numFmt` value | Output | Notes |
|---|---|---|
| `decimal` | 1, 2, 3... | |
| `upperLetter` | A, B, C... | |
| `lowerLetter` | a, b, c... | |
| `upperRoman` | I, II, III... | |
| `lowerRoman` | i, ii, iii... | |
| `chineseCounting` | 一, 二, 三... | Chinese ordinal |
| `chineseCountingThousand` | 一, 二, 三... | Same glyphs, different internal encoding |
| `ideographTraditional` | 甲, 乙, 丙... | |
| `bullet` | • | For unordered lists |
| `none` | (nothing) | Suppress number display |

## Common Multi-Level Heading Patterns

`suff` controls the marker→text gap: `"space"` when the marker ends in a digit/character (`1. Title`), `"nothing"` when trailing punctuation already separates (`一、Title`). `"tab"` is rare, wide-list layouts only.

### Academic Thesis (Chinese)
```
Level 0: 第1章  / 第2章  / 第3章       numFmt=chineseCounting  lvlText="第%1章"     suff="space"
Level 1: 1.1   / 1.2    / 2.1          numFmt=decimal          lvlText="%1.%2"      suff="space"
Level 2: 1.1.1 / 1.1.2  / 2.1.1       numFmt=decimal          lvlText="%1.%2.%3"   suff="space"
```

### Technical Document (Decimal)
```
Level 0: 1.    / 2.     / 3.           numFmt=decimal          lvlText="%1."        suff="space"
Level 1: 1.1   / 1.2    / 2.1          numFmt=decimal          lvlText="%1.%2"      suff="space"
Level 2: 1.1.1 / 1.1.2  / 2.1.1       numFmt=decimal          lvlText="%1.%2.%3"   suff="space"
```

### Government Document (Chinese)
```
Level 0: 一、   / 二、   / 三、        numFmt=chineseCounting  lvlText="%1、"       suff="nothing"
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%2）"     suff="nothing"
Level 2: 1.     / 2.     / 3.          numFmt=decimal          lvlText="%3."        suff="space"
Level 3: （1）  / （2）  / （3）       numFmt=decimal          lvlText="（%4）"     suff="nothing"
```

### Mixed Counting Schemes (Chinese outer, Arabic inner)
```
Level 0: 一、   / 二、   / 三、        numFmt=chineseCounting  lvlText="%1、"       suff="nothing"
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%2）"     suff="nothing"
Level 2: 1.1    / 1.2    / 2.1         numFmt=decimal          lvlText="%1.%3"      suff="space"  isLgl=true
```
A level mixing arabic counters with cross-references to a non-arabic outer level needs `isLgl=true`. Without it, Word renders each `%N` placeholder using *that level's* `numFmt` — so `%1.%3` on level 2 above would display `一.1`, not `1.1`. `isLgl` overrides cross-level placeholders to arabic regardless of the referenced level's format.

### Legal Document
```
Level 0: 第一条 / 第二条 / 第三条      numFmt=chineseCounting  lvlText="第%1条"     suff="space"
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%2）"     suff="nothing"
Level 2: 1.     / 2.     / 3.          numFmt=decimal          lvlText="%3."        suff="space"
```

## Single-Counter Patterns (Reference List, Procedural Lists, …)

For paragraph roles that need a sequence but not a hierarchy — reference list entries, procedural numbered steps, appendix items — bind a one-level numbering scheme. **Caption-class numbering (figures / tables / equations / theorems) lives in the top-level `captions` table, not here — see [`captions.md`](captions.md).**

Default counter scope: one continuous counter across the document
(`restart: "continuous"`, implicit). The only opt-out is procedural
`1./2./3.` list shapes (see "Procedural Numbered Lists" below), which
need `restart: "perInstance"`.

### Reference List
```
Level 0: [1] / [2] / [3]                numFmt=decimal  lvlText="[%1]"   suff="space"
```
Body-text cites to these (or to any auto-numbered caption / heading) go through `InlineRef` in `edits[]`, not literal text. See [`cross-references.md`](cross-references.md).

### restart

Scheme-level `restart` controls how the counter sequence resets across the document. Default is `"continuous"` (one counter for the whole doc).

| Value | Behavior |
|---|---|
| `"perInstance"` | Each contiguous run of list items gets its own numId; restart at run boundaries (broken by any non-target paragraph). |
| `"continuous"` | One numId across the whole doc; items continue regardless of intervening paragraphs (default). |
| `"byHeading"` | Restart whenever the nearest preceding heading-styled paragraph (any style with `outlineLvl`) changes. |
| `{ "atStyleChange": "ProposalH2" }` | Restart whenever a paragraph bound to the named styleId appears. |

Block-level `numbering: { numId, level, restart: true }` overrides scheme-level behavior at one paragraph. The engine forks a fresh numId with `<w:startOverride val="1"/>` at that point; use when a single mid-list position needs a hard reset the scheme-level value can't express.

## Explicit `numId` on a scheme

By default, the engine allocates fresh numIds for declared schemes. To pin a scheme to a specific id — so block-level `numbering: { numId }` references resolve predictably — set `"numId": N` on the scheme. The dry-run report includes a scheme → numId allocation table showing which id each scheme was assigned and whether it was pinned or allocated.

Collision: two schemes requesting the same `numId` cause apply to throw, naming both conflicting entries.

## lvlText Syntax

- `%N` = counter at level N (1-indexed)
- Composite: `%1.%2` → "1.2"; literal wraps variables: `第%1章` → "第1章"
- `isLgl: true` on a level forces every cross-level `%N` to render as arabic regardless of the referenced level's `numFmt` (use when an outer level is chineseCounting / roman but you want its digit form here)

