# Numbering Format Reference

## OOXML numFmt Values

| `numFmt` value | Output | Notes |
|---|---|---|
| `decimal` | 1, 2, 3... | Most common |
| `upperLetter` | A, B, C... | |
| `lowerLetter` | a, b, c... | |
| `upperRoman` | I, II, III... | |
| `lowerRoman` | i, ii, iii... | |
| `chineseCounting` | 一, 二, 三... | Chinese ordinal |
| `chineseCountingThousand` | 一, 二, 三... | Same glyphs, different internal encoding |
| `ideographTraditional` | 甲, 乙, 丙... | Rare |
| `bullet` | • | For unordered lists |
| `none` | (nothing) | Suppress number display |

## Common Multi-Level Heading Patterns

`suff` is shown alongside `lvlText` because the marker→text gap is part
of the visual identity of each pattern. Pick `suff="space"` when the
marker ends in a digit or character (the gap separates them from the
title text), `"nothing"` when the marker's trailing punctuation already
provides the separation. `"tab"` is rare and only suits wide-list
layouts that align titles in a column.

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

## Single-Counter Patterns (Captions, References, …)

For paragraph roles that need a sequence but not a hierarchy — figure captions, table captions, reference list entries, theorem numbering, appendix items — bind a one-level numbering scheme. Never let the agent emit the counter as literal text (`图 1`, `[1]`); it desyncs the moment a figure is inserted or a reference reordered.

### Figure / Table Caption (flat counter)
```
Level 0: 图 1 / 图 2 / 图 3            numFmt=decimal  lvlText="图 %1"   suff="space"
Level 0: 表 1 / 表 2 / 表 3            numFmt=decimal  lvlText="表 %1"   suff="space"
```
Use a separate `numId` per caption family so figures and tables count independently.

### Chapter-Prefixed Caption (`图 1-1`, `表 1-1`, `图 2-1`, ...)
```
Level 0: (mirrors Heading1's counter; no display)  numFmt=decimal  lvlText=""        suff="nothing"
Level 1: 图 1-1 / 图 1-2 / 图 2-1                  numFmt=decimal  lvlText="图 %1-%2" suff="space"
```
Captions sit on level 1; level 0 silently tracks the chapter number, restart-on-Heading1 via the chapter style's own numbering. Bind caption paragraphs to level 1; the body never sees level 0. For tables use the same shape with `lvlText="表 %1-%2"` on a separate `numId` so figures and tables count independently.

### Reference List
```
Level 0: [1] / [2] / [3]                numFmt=decimal  lvlText="[%1]"   suff="space"
```
Body-text cites to these (or to any auto-numbered caption / heading) go through `InlineRef` in `edits[]`, not literal text. See [`cross-references.md`](cross-references.md).

## lvlText Syntax

- `%1` = current value of level 1
- `%2` = current value of level 2
- `%1.%2` = "1.2" (composite reference to multiple levels)
- Literal text wraps the variables: `第%1章` = "第1章"
- Each level can only reference its own level and higher levels
- `isLgl: true` on a level forces every cross-level `%N` to render as arabic, regardless of the referenced level's `numFmt` (use when an outer level is chineseCounting / roman / etc. but you want it to appear as a digit inside this level's marker)

