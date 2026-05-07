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
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%1）"     suff="nothing"
Level 2: 1.     / 2.     / 3.          numFmt=decimal          lvlText="%1."        suff="space"
Level 3: （1）  / （2）  / （3）       numFmt=decimal          lvlText="（%1）"     suff="nothing"
Level 4: ①     / ②     / ③           numFmt=decimal          lvlText=special      suff="nothing"
```

### Legal Document
```
Level 0: 第一条 / 第二条 / 第三条      numFmt=chineseCounting  lvlText="第%1条"     suff="space"
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%1）"     suff="nothing"
Level 2: 1.     / 2.     / 3.          numFmt=decimal          lvlText="%1."        suff="space"
```

## lvlText Syntax

- `%1` = current value of level 1
- `%2` = current value of level 2
- `%1.%2` = "1.2" (composite reference to multiple levels)
- Literal text wraps the variables: `第%1章` = "第1章"
- Each level can only reference its own level and higher levels

