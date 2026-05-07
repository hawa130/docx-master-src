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

### Academic Thesis (Chinese)
```
Level 0: 第1章  / 第2章  / 第3章       numFmt=chineseCounting  lvlText="第%1章"
Level 1: 1.1   / 1.2    / 2.1          numFmt=decimal          lvlText="%1.%2"
Level 2: 1.1.1 / 1.1.2  / 2.1.1       numFmt=decimal          lvlText="%1.%2.%3"
```

### Technical Document (Decimal)
```
Level 0: 1     / 2      / 3            numFmt=decimal          lvlText="%1"
Level 1: 1.1   / 1.2    / 2.1          numFmt=decimal          lvlText="%1.%2"
Level 2: 1.1.1 / 1.1.2  / 2.1.1       numFmt=decimal          lvlText="%1.%2.%3"
```

### Government Document (Chinese)
```
Level 0: 一、   / 二、   / 三、        numFmt=chineseCounting  lvlText="%1、"
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%1）"
Level 2: 1.     / 2.     / 3.          numFmt=decimal          lvlText="%1."
Level 3: （1）  / （2）  / （3）       numFmt=decimal          lvlText="（%1）"
Level 4: ①     / ②     / ③           numFmt=decimal          lvlText=special
```

### Legal Document
```
Level 0: 第一条 / 第二条 / 第三条      numFmt=chineseCounting  lvlText="第%1条"
Level 1: （一） / （二） / （三）      numFmt=chineseCounting  lvlText="（%1）"
Level 2: 1.     / 2.     / 3.          numFmt=decimal          lvlText="%1."
```

## lvlText Syntax

- `%1` = current value of level 1
- `%2` = current value of level 2
- `%1.%2` = "1.2" (composite reference to multiple levels)
- Literal text wraps the variables: `第%1章` = "第1章"
- Each level can only reference its own level and higher levels

