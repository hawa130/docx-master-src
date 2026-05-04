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

## docx-js Implementation

```javascript
// In numbering config:
{
  reference: "heading-numbering",
  levels: [
    {
      level: 0,
      format: LevelFormat.DECIMAL,           // numFmt
      text: "%1",                             // lvlText
      alignment: AlignmentType.LEFT,
      start: 1,                               // w:start
      style: {
        paragraph: {
          indent: { left: 0, hanging: 0 }
        }
      }
    },
    {
      level: 1,
      format: LevelFormat.DECIMAL,
      text: "%1.%2",
      alignment: AlignmentType.LEFT,
      start: 1,
      style: {
        paragraph: {
          indent: { left: 0, hanging: 0 }
        }
      }
    }
  ]
}

// In style definition, bind numbering:
{
  id: "Heading1",
  paragraph: {
    numbering: { reference: "heading-numbering", level: 0 }
  }
}
```

## LevelFormat Constants (docx-js)

| LevelFormat constant | Equivalent numFmt |
|---|---|
| `LevelFormat.DECIMAL` | `decimal` |
| `LevelFormat.UPPER_LETTER` | `upperLetter` |
| `LevelFormat.LOWER_LETTER` | `lowerLetter` |
| `LevelFormat.UPPER_ROMAN` | `upperRoman` |
| `LevelFormat.LOWER_ROMAN` | `lowerRoman` |
| `LevelFormat.BULLET` | `bullet` |
| `LevelFormat.CHINESE_COUNTING` | `chineseCounting` |
| `LevelFormat.CHINESE_COUNTING_THOUSAND` | `chineseCountingThousand` |

## Manual Numbering Detection Patterns

When documents use typed numbering instead of `numPr`, detect these regex patterns:

```
/^第[一二三四五六七八九十\d]+章\s/     → Level 0, Chinese chapter
/^第\d+章\s/                          → Level 0, Arabic chapter  
/^(\d+)\s/                            → Level 0, plain decimal
/^(\d+)\.(\d+)\s/                     → Level 1, decimal.decimal
/^(\d+)\.(\d+)\.(\d+)\s/              → Level 2, three-level decimal
/^[一二三四五六七八九十]+、/            → Level 0, Chinese with 、
/^（[一二三四五六七八九十]+）/          → Level 1, Chinese in parens
/^(\d+)\.\s/                          → Varies, decimal with dot
/^（(\d+)）/                           → Varies, Arabic in parens
/^[①②③④⑤⑥⑦⑧⑨⑩]/                   → Deepest level, circled numbers
/^附录\s*[A-Z]/                        → Appendix heading
/^[A-Z]\.(\d+)\s/                      → Appendix sub-section (A.1, B.2)
```

**Important:** These patterns help the Agent recognize manual numbering. The Agent should not rely solely on regex — it should also consider font size, boldness, and position to confirm that a numbered paragraph is actually a heading vs. a list item vs. body text that happens to start with a number.
