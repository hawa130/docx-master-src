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

## Single-Counter Patterns (Captions, References, …)

For paragraph roles that need a sequence but not a hierarchy — figure captions, table captions, reference list entries, theorem numbering, appendix items — bind a one-level numbering scheme.

Default counter scope: one continuous counter across the document (`restart: "continuous"`, implicit). Every paragraph bound to the scheme shares one running counter — what captions, references, equations, and appendix items want. The only opt-out is procedural `1./2./3.` list shapes (see "Procedural Numbered Lists" below), which need `restart: "perInstance"`.

### Figure / Table Caption (flat counter)
```
Level 0: 图 1 / 图 2 / 图 3            numFmt=decimal  lvlText="图 %1"   suff="space"
Level 0: 表 1 / 表 2 / 表 3            numFmt=decimal  lvlText="表 %1"   suff="space"
```
Use a separate `numId` per caption family so figures and tables count independently.

Pair the caption style with a `FigureImage` body style (`alignment: "center"`, small `spaceBefore` / `spaceAfter` to butt the image against its caption) and set the image block's `styleId: "FigureImage"` — without it the image paragraph emits no pPr and renders left-aligned with default spacing regardless of what the caption style declares.

**Position convention.** `FigureCaption` paragraphs sit **below** the image. `TableCaption` paragraphs sit **above** the `{ "type": "table", ... }` block. `EquationNumber` paragraphs sit to the **right** of the equation on the same line (compose via 3-column borderless `TableBlock`; see [`equations.md`](equations.md)). All bind to single-level continuous counters (default `restart`); body text refs target the caption / number paragraph's `anchor`, never the figure / table / equation paragraph itself.

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

### Procedural Numbered Lists (`restart: "perInstance"`)
```
Level 0: 1. / 2. / 3.                   numFmt=decimal  lvlText="%1."  suff="space"  restart="perInstance"
```
When the same scheme is meant to drive multiple separate list blocks (steps in Chapter 1 numbered 1./2./3., independent steps in Chapter 2 also starting from 1.), opt in to per-instance restart. The engine forks a fresh `numId` per contiguous run of paragraphs bound to the scheme and writes `<w:startOverride val="1"/>` on each fork. An "instance" is broken by any paragraph not bound to the scheme's styleId — e.g. a heading or body paragraph between two list blocks. Use sparingly; only this case actually wants it.

## lvlText Syntax

- `%N` = counter at level N (1-indexed)
- Composite: `%1.%2` → "1.2"; literal wraps variables: `第%1章` → "第1章"
- `isLgl: true` on a level forces every cross-level `%N` to render as arabic regardless of the referenced level's `numFmt` (use when an outer level is chineseCounting / roman but you want its digit form here)

