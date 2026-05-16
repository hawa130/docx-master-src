/**
 * Shared constants for the MathML → OMML converter.
 *
 * Sources:
 *   - n-ary chr table: ECMA-376 Part 1 §22.1.2.70 (CT_NaryPr/chr) plus
 *     the operator class assignments in MathML Core Appendix B (Operator
 *     Dictionary). Cross-checked against TEI Stylesheets' mml2omml.xsl
 *     (BSD-2 / CC-BY-SA 3.0, Copyright 2011–2020 TEI Consortium).
 *   - Cambria Math font: Word's hard-coded math font; every <m:r> wears
 *     it regardless of declared family. Matches Word's own export.
 */

export const MML_NS = "http://www.w3.org/1998/Math/MathML"
export const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

/** Operators that promote a surrounding <mrow> into <m:nary>. The string
 *  value is what we emit for `<m:chr m:val="…">`; if the symbol matches
 *  Word's default for that nary (∑/∏/∫), Word lets us omit chr — but we
 *  always emit it for explicitness (matches Word's own round-tripping).
 *  Keys are Unicode codepoints as decimal numbers for fast lookup. */
export const NARY_OPERATORS: ReadonlyMap<string, string> = new Map([
  ["∑", "∑"], // U+2211 N-ARY SUMMATION
  ["∏", "∏"], // U+220F N-ARY PRODUCT
  ["∐", "∐"], // U+2210 N-ARY COPRODUCT
  ["∫", "∫"], // U+222B INTEGRAL
  ["∬", "∬"], // U+222C DOUBLE INTEGRAL
  ["∭", "∭"], // U+222D TRIPLE INTEGRAL
  ["∮", "∮"], // U+222E CONTOUR INTEGRAL
  ["∯", "∯"], // U+222F SURFACE INTEGRAL
  ["∰", "∰"], // U+2230 VOLUME INTEGRAL
  ["⨀", "⨀"], // U+2A00 N-ARY CIRCLED DOT
  ["⨁", "⨁"], // U+2A01 N-ARY CIRCLED PLUS
  ["⨂", "⨂"], // U+2A02 N-ARY CIRCLED TIMES
  ["⋃", "⋃"], // U+22C3 N-ARY UNION
  ["⋂", "⋂"], // U+22C2 N-ARY INTERSECTION
  ["⋁", "⋁"], // U+22C1 N-ARY LOGICAL OR
  ["⋀", "⋀"], // U+22C0 N-ARY LOGICAL AND
])

/** Infix binary operators that terminate an n-ary's operand scan. The
 *  scan inside an mrow grabs siblings following the n-ary until one of
 *  these (or another n-ary, or end of mrow) is hit. Excludes parens/
 *  brackets — those bind tighter and belong inside the operand. */
export const NARY_OPERAND_TERMINATORS: ReadonlySet<string> = new Set([
  "+",
  "−", // U+2212 MINUS SIGN
  "-",
  "=",
  "≠",
  "<",
  ">",
  "≤",
  "≥",
  "≈",
  "≡",
  "∼",
  "∝",
  "∈",
  "∉",
  "⊆",
  "⊂",
  "⊇",
  "⊃",
  "⊕",
  "⊖",
  "⊗",
  "⊘",
  "→",
  "←",
  "↔",
  "⇒",
  "⇐",
  "⇔",
  "↦",
  "∧",
  "∨",
  "∪",
  "∩",
  "⊢",
  "⊨",
  "±",
  "∓",
])

/** Characters that, when wrapped by `<mover>` / `<munder>` with the
 *  `accent="true"` attribute (or implied by the character class), map
 *  to OMML `<m:acc>` with the given chr value. The chr attribute drops
 *  when the accent is the default (combining acute = U+0301). */
export const ACCENT_CHARS: ReadonlySet<string> = new Set([
  "ˇ", // ˇ caron
  "˘", // ˘ breve
  "˙", // ˙ dot above
  "˚", // ˚ ring above
  "˜", // ˜ tilde
  "^", // ^ circumflex
  "`", // ` grave (rare)
  "´", // ´ acute
  "¯", // ¯ macron
  "⃐", // ⃐ combining left arrow above
  "⃑", // ⃑ combining right arrow above
  "⃖", // ⃖ combining left arrow above (alt)
  "⃗", // ⃗ combining right arrow above (vector hat)
  "⃛", // ⃛ combining three dots above
  "⃜", // ⃜ combining four dots above
  "⃡", // ⃡ combining left-right arrow above
])

/** Characters that as the `<mover>` child mean an over-bar (m:bar).
 *  Bracket-class chars (⏞ ⏜ ⏝ ⏟) use m:groupChr instead — handled in
 *  GROUP_CHR_MAP below. */
export const BAR_OVER_CHARS: ReadonlySet<string> = new Set([
  "¯", // ¯ macron
  "̄", // ̄ combining macron
  "‾", // ‾ overline
])

/** Characters that as the `<munder>` child mean an under-bar
 *  (m:bar pos="bot"). Symmetric with BAR_OVER_CHARS. */
export const BAR_UNDER_CHARS: ReadonlySet<string> = new Set([
  "_", // _ low line
  "̲", // ̲ combining low line
  "‗", // ‗ double low line
])

/** Bracket-class over/under characters that map to OMML <m:groupChr>
 *  (a grouping bracket sized to the base). Value is the position the
 *  bracket is rendered on. */
export const GROUP_CHR_MAP: ReadonlyMap<string, "top" | "bot"> = new Map([
  ["⏞", "top"], // U+23DE TOP CURLY BRACKET (\overbrace)
  ["⏟", "bot"], // U+23DF BOTTOM CURLY BRACKET (\underbrace)
  ["⏜", "top"], // U+23DC TOP PARENTHESIS
  ["⏝", "bot"], // U+23DD BOTTOM PARENTHESIS
  ["⎴", "top"], // U+23B4 TOP SQUARE BRACKET (legacy)
  ["⎵", "bot"], // U+23B5 BOTTOM SQUARE BRACKET (legacy)
  ["︷", "top"], // U+FE37 PRESENTATION FORM (older fonts)
  ["︸", "bot"], // U+FE38 PRESENTATION FORM (older fonts)
])

/** mathvariant values → OMML `<m:sty m:val="…">`. Unknown variants drop
 *  back to default (Word's italic for length-1 mi, plain for the rest). */
export const MATHVARIANT_STYLE: ReadonlyMap<string, "p" | "b" | "i" | "bi"> = new Map([
  ["normal", "p"],
  ["bold", "b"],
  ["italic", "i"],
  ["bold-italic", "bi"],
])
