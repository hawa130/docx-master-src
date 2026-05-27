/**
 * Caption counter simulator.
 *
 * Walks the document body in order, advancing per-identifier counters
 * for caption-class paragraphs, resolving STYLEREF chapter prefixes
 * against outline numbering outputs, producing:
 *
 *   - Per-field result values to write into SEQ / STYLEREF / sub-counter
 *     result runs (so caption paragraphs read correctly without Word F9)
 *   - Per-caption-paragraph full rendered text (prefix + chapter +
 *     parent + sub + suffix), consumed by REF placeholder backfill so
 *     body cross-references like "see (2.3)" render pre-F9
 *
 * Pure function over input data structures — no styles.xml or outline
 * simulator access. Caller (apply pipeline) preprocesses:
 *
 *   - Walks outline simulator's output to build `outlineParagraphs` map
 *     (each tagged with its styleName + rendered number)
 *   - Resolves agent's captions config into ResolvedCaptionConfig per
 *     identifier (styleId → styleName + outline level via styles.xml)
 *   - Caption emitters create `fills` + `resets` records pointing at
 *     the actual result `<w:t>` elements
 *
 * State machine per identifier:
 *
 *   subGroup omitted (standalone): parent++, sub = 0, close open subgroup
 *   subGroup "start":              parent++, sub = 1, open subgroup
 *   subGroup "continue":           requires open subgroup; sub++, parent unchanged
 *   CaptionCounterReset:           parent = newValue, sub = 0, close subgroup
 *
 * Subgroup continue without preceding start throws — pre-scan should
 * catch this, but the simulator double-checks (the resulting render
 * would silently look wrong otherwise).
 */

import { NS } from "@lib/parse/types.ts"
import { firstChildNS, walkBodyParagraphs } from "@lib/xml/xml-utils.ts"
import { toAlphaCounter, toRoman } from "@lib/parse/counter-format.ts"
import type { SeqFormat } from "@lib/edit/fields/seq-field.ts"

const w = NS.w

/** Resolved form of a `captions[id]` config entry. styleId references
 * resolved to styleName + outlineLevel via styles.xml lookup upstream. */
export interface ResolvedChapterPrefixEntry {
  styleName: string
  /** Heading's styleId — used by the sim to count occurrences when a
   * `format` override is in effect (count integer instead of relying on
   * the heading's rendered text). */
  styleId: string
  outlineLevel: number
  /** When set, override the heading's native number rendering with this
   * format. Emitted to Word's runtime as a `\* <FORMAT>` switch on
   * STYLEREF; the counter sim re-formats the per-style integer
   * occurrence count. */
  format?: SeqFormat
}

export interface ResolvedCaptionConfig {
  identifier: string
  prefix: string
  suffix: string
  format: SeqFormat
  chapterPrefix: ResolvedChapterPrefixEntry[]
  chapterSeparator: string
  bodySeparator: string
  paragraphStyleId: string
  /** Derived from the last chapterPrefix entry's outlineLevel; undefined
   * when chapterPrefix is empty. Used as the SEQ field's `\s` value. */
  restartAtOutlineLevel: number | undefined
  /** Set when this identifier supports subequations. */
  subCounter: { format: SeqFormat; prefix: string; suffix: string } | undefined
}

/** Caption emitter's record per caption paragraph, telling the simulator
 * where to write the rendered values. Created at emit time (for new
 * captions) or during standardize re-emit (for existing captions
 * detected in source). */
export interface PendingCaptionFill {
  /** The caption paragraph element. Body-walk order determines counter
   * sequence. */
  paragraph: Element
  /** SEQ identifier (e.g. "Equation"). Must match a key in
   * `CaptionSimulatorInput.configs`. */
  identifier: string
  /** Subequation membership. Undefined = standalone. */
  subGroup: "start" | "continue" | undefined
  /** STYLEREF result text elements, ordered by ResolvedCaptionConfig's
   * `chapterPrefix` entries. Empty when chapterPrefix is `[]`. */
  chapterPrefixResults: Element[]
  /** SEQ field's result text element for the parent counter. */
  parentSeqResult: Element
  /** Sub-counter SEQ's result text element. Set only when `subGroup`
   * is `"start"` or `"continue"`. */
  subSeqResult?: Element
  /** Caption body text (everything after the bodySeparator). Present for
   * CaptionBlock and standardize re-emitted paragraphs; absent for
   * EquationBlock captions which carry no body text. Used by the dry-run
   * preview to show the full visible caption string (counter + body). */
  bodyText?: string
}

/** CaptionCounterReset emit record. Distinct from caption fills since
 * it advances counter state without rendering a visible value. */
export interface PendingCaptionReset {
  paragraph: Element
  identifier: string
  newValue: number
}

export interface CaptionSimulatorInput {
  fills: PendingCaptionFill[]
  resets: PendingCaptionReset[]
  configs: Map<string, ResolvedCaptionConfig>
  /** Outline-numbered paragraphs tagged with their styleName + rendered
   * number. Produced by the existing outline numbering simulator;
   * caller annotates with styleName from styles.xml. */
  outlineParagraphs: Map<Element, { styleName: string; rendered: string }>
}

export interface CaptionSimulatorOutput {
  /** Per-result-element values to write into SEQ / STYLEREF / sub-counter
   * result text elements. */
  fieldValues: Map<Element, string>
  /** Per-caption-paragraph full rendered text — prefix + chapterPrefix
   * values (joined by chapterSeparator) + parent + sub (with subCounter
   * wrap) + suffix. Used by REF placeholder backfill. */
  fullCaptionText: Map<Element, string>
}

interface CounterState {
  parent: number
  sub: number
  openSubGroup: boolean
}

export function simulateCaptions(
  documentDoc: Document,
  input: CaptionSimulatorInput,
): CaptionSimulatorOutput {
  const fieldValues = new Map<Element, string>()
  const fullCaptionText = new Map<Element, string>()
  const states = new Map<string, CounterState>()
  /** styleName → heading's rendered counter text (e.g. "一"). Used when
   * the chapterPrefix entry has no `format` override — STYLEREF returns
   * the heading's native rendering. */
  const latestHeading = new Map<string, string>()
  /** styleName → integer occurrence count. Used when a chapterPrefix
   * entry's `format` override is in effect — counter sim re-formats the
   * integer per the override (matching Word's `\* <FORMAT>` runtime). */
  const latestHeadingCount = new Map<string, number>()

  const fillByPara = new Map<Element, PendingCaptionFill>()
  for (const f of input.fills) fillByPara.set(f.paragraph, f)
  const resetByPara = new Map<Element, PendingCaptionReset>()
  for (const r of input.resets) resetByPara.set(r.paragraph, r)

  // styleName → list of identifiers that reset when a paragraph with
  // that style is encountered. Derived from each config's last
  // chapterPrefix entry (the deepest level controls restart, per SEQ \s
  // semantics).
  const resetTriggers = new Map<string, string[]>()
  for (const [identifier, config] of input.configs) {
    if (config.chapterPrefix.length === 0) continue
    const lastStyleName = config.chapterPrefix[config.chapterPrefix.length - 1]!.styleName
    const list = resetTriggers.get(lastStyleName) ?? []
    list.push(identifier)
    resetTriggers.set(lastStyleName, list)
  }

  const root = documentDoc.documentElement
  if (!root) return { fieldValues, fullCaptionText }
  const body = firstChildNS(root, w, "body")
  if (!body) return { fieldValues, fullCaptionText }

  for (const para of walkBodyParagraphs(body)) {
    const outlineInfo = input.outlineParagraphs.get(para)
    if (outlineInfo) {
      latestHeading.set(outlineInfo.styleName, outlineInfo.rendered)
      latestHeadingCount.set(
        outlineInfo.styleName,
        (latestHeadingCount.get(outlineInfo.styleName) ?? 0) + 1,
      )
      // Reset any caption counters whose chapter restart hangs off this
      // heading style. Mirrors Word's SEQ \s N behavior.
      const triggered = resetTriggers.get(outlineInfo.styleName)
      if (triggered) {
        for (const identifier of triggered) {
          const state = stateFor(states, identifier)
          state.parent = 0
          state.sub = 0
          state.openSubGroup = false
        }
      }
      continue
    }

    const reset = resetByPara.get(para)
    if (reset) {
      // newValue is "the value the next caption renders" (agent
      // intent). Pre-decrement state so the next caption's parent++
      // lands on newValue. Mirrors the `\r (newValue - 1)` convention
      // emitCaptionReset uses for Word's runtime SEQ resolution.
      const state = stateFor(states, reset.identifier)
      state.parent = reset.newValue - 1
      state.sub = 0
      state.openSubGroup = false
      continue
    }

    const fill = fillByPara.get(para)
    if (!fill) continue

    const config = input.configs.get(fill.identifier)
    if (!config) {
      // Unknown identifier — pass through unchanged. Standardize re-emit
      // surfaces this case via a separate warn; simulator just skips.
      continue
    }

    const state = stateFor(states, fill.identifier)
    applyFillToState(state, fill)

    // Resolve STYLEREF chapter prefix values.
    // With format override: count integer occurrences of the styled
    // paragraph and format per override (mirrors Word's STYLEREF
    // \* <FORMAT> runtime). Without override: use the heading's native
    // rendering captured at the latestHeading map.
    for (let i = 0; i < fill.chapterPrefixResults.length; i++) {
      const entry = config.chapterPrefix[i]
      const resultEl = fill.chapterPrefixResults[i]!
      if (!entry) continue
      let rendered: string
      if (entry.format !== undefined) {
        const count = latestHeadingCount.get(entry.styleName) ?? 0
        rendered = count > 0 ? formatCounter(count, entry.format) : "0"
      } else {
        rendered = latestHeading.get(entry.styleName) ?? "0"
      }
      fieldValues.set(resultEl, rendered)
    }

    // Parent counter
    const parentRendered = formatCounter(state.parent, config.format)
    fieldValues.set(fill.parentSeqResult, parentRendered)

    // Sub-counter
    let subRendered = ""
    if (fill.subSeqResult && config.subCounter) {
      subRendered = formatCounter(state.sub, config.subCounter.format)
      fieldValues.set(fill.subSeqResult, subRendered)
    }

    // Full text for REF backfill: prefix + chapter parts + parent + sub + suffix
    const chapterParts: string[] = []
    for (let i = 0; i < fill.chapterPrefixResults.length; i++) {
      const resultEl = fill.chapterPrefixResults[i]!
      chapterParts.push(fieldValues.get(resultEl) ?? "0")
    }
    const chapterStr =
      chapterParts.length > 0
        ? chapterParts.join(config.chapterSeparator) + config.chapterSeparator
        : ""
    let subStr = ""
    if (fill.subSeqResult && config.subCounter && state.sub > 0) {
      subStr = config.subCounter.prefix + subRendered + config.subCounter.suffix
    }
    const full = config.prefix + chapterStr + parentRendered + subStr + config.suffix
    fullCaptionText.set(para, full)
  }

  return { fieldValues, fullCaptionText }
}

function stateFor(states: Map<string, CounterState>, identifier: string): CounterState {
  let s = states.get(identifier)
  if (!s) {
    s = { parent: 0, sub: 0, openSubGroup: false }
    states.set(identifier, s)
  }
  return s
}

function applyFillToState(state: CounterState, fill: PendingCaptionFill): void {
  if (fill.subGroup === "continue") {
    if (!state.openSubGroup) {
      throw new Error(
        `Caption "${fill.identifier}": subGroup "continue" without preceding "start". ` +
          `Pre-scan should have caught this — investigate the pipeline if it reaches the simulator.`,
      )
    }
    state.sub++
  } else if (fill.subGroup === "start") {
    state.parent++
    state.sub = 1
    state.openSubGroup = true
  } else {
    state.parent++
    state.sub = 0
    state.openSubGroup = false
  }
}

/* ------------------------ number formatting ------------------------ */

export function formatCounter(n: number, format: SeqFormat): string {
  switch (format) {
    case "arabic":
      return String(n)
    case "alphabetic":
      return toAlphaCounter(n, false)
    case "ALPHABETIC":
      return toAlphaCounter(n, true)
    case "roman":
      return toRoman(n).toLowerCase()
    case "ROMAN":
      return toRoman(n)
    case "chinese":
      return toChineseNum(n, false)
    case "chinese-formal":
      return toChineseNum(n, true)
  }
}

function toChineseNum(n: number, formal: boolean): string {
  const digits = formal
    ? ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
    : ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
  if (n <= 0) return ""
  if (n < 10) return digits[n]!
  if (n < 20) return n === 10 ? "十" : "十" + digits[n - 10]!
  if (n < 100) {
    const tens = Math.floor(n / 10)
    const ones = n % 10
    return digits[tens]! + "十" + (ones === 0 ? "" : digits[ones]!)
  }
  // Captions rarely exceed 99 occurrences in practice; for 100+ fall
  // back to Arabic. If real use cases appear, extend the table here.
  return String(n)
}
