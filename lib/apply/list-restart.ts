import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildrenNS, wAttr, walkBodyParagraphs } from "@lib/xml/xml-utils.ts"
import { forkNumWithStartOverride, setParagraphNumPr } from "@lib/apply/numbering-mutation.ts"

/**
 * List-restart pass for single-level numbering schemes that opt in via a
 * scheme-level `restart` value.
 *
 * Word's `<w:num>` element IS the counter. All paragraphs sharing a numId
 * share one counter — so a style-level binding makes the counter run
 * continuously across the whole document. That's the right behavior for
 * caption / reference / global-counter shapes and is the default
 * (`restart: "continuous"`).
 *
 * The engine supports three opt-in restart modes:
 *
 *   "perInstance"        — restart at each contiguous run of paragraphs
 *                          bound to the scheme's styleId. A non-target
 *                          paragraph breaks the run and triggers a restart.
 *                          Classic use: 1./2./3. lists in Chapter 1 that
 *                          should start fresh in Chapter 2.
 *
 *   "byHeading"          — restart whenever a heading-styled paragraph
 *                          (any style with <w:outlineLvl>, or a paragraph
 *                          with direct <w:outlineLvl> in its own pPr) is
 *                          encountered. Each heading increments a per-target
 *                          epoch counter; when a list paragraph's last-seen
 *                          epoch differs from the current epoch, a new fork
 *                          is started. Use for "each chapter gets its own
 *                          1, 2, 3, …" without caring which heading level
 *                          triggered it.
 *
 *   { atStyleChange: S } — restart whenever a paragraph bound to styleId S
 *                          precedes the current list paragraph. Useful when
 *                          chapter boundaries are marked by a custom style
 *                          that doesn't carry outlineLvl (e.g. "ProposalH2").
 *
 * All three modes share the same OOXML mechanism: fork a fresh `<w:num>`
 * pointing to the same abstractNumId but carrying
 * `<w:lvlOverride><w:startOverride val="1"/></w:lvlOverride>`, and write
 * paragraph-level `<w:numPr>` on each affected paragraph so it overrides
 * the style-level binding.
 *
 * Only single-level schemes with a non-"continuous" restart are forked;
 * multi-level schemes and continuous single-level schemes are skipped.
 */

/**
 * Build the set of styleIds that declare <w:outlineLvl> in their <w:pPr>
 * in styles.xml. Used by applyListRestartPass for byHeading mode to detect
 * heading paragraphs via the style cascade, not just via direct pPr overrides.
 */
export function buildHeadingStyleIdSet(stylesDoc: Document): Set<string> {
  const w = NS.w
  const result = new Set<string>()
  const root = stylesDoc.documentElement
  if (!root) return result
  for (const styleEl of getChildrenNS(root, w, "style")) {
    const id = wAttr(styleEl, "styleId")
    if (!id) continue
    const pPr = firstChildNS(styleEl, w, "pPr")
    if (!pPr) continue
    const outlineLvlEl = firstChildNS(pPr, w, "outlineLvl")
    if (outlineLvlEl) result.add(id)
  }
  return result
}

export function applyListRestartPass(
  documentDoc: Document,
  numberingDoc: Document,
  installedSchemes: ReadonlyArray<{
    levels: ReadonlyArray<{
      level: number
      styleId: string
      restart: "continuous" | "perInstance" | "byHeading" | { atStyleChange: string }
    }>
    numId: string
    abstractNumId: string
  }>,
  headingStyleIds?: ReadonlySet<string>,
): void {
  type RestartMode = "perInstance" | "byHeading" | { atStyleChange: string }
  type Target = { styleId: string; abstractNumId: string; level: number; mode: RestartMode; baseNumId: string }
  const targets: Target[] = []
  for (const scheme of installedSchemes) {
    if (scheme.levels.length !== 1) continue
    const lvl = scheme.levels[0]!
    if (lvl.restart === "continuous" || !lvl.restart) continue
    targets.push({
      styleId: lvl.styleId,
      abstractNumId: scheme.abstractNumId,
      level: lvl.level,
      mode: lvl.restart as RestartMode,
      baseNumId: scheme.numId,
    })
  }
  if (targets.length === 0) return

  const styleIdToTarget = new Map(targets.map((t) => [t.styleId, t]))

  // For each target, accumulate the list of forked runs (each run = Element[]).
  // A "run" here means paragraphs that share the same counter fork.
  const runs = new Map<string, Element[][]>() // styleId → array of runs
  const currentRunByStyle = new Map<string, Element[]>() // styleId → current open run
  for (const t of targets) {
    runs.set(t.styleId, [])
    currentRunByStyle.set(t.styleId, [])
  }

  // Flush the current open run for a style into the runs list.
  const flush = (styleId: string) => {
    const cur = currentRunByStyle.get(styleId)!
    if (cur.length > 0) {
      runs.get(styleId)!.push(cur)
      currentRunByStyle.set(styleId, [])
    }
  }

  const w = NS.w

  // Per-target boundary tracker state.
  //   "byHeading":          headingEpoch tracks a monotonically-incrementing
  //                         counter that advances each time any heading paragraph
  //                         is encountered. lastSeenEpoch[targetStyleId] records
  //                         the epoch at the time of the most recent list paragraph.
  //                         When the current epoch differs from lastSeenEpoch, a
  //                         boundary has been crossed since the last list paragraph
  //                         and the run is flushed. This correctly handles multiple
  //                         chapters using the same heading styleId (e.g. Heading1).
  //   { atStyleChange: S }: atStylePending is set true when the marker style is seen;
  //                         the next list paragraph flushes and clears it.
  const headingEpoch = { value: 0 } // global epoch; increments on every heading paragraph
  const lastSeenEpochByStyle = new Map<string, number>() // targetStyleId → epoch at last list para
  const atStylePending = new Map<string, boolean>() // targetStyleId → marker seen?
  for (const t of targets) {
    if (t.mode === "byHeading") lastSeenEpochByStyle.set(t.styleId, 0)
    else if (typeof t.mode === "object") atStylePending.set(t.styleId, false)
  }

  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return

  for (const child of walkBodyParagraphs(body)) {
    const pPr = firstChildNS(child, w, "pPr")
    const pStyleEl = pPr ? firstChildNS(pPr, w, "pStyle") : null
    const paragraphStyleId = pStyleEl ? wAttr(pStyleEl, "val") : ""

    // Detect whether this paragraph is a heading boundary. Two detection paths:
    //   1. The paragraph's styleId is in the heading style set (style declares
    //      <w:outlineLvl> in styles.xml — the normal case for all Heading1-N).
    //   2. The paragraph has a direct <w:outlineLvl> in its own pPr (one-off
    //      override that doesn't go through a heading styleId).
    // The headingStyleIds set is pre-computed by the caller from styles.xml so
    // this pass doesn't need to receive the stylesDoc directly.
    let isHeading = false
    if (paragraphStyleId && headingStyleIds?.has(paragraphStyleId)) {
      isHeading = true
    } else if (pPr && firstChildNS(pPr, w, "outlineLvl")) {
      isHeading = true
    }

    const target = paragraphStyleId ? styleIdToTarget.get(paragraphStyleId) : undefined

    if (target) {
      const { mode } = target
      if (mode === "perInstance") {
        // Any prior run for OTHER perInstance targets is already flushed below
        // by their own non-target logic; here we just accumulate this paragraph.
        currentRunByStyle.get(target.styleId)!.push(child)
        // Flush all other perInstance targets whose run this paragraph interrupts.
        for (const t of targets) {
          if (t.styleId !== target.styleId && t.mode === "perInstance") flush(t.styleId)
        }
      } else if (mode === "byHeading") {
        // Check whether a heading boundary was crossed since the last list
        // paragraph for this target. The epoch counter advances on every
        // heading paragraph, so even if all chapters use the same styleId,
        // each chapter increments the epoch and triggers a flush.
        const lastEpoch = lastSeenEpochByStyle.get(target.styleId)!
        if (lastEpoch !== headingEpoch.value) {
          flush(target.styleId)
          lastSeenEpochByStyle.set(target.styleId, headingEpoch.value)
        }

        // Check for block-level restart override: if this paragraph already has
        // a direct <w:numPr> with a numId different from the scheme's base numId,
        // it was forked by a block-level restart:true pass — treat it as an
        // explicit override and don't rewrite it. Flush the current run so the
        // next paragraph starts fresh, but don't add this paragraph to any run.
        const existingNumPr = pPr ? firstChildNS(pPr, w, "numPr") : null
        if (existingNumPr) {
          const numIdEl = firstChildNS(existingNumPr, w, "numId")
          const existingNumId = numIdEl ? wAttr(numIdEl, "val") : null
          if (existingNumId !== null && existingNumId !== target.baseNumId) {
            flush(target.styleId)
            continue
          }
        }

        currentRunByStyle.get(target.styleId)!.push(child)
      } else {
        // atStyleChange: if the marker was seen since the last list paragraph,
        // flush now so this paragraph starts a fresh fork.
        if (atStylePending.get(target.styleId)) {
          flush(target.styleId)
          atStylePending.set(target.styleId, false)
        }
        currentRunByStyle.get(target.styleId)!.push(child)
      }
    } else {
      // Non-target paragraph — update boundary tracking state.
      if (isHeading) {
        // Advance the global epoch so every byHeading target detects the boundary.
        headingEpoch.value++
      }
      for (const t of targets) {
        if (t.mode === "perInstance") {
          // Any non-target paragraph breaks the run.
          flush(t.styleId)
        } else if (typeof t.mode === "object") {
          if (paragraphStyleId === t.mode.atStyleChange) {
            // Marker style seen — next list paragraph starts a new run.
            atStylePending.set(t.styleId, true)
          }
        }
      }
    }
  }
  for (const t of targets) flush(t.styleId)

  // Fork every collected run (including the first, so the style-level numId
  // becomes a declaration only; all list paragraphs reference per-run numIds).
  for (const t of targets) {
    for (const run of runs.get(t.styleId)!) {
      const newNumId = forkNumWithStartOverride(numberingDoc, t.abstractNumId, t.level)
      for (const p of run) setParagraphNumPr(p, newNumId, t.level)
    }
  }
}

/**
 * Fork a fresh `<w:num>` for a paragraph-level explicit restart
 * (`numbering.restart: true` in a ParagraphBlock).
 *
 * Looks up the abstractNumId from the paragraph's current numId in
 * numberingDoc, forks it with `<w:startOverride val="1"/>`, and rewrites the
 * paragraph's `<w:numPr>` to use the new numId. Restart applies only to this
 * paragraph; subsequent paragraphs continue with their own declared numbering.
 *
 * Throws when numId doesn't resolve in numberingDoc — schema-valid input but
 * doctree-inconsistent (caller passed a numId not installed in the doc).
 */
export function applyParagraphLevelRestart(
  pEl: Element,
  numberingDoc: Document,
  numId: string,
  level: number,
): void {
  const w = NS.w
  const root = numberingDoc.documentElement!
  let abstractNumId: string | null = null
  for (const num of getChildrenNS(root, w, "num")) {
    if (wAttr(num, "numId") === numId) {
      const absRef = firstChildNS(num, w, "abstractNumId")
      abstractNumId = absRef ? wAttr(absRef, "val") : null
      break
    }
  }
  if (!abstractNumId) {
    throw new Error(
      `numbering.restart: numId "${numId}" not found in numbering.xml. ` +
        `Fix: use a numId that exists in the document's numbering.xml, or install a numbering scheme via config.numbering first.`,
    )
  }
  const newNumId = forkNumWithStartOverride(numberingDoc, abstractNumId, level)
  setParagraphNumPr(pEl, newNumId, level)
}
