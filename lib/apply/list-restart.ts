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
 *   "byHeading"          — restart whenever the nearest preceding heading-
 *                          styled paragraph (any style with <w:outlineLvl>)
 *                          changes. The engine tracks the heading styleId
 *                          seen most recently before each list paragraph;
 *                          when it changes, a new fork is started. Use for
 *                          "each chapter gets its own 1, 2, 3, …" without
 *                          caring which heading level triggered it.
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
): void {
  type RestartMode = "perInstance" | "byHeading" | { atStyleChange: string }
  type Target = { styleId: string; abstractNumId: string; level: number; mode: RestartMode }
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
  //   "byHeading":          headingOwner tracks the styleId of the most-recently-
  //                         seen heading paragraph. When it changes, the current run
  //                         is flushed so the next item starts a fresh fork.
  //   { atStyleChange: S }: atStylePending is set true when the marker style is seen;
  //                         the next list paragraph flushes and clears it.
  const headingOwner = new Map<string, string | null>() // targetStyleId → heading styleId
  const atStylePending = new Map<string, boolean>() // targetStyleId → marker seen?
  for (const t of targets) {
    if (t.mode === "byHeading") headingOwner.set(t.styleId, null)
    else if (typeof t.mode === "object") atStylePending.set(t.styleId, false)
  }

  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return

  for (const child of walkBodyParagraphs(body)) {
    const pPr = firstChildNS(child, w, "pPr")
    const pStyleEl = pPr ? firstChildNS(pPr, w, "pStyle") : null
    const paragraphStyleId = pStyleEl ? wAttr(pStyleEl, "val") : ""

    // Detect whether this paragraph is a heading boundary (has outlineLvl in
    // its direct pPr). Style-cascade-inherited outlineLvl is not checked here
    // because list-restart.ts doesn't receive styles.xml; heading styles always
    // set outlineLvl directly (that's how Word populates TOC navigation), so
    // direct-pPr detection is sufficient in practice.
    let isHeading = false
    let isHeadingStyleId: string | null = null
    if (pPr) {
      const outlineLvlEl = firstChildNS(pPr, w, "outlineLvl")
      if (outlineLvlEl) {
        isHeading = true
        isHeadingStyleId = paragraphStyleId || null
      }
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
        // Run continues; heading-boundary flushes happen in the non-target branch.
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
      for (const t of targets) {
        if (t.mode === "perInstance") {
          // Any non-target paragraph breaks the run.
          flush(t.styleId)
        } else if (t.mode === "byHeading") {
          // Only heading paragraphs trigger a boundary for this mode.
          if (isHeading) {
            const prevOwner = headingOwner.get(t.styleId)
            if (prevOwner !== isHeadingStyleId) {
              flush(t.styleId)
              headingOwner.set(t.styleId, isHeadingStyleId)
            }
          }
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
