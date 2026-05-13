import { NS } from "@lib/parse/types.ts"
import { firstChildNS, wAttr, walkBodyParagraphs } from "@lib/xml/xml-utils.ts"
import { forkNumWithStartOverride, setParagraphNumPr } from "@lib/apply/numbering-mutation.ts"

/**
 * Per-instance restart pass for single-level numbering schemes that opt in
 * via `restart: "perInstance"`.
 *
 * Word's `<w:num>` element IS the counter. All paragraphs sharing a numId
 * share one counter — so a style-level binding makes the counter run
 * continuously across the whole document. That's the right behavior for
 * caption / reference / global-counter shapes and is the default
 * (`restart: "continuous"`).
 *
 * For procedural list shapes (ListNumber / ListBullet 1./2./3. lists), a
 * shared counter is wrong: a list in Chapter 1 would continue numbering in
 * Chapter 2. Multi-level heading schemes solve this with `<w:lvlRestart>`,
 * but single-level schemes have no analogous mechanism. The workaround Word
 * itself uses: each list instance gets its own `<w:num>` pointing to the
 * same abstractNumId but carrying
 * `<w:lvlOverride><w:startOverride val="1"/></w:lvlOverride>`. Each
 * instance's paragraphs get paragraph-level `<w:numPr>` overriding the
 * style-level binding.
 *
 * "Instance" = contiguous run of paragraphs with the target styleId in
 * document tree order. Non-target paragraphs break the run. Matches user
 * intuition: a list interrupted by a heading or body paragraph starts fresh.
 *
 * Only single-level schemes with `restart === "perInstance"` are forked;
 * multi-level schemes and continuous single-level schemes are skipped.
 */
export function applyListRestartPass(
  documentDoc: Document,
  numberingDoc: Document,
  installedSchemes: ReadonlyArray<{
    levels: ReadonlyArray<{ level: number; styleId: string; restart: "continuous" | "perInstance" }>
    numId: string
    abstractNumId: string
  }>,
): void {
  type Target = { styleId: string; abstractNumId: string; level: number }
  const targets: Target[] = []
  for (const scheme of installedSchemes) {
    if (scheme.levels.length !== 1) continue
    const lvl = scheme.levels[0]!
    if (lvl.restart !== "perInstance") continue
    targets.push({
      styleId: lvl.styleId,
      abstractNumId: scheme.abstractNumId,
      level: lvl.level,
    })
  }
  if (targets.length === 0) return

  const styleIdToTarget = new Map(targets.map((t) => [t.styleId, t]))
  const runs = new Map<string, Element[][]>()
  const currentRunByStyle = new Map<string, Element[]>()
  for (const t of targets) {
    runs.set(t.styleId, [])
    currentRunByStyle.set(t.styleId, [])
  }

  const flush = (styleId: string) => {
    const cur = currentRunByStyle.get(styleId)!
    if (cur.length > 0) {
      runs.get(styleId)!.push(cur)
      currentRunByStyle.set(styleId, [])
    }
  }

  const w = NS.w
  // Scope to body + tbl/tr/tc paragraphs (the canonical walker). The cross-
  // ref pipeline downstream (`injectChapterCounters`, caption sim, numbering
  // sim) all share this scope; forking a numId for a list paragraph whose
  // counter no downstream pass would write would leave the forked numId
  // unreferenced. Footnotes / endnotes / comments aren't in scope.
  const body = firstChildNS(documentDoc.documentElement, w, "body")
  if (!body) return
  for (const child of walkBodyParagraphs(body)) {
    const pPr = firstChildNS(child, w, "pPr")
    const pStyle = pPr ? firstChildNS(pPr, w, "pStyle") : null
    const styleId = pStyle ? wAttr(pStyle, "val") : ""
    const target = styleId ? styleIdToTarget.get(styleId) : undefined
    if (target) {
      currentRunByStyle.get(target.styleId)!.push(child)
      // Any non-target run gets flushed by this paragraph not matching
      // its styleId — handled below.
      for (const t of targets) {
        if (t.styleId !== target.styleId) flush(t.styleId)
      }
    } else {
      for (const t of targets) flush(t.styleId)
    }
  }
  for (const t of targets) flush(t.styleId)

  // Fork every run (including the first). The style-level numId becomes a
  // binding declaration only; every list paragraph references its run's
  // forked numId via paragraph-level override.
  for (const t of targets) {
    for (const run of runs.get(t.styleId)!) {
      const newNumId = forkNumWithStartOverride(numberingDoc, t.abstractNumId, t.level)
      for (const p of run) setParagraphNumPr(p, newNumId, t.level)
    }
  }
}
