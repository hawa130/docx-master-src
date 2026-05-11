import { NS } from "@lib/parse/types.ts"
import { firstChildNS, getChildren, wAttr } from "@lib/xml/xml-utils.ts"
import { forkNumWithStartOverride, setParagraphNumPr } from "@lib/apply/numbering-mutation.ts"

/**
 * Per-instance restart pass for single-level numbering schemes.
 *
 * Word's `<w:num>` element IS the counter. All paragraphs sharing a numId
 * share one counter — so a style-level binding makes the counter run
 * continuously across the whole document. Multi-level heading schemes use
 * `<w:lvlRestart>` to reset sub-levels at parent boundaries, but single-level
 * "list-shaped" schemes have no analogous mechanism: a 1./2./3. list in
 * Chapter 1 would continue as 4./5./6. in Chapter 2.
 *
 * The fix Word itself uses internally: each separate list instance gets its
 * own `<w:num>` that points to the same abstractNumId but carries
 * `<w:lvlOverride><w:startOverride val="1"/></w:lvlOverride>`. Each instance
 * gets paragraph-level `<w:numPr>` overriding the style-level binding.
 *
 * "Instance" = contiguous run of paragraphs with the target styleId in
 * document tree order. Non-target paragraphs break the run. Matches user
 * intuition: a list interrupted by a heading or body paragraph starts fresh.
 *
 * Auto-applies to single-level schemes only — multi-level heading schemes
 * use lvlRestart and don't need forking.
 */
export function applyListRestartPass(
  documentDoc: Document,
  numberingDoc: Document,
  installedSchemes: ReadonlyArray<{
    levels: ReadonlyArray<{ level: number; styleId: string }>
    numId: string
    abstractNumId: string
  }>,
): void {
  type Target = { styleId: string; abstractNumId: string; level: number }
  const targets: Target[] = []
  for (const scheme of installedSchemes) {
    if (scheme.levels.length !== 1) continue
    const lvl = scheme.levels[0]!
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
  const visit = (el: Element) => {
    for (const child of getChildren(el)) {
      if (child.namespaceURI === w && child.localName === "p") {
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
      } else {
        visit(child)
      }
    }
  }
  visit(documentDoc.documentElement!)
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
