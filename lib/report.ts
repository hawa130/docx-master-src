import type { EditsPreviewEntry } from "@lib/edit-engine.ts"
import type { ImportResult } from "@lib/template-import.ts"
import type {
  FlagRecord,
  RestyleSample,
  StyleConfigEntry,
  StyleResolutionEntry,
} from "./config-types.ts"

/* ------------- display helpers ------------- */

/**
 * Pick the user-facing typographic fields from a resolved style for
 * side-by-side display. Excludes mechanical fields (id, name, basedOn,
 * fromParagraph, overrides) the user wouldn't recognize.
 */
export function extractDisplayFields(def: StyleConfigEntry): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  // fontCJK first when set: in Chinese docs the CJK font is what users
  // perceive as "the font", so leading with it avoids the "why is the body
  // Arial?" double-take when the agent reviews the resolution block. When
  // only fontLatin is set, fontCJK is absent and the order collapses to
  // the natural one anyway.
  const interesting: (keyof StyleConfigEntry)[] = [
    "fontCJK",
    "fontLatin",
    "size",
    "bold",
    "italic",
    "color",
    "alignment",
    "lineSpacing",
    "lineRule",
    "spaceBefore",
    "spaceAfter",
    "firstLineIndent",
    "hangingIndent",
    "outlineLevel",
  ]
  for (const k of interesting) {
    if (def[k] !== undefined) out[k] = def[k]
  }
  return out
}

function formatResolvedFields(fields: Record<string, unknown>): string {
  const parts: string[] = []
  for (const [k, v] of Object.entries(fields)) {
    parts.push(`${k}=${JSON.stringify(v)}`)
  }
  return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`
}

/* ------------- change report ------------- */

export function printReport(args: {
  source: string
  injected: string[]
  updated: string[]
  restyleStats: Map<string, number>
  flags: FlagRecord[]
  manualNumberingRemoved: Map<string, number>
  manualNumberingByStyle: Map<string, Map<string, number>>
  patternMatchStats: Map<string, number>
  patternStripStats: Map<string, number>
  styleResolutions: StyleResolutionEntry[]
  derivedFrom: Map<string, number>
  output: string
  dryRun: boolean
  samples: Map<string, RestyleSample[]>
  implicitKeepByFingerprint: Map<
    string,
    { empty: number; nonEmpty: number; nonEmptySamples: string[] }
  >
  unstrippedByStyle: Map<string, { count: number; samples: string[] }>
  /** dry-run only: heuristic detection of typed-prefix residue in paragraphs
   * bound to auto-numbered styles (catches inserted-with-prefix paragraphs
   * that don't go through the rule-routing path unstrippedByStyle covers). */
  manualNumberingDetected?: Map<string, { count: number; samples: string[] }>
  numberingBindings: Array<{
    styleId: string
    level: number
    lvlText: string
    suff: "tab" | "space" | "nothing"
  }>
  templateImport: ImportResult | null
  /** dry-run only: per-op preview of edits[] (locator-resolved, not mutated). */
  editsPreview: EditsPreviewEntry[]
}) {
  const lines: string[] = []
  lines.push(
    args.dryRun ? "=== Change Report (DRY RUN — no file written) ===" : "=== Change Report ===",
  )
  // Echo absolute paths up top so an agent that may have changed cwd between
  // calls (or is reading a captured report later) can see exactly which files
  // were involved without re-deriving from `config.source` / `config.output`.
  lines.push(`Source: ${args.source}`)
  lines.push(`Output: ${args.output}`)
  lines.push("")
  if (args.templateImport) {
    const ti = args.templateImport
    const directly = ti.imported.filter((id) => !ti.pulledAncestors.includes(id))
    lines.push(
      `Imported from template: ${directly.length} requested + ${ti.pulledAncestors.length} basedOn ancestors`,
    )
    lines.push(`  styles: [${ti.imported.join(", ")}]`)
    if (ti.pulledAncestors.length > 0)
      lines.push(`  pulled ancestors: [${ti.pulledAncestors.join(", ")}]`)
    if (ti.numIdRemap.size > 0) {
      const remaps = [...ti.numIdRemap.entries()].map(([o, n]) => `${o}→${n}`).join(", ")
      lines.push(`  numIds migrated: ${remaps}`)
    }
    lines.push("")
  }
  const annotate = (id: string) => {
    const src = args.derivedFrom.get(id)
    return src !== undefined ? `${id} (from #${src})` : id
  }
  lines.push(`Styles injected: ${args.injected.length} (${args.injected.map(annotate).join(", ")})`)
  lines.push(`Styles updated:  ${args.updated.length} (${args.updated.map(annotate).join(", ")})`)
  lines.push("")
  let totalRestyled = 0
  for (const c of args.restyleStats.values()) totalRestyled += c
  lines.push(`Paragraphs restyled: ${totalRestyled}`)
  for (const [styleId, count] of args.restyleStats) {
    lines.push(`  ${styleId}: ${count} paragraphs`)
  }
  if (args.implicitKeepByFingerprint.size > 0) {
    let totalEmpty = 0
    let totalNonEmpty = 0
    for (const v of args.implicitKeepByFingerprint.values()) {
      totalEmpty += v.empty
      totalNonEmpty += v.nonEmpty
    }
    lines.push(`Paragraphs untouched: ${totalEmpty + totalNonEmpty}`)
    if (totalEmpty > 0) {
      lines.push(`  empty (likely spacers): ${totalEmpty}`)
    }
    if (totalNonEmpty > 0) {
      // Non-empty untouched are the coverage signal — break down by
      // fingerprint with up-to-2 sample texts so the agent can spot a missed
      // role at a glance. On the Targeted Edit path this is expected (only
      // intentional changes apply); on Full Standardization, an unfamiliar
      // entry here means a fingerprint slipped through and the samples make
      // it cheap to confirm.
      lines.push(`  non-empty (verify coverage): ${totalNonEmpty}`)
      const sortedEntries = [...args.implicitKeepByFingerprint.entries()]
        .filter(([, v]) => v.nonEmpty > 0)
        .sort((a, b) => b[1].nonEmpty - a[1].nonEmpty)
      for (const [fp, v] of sortedEntries) {
        const samples =
          v.nonEmptySamples.length > 0
            ? `  e.g. ${v.nonEmptySamples.map((s) => `"${s}"`).join(" / ")}`
            : ""
        lines.push(`    ${fp}×${v.nonEmpty}${samples}`)
      }
    }
  }
  lines.push("")
  // Auto-numbering bindings: which lvlText each numbered style will gain at
  // render. We don't compute the actual rendered numbers (would require
  // walking H1/H2/H3 counters across the whole doc); the agent gets the
  // structural binding instead, which is enough to confirm "Heading2 → '%1.%2'"
  // is what they intended without doing the arithmetic.
  if (args.numberingBindings.length > 0) {
    lines.push("Auto-numbering bindings (will prepend at render):")
    for (const b of args.numberingBindings) {
      lines.push(`  ${b.styleId} → "${b.lvlText}" + suff:${b.suff} (level ${b.level})`)
    }
    lines.push("")
  }
  if (args.manualNumberingRemoved.size > 0) {
    lines.push("Manual numbering converted:")
    for (const [pat, count] of args.manualNumberingRemoved) {
      lines.push(`  Prefix removed: "${pat}" (${count})`)
    }
    lines.push("")
  }
  // Mixed-scheme detection: a heading style that had >1 distinct strip pattern
  // hit means the source document used inconsistent manual numbering within
  // one logical level (e.g. chapter 1's H2 "1.1 ..." and chapter 2's H2
  // "1. ..."). This was already handled correctly by stripPrefixPatterns
  // matching in priority order; surfacing it here lets the agent tell the
  // user explicitly that normalization changed an inconsistent input — which
  // SKILL.md flags as a normalization decision worth confirming.
  const mixedStyles = [...args.manualNumberingByStyle.entries()].filter(([, m]) => m.size >= 2)
  if (mixedStyles.length > 0) {
    lines.push("Mixed manual numbering detected (source inconsistent):")
    for (const [styleId, patMap] of mixedStyles) {
      const breakdown = [...patMap.entries()]
        .sort((a, b) => b[1] - a[1])
        .map(([pat, n]) => `"${pat}"×${n}`)
        .join(", ")
      lines.push(`  ${styleId}: ${breakdown}`)
    }
    lines.push(
      "  Normalization unified these to one scheme — worth confirming with the user before final write.",
    )
    lines.push("")
  }
  // Numbered-style paragraphs that fell through every stripPrefixPattern.
  // We surface the leading-text samples — NOT a classified shape — and let
  // the agent read them: if the samples show "1. 数据集管理" / "2. ..." the
  // agent recognises a missed "%1." pattern and adds it; if the samples
  // show "研究方法" / "实验结果" the agent recognises clean headings without
  // manual numbering and does nothing. Avoids hardcoding a shape list that
  // would always lag the variety of typed prefixes real docs use.
  if (args.unstrippedByStyle.size > 0) {
    lines.push("Numbered-style paragraphs not matched by any stripPrefixPattern:")
    for (const [styleId, info] of args.unstrippedByStyle) {
      const samples = info.samples.map((s) => `"${s}"`).join(" / ")
      lines.push(`  ${styleId}: ${info.count} paragraphs  e.g. ${samples}`)
    }
    lines.push("  → Read the samples: a typed prefix you missed (add to stripPrefixPatterns)")
    lines.push("    or a clean heading without manual numbering (no action needed).")
    lines.push("")
  }
  // Dry-run only: heuristic catch for typed prefixes that survived into
  // numbered-style paragraphs. Distinct from unstrippedByStyle above:
  // unstripped surfaces leading-text *samples* for agent classification;
  // this section reports likely typed prefixes the regex inventory
  // recognised, so the action is more concrete (strip on insert, or add a
  // stripPrefixPattern for the chrome shape).
  if (args.manualNumberingDetected && args.manualNumberingDetected.size > 0) {
    lines.push("Manual numbering detected in numbered-style paragraphs:")
    for (const [styleId, info] of args.manualNumberingDetected) {
      const samples = info.samples.map((s) => `"${s}"`).join(" / ")
      lines.push(`  ${styleId}: ${info.count} paragraphs  e.g. ${samples}`)
    }
    lines.push("  → For inserts (in `edits[]`): drop the typed prefix from `text` —")
    lines.push("    the styleId's auto-numbering scheme already emits the marker.")
    lines.push("    For chrome paragraphs: add a matching `stripPrefixPatterns` entry")
    lines.push("    on the bound numbering level so the prefix is stripped on retag.")
    lines.push("")
  }
  if (args.patternMatchStats.size > 0) {
    lines.push("Pattern rules matched:")
    for (const [src, count] of args.patternMatchStats) {
      const stripped = args.patternStripStats.get(src) ?? 0
      const stripNote = stripped > 0 ? ` (stripped match in ${stripped})` : ""
      lines.push(`  /${src}/: ${count} paragraphs${stripNote}`)
    }
    lines.push("")
  }
  if (args.flags.length > 0) {
    lines.push(`Flagged (not modified): ${args.flags.length}`)
    for (const f of args.flags) {
      lines.push(`  #${f.paraIndex}: ${f.reason}`)
    }
    lines.push("")
  }
  if (args.styleResolutions.length > 0) {
    lines.push("=== Style Resolution (verify by reading) ===")
    lines.push("  The script does not parse natural language. For styles with a")
    lines.push("  user spec, compare it to the agent-resolved fields by eye —")
    lines.push("  any mismatch means the agent's translation needs adjustment.")
    lines.push("  Styles without a spec are still listed so the resolved fields")
    lines.push("  are auditable.")
    lines.push("")
    for (const r of args.styleResolutions) {
      lines.push(`  ${r.styleId}`)
      if (r.userSpec !== null) {
        lines.push(`    User specified: "${r.userSpec}"`)
      } else {
        lines.push(`    User specified: (none — no requirements entry)`)
      }
      lines.push(`    Agent resolved: ${formatResolvedFields(r.resolved)}`)
    }
    lines.push("")
  }
  if (args.editsPreview.length > 0) {
    lines.push("=== Edits Preview (locator-resolved; not yet applied in dry-run) ===")
    let totalReplaceDelete = 0
    let totalInsert = 0
    for (const e of args.editsPreview) {
      const targetSpan =
        e.targetParaIndices.length === 0
          ? "(end-of-body)"
          : e.targetParaIndices.length === 1
            ? `#${e.targetParaIndices[0]}`
            : `#${e.targetParaIndices[0]}–#${e.targetParaIndices[e.targetParaIndices.length - 1]} (${e.targetParaIndices.length})`
      const insertNote = e.willInsertCount > 0 ? `; +${e.willInsertCount} new` : ""
      const replaceNote =
        e.willReplaceOrDeleteIndices.length > 0
          ? `; -${e.willReplaceOrDeleteIndices.length} replaced/deleted`
          : ""
      const containerNote = e.container === "cell" ? " [in cell]" : ""
      lines.push(`  edits[${e.index}] ${e.op} → ${targetSpan}${containerNote}${replaceNote}${insertNote}`)
      totalReplaceDelete += e.willReplaceOrDeleteIndices.length
      totalInsert += e.willInsertCount
    }
    lines.push(
      `  Total: ${args.editsPreview.length} ops; ${totalReplaceDelete} paragraphs replaced/deleted, ${totalInsert} new paragraphs inserted.`,
    )
    lines.push(
      "  Note: implicit-keep counts above already exclude paragraphs the edits[] pass will replace/delete.",
    )
    lines.push("")
  }
  if (args.samples.size > 0) {
    lines.push("=== Sample Affected Paragraphs (first per style) ===")
    for (const [styleId, samples] of args.samples) {
      lines.push(`  ${styleId}:`)
      for (const s of samples) {
        const notes = s.notes.length > 0 ? `  [${s.notes.join("; ")}]` : ""
        lines.push(
          `    #${s.paraIndex} via=${s.via}${s.patternSource ? ` /${s.patternSource}/` : ""}: "${s.textPreview}"${notes}`,
        )
      }
    }
    lines.push("")
  }
  if (args.dryRun) {
    lines.push("Dry run — no file written, no validation performed.")
    lines.push("Re-run without --dry-run to commit changes.")
  } else {
    lines.push("Validation: PASS")
    lines.push(`Output: ${args.output}`)
  }
  console.log(lines.join("\n"))
}
