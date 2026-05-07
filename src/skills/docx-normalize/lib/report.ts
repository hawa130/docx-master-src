import type { ImportResult } from "@core/template-import.ts"
import type {
  FlagRecord,
  RestyleSample,
  StyleConfigEntry,
  StyleResolutionEntry,
} from "./types.ts"

/* ------------- display helpers ------------- */

/**
 * Pick the user-facing typographic fields from a resolved style for
 * side-by-side display. Excludes mechanical fields (id, name, basedOn,
 * fromParagraph, overrides) the user wouldn't recognize.
 */
export function extractDisplayFields(def: StyleConfigEntry): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  // fontEastAsia first when set: in Chinese docs the CJK font is what users
  // perceive as "the font", so leading with it avoids the "why is the body
  // Arial?" double-take when the agent reviews the resolution block. When
  // only `font` (Latin/ASCII) is set, fontEastAsia is absent and the order
  // collapses to the natural one anyway.
  const interesting: (keyof StyleConfigEntry)[] = [
    "fontEastAsia", "font", "size", "bold", "italic", "color",
    "alignment", "lineSpacing", "lineRule", "spaceBefore", "spaceAfter",
    "firstLineIndent", "hangingIndent", "outlineLevel",
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
  implicitKeepByFingerprint: Map<string, { empty: number; nonEmpty: number; nonEmptySamples: string[] }>
  unstrippedShapesByStyle: Map<string, Map<string, number>>
  templateImport: ImportResult | null
}) {
  const lines: string[] = []
  lines.push(args.dryRun ? "=== Change Report (DRY RUN — no file written) ===" : "=== Change Report ===")
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
  lines.push(
    `Styles injected: ${args.injected.length} (${args.injected.map(annotate).join(", ")})`,
  )
  lines.push(
    `Styles updated:  ${args.updated.length} (${args.updated.map(annotate).join(", ")})`,
  )
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
        const samples = v.nonEmptySamples.length > 0
          ? `  e.g. ${v.nonEmptySamples.map((s) => `"${s}"`).join(" / ")}`
          : ""
        lines.push(`    ${fp}×${v.nonEmpty}${samples}`)
      }
    }
  }
  lines.push("")
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
  const mixedStyles = [...args.manualNumberingByStyle.entries()].filter(
    ([, m]) => m.size >= 2,
  )
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
  // Loud-fail on uncovered typed-prefix shapes: paragraphs assigned to a
  // numbered style whose leading text matched a known prefix shape that the
  // level's stripPrefixPatterns didn't cover. This is the "agent thought
  // they handled it but missed half the doc" case — auto-numbering will fire
  // AND the manual prefix will stay, producing "1. 1. Heading text" output.
  // Surface here, not silently in the existing Mixed-detected section,
  // because that section only fires when patterns ALREADY hit two shapes;
  // it can't see what's still uncovered.
  if (args.unstrippedShapesByStyle.size > 0) {
    lines.push(
      "Uncovered manual prefixes (will double-number — auto-number AND keep manual prefix):",
    )
    for (const [styleId, shapeMap] of args.unstrippedShapesByStyle) {
      const breakdown = [...shapeMap.entries()]
        .sort((a, b) => b[1] - a[1])
        .map(([shape, n]) => `"${shape}"×${n}`)
        .join(", ")
      lines.push(`  ${styleId}: ${breakdown}`)
    }
    lines.push(
      "  → Add the missing shape(s) to the style's numbering.levels[i].stripPrefixPatterns",
    )
    lines.push(
      "    (longer patterns first, e.g. [\"%1.%2\", \"%1.\"]).",
    )
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
