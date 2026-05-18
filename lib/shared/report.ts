import type { EditsPreviewEntry } from "@lib/edit/edit-engine.ts"
import type { ImportResult } from "@lib/apply/template-import.ts"
import type { PageSetupReport } from "@lib/apply/page-setup-mutation.ts"
import type {
  HeaderFooterReport,
  HeaderFooterBindingReport,
} from "@lib/apply/header-footer-mutation.ts"
import { type LineSpacingInput, parseLineSpacing, twipsToCmString } from "@lib/shared/units.ts"
import { sameValue, type VsDirectReport } from "@lib/shared/vs-direct.ts"
import type {
  FlagRecord,
  RestyleSample,
  StyleConfigEntry,
  StyleResolutionEntry,
} from "@lib/config/config-types.ts"

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
    "vertAlign",
    "alignment",
    "lineSpacing",
    "spaceBefore",
    "spaceAfter",
    "firstLineIndent",
    "hangingIndent",
    "outlineLevel",
  ]
  for (const k of interesting) {
    if (def[k] === undefined) continue
    // lineSpacing's three surface forms (number multiplier / "Npt" exact /
    // { atLeast } at-least) flatten to a "mode:value" string for diff parity
    // with cascade-resolved source values (which also flatten through
    // parseLineSpacing).
    if (k === "lineSpacing") {
      const ls = parseLineSpacing(def[k] as LineSpacingInput, "lineSpacing")
      out[k] = `${ls.mode}:${ls.value}`
    } else {
      out[k] = def[k]
    }
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

/** Compact Δ-line per the legend in the Style Resolution header.
 *  changes/+new/~matches stay in that order so the most actionable signal
 *  (changes) reads first. Returns empty string when nothing to show. */
function formatDeltaLine(
  resolved: Record<string, unknown>,
  prior: Record<string, unknown>,
): string {
  const changes: string[] = []
  const fresh: string[] = []
  const matches: string[] = []
  for (const [k, v] of Object.entries(resolved)) {
    if (!(k in prior)) {
      fresh.push(`+${k}=${shortVal(v)}`)
    } else if (sameValue(prior[k], v)) {
      matches.push(`~${k}`)
    } else {
      changes.push(`${k} ${shortVal(prior[k])}→${shortVal(v)}`)
    }
  }
  return [...changes, ...fresh, ...matches].join(", ")
}

function shortVal(v: unknown): string {
  if (typeof v === "string") return v
  if (typeof v === "boolean") return v ? "true" : "false"
  return String(v)
}

/** Render the vs-target-direct sub-block under a styleResolution entry.
 *  Hides the whole block when there's nothing to learn:
 *    - 0 target paragraphs → emit a single placeholder line.
 *    - All declared fields land in "new" (no override or redundant on any
 *      target) → skip — sparse declaration that adds without conflict. */
function renderVsDirect(r: VsDirectReport, lines: string[]): void {
  if (r.targetCount === 0) {
    lines.push(`    vs target direct: (no chrome targets; edits[]-inserts excluded from analysis)`)
    return
  }
  const interesting = r.fields.filter((f) => f.override > 0 || f.redundant > 0)
  if (interesting.length === 0) return // all-new sub-block hidden
  lines.push(`    vs target direct (${r.targetCount} paragraphs):`)
  for (const f of r.fields) {
    if (f.override === 0 && f.redundant === 0 && f.fresh === 0) continue
    const n = r.targetCount
    if (f.override > 0) {
      const tag = f.willStrip ? "" : " [mixed runs — direct stays]"
      const fromTo =
        f.overrideFrom !== undefined
          ? `${shortVal(f.overrideFrom)} → ${shortVal(f.declared)}`
          : `${shortVal(f.declared)}`
      lines.push(`      override:  ${f.field} ${fromTo} (${f.override}/${n})${tag}`)
    }
    if (f.redundant > 0) {
      lines.push(
        `      redundant: ${f.field}=${shortVal(f.declared)} (${f.redundant}/${n} already match direct)`,
      )
    }
    if (f.fresh > 0) {
      lines.push(
        `      new:       ${f.field}=${shortVal(f.declared)} (${f.fresh}/${n} no direct equivalent)`,
      )
    }
  }
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
  /** dry-run only: heuristic detection of typed-prefix residue in
   * `edits[]` Block paragraphs. `bound` = styleId has numbering scheme
   * (double-print risk); `unbound` = styleId lacks numbering (potential
   * block-enumeration miscategorisation). Chrome paragraphs through the
   * rule-routing path are covered separately by unstrippedByStyle. */
  manualNumberingDetected?: Map<
    string,
    { count: number; samples: string[]; kind: "bound" | "unbound" }
  >
  /** dry-run only: text snippet for each excluded paragraph so the agent can
   * spot index drift after document edits (exclude entries are bare numbers). */
  excludeSamples?: Array<{ index: number; snippet: string }>
  numberingBindings: Array<{
    styleId: string
    level: number
    lvlText: string
    suff: "tab" | "space" | "nothing"
  }>
  templateImport: ImportResult | null
  /** dry-run only: per-op preview of edits[] (locator-resolved, not mutated). */
  editsPreview: EditsPreviewEntry[]
  /** dry-run only: summary of the cross-reference post-pass (chapter SEQ
   * injection, caption standardize re-emit, predicted caption text). Null
   * when the run had no captions config or no captions/refs in edits. */
  captionsPreview: {
    chapterSeqsInjected: number
    standardizeReemitted: number
    freshlyEmitted: number
    samples: Array<{ identifier: string; text: string }>
  } | null
  /** dry-run only: heuristic flags for likely Pangu (CJK ↔ Latin/digit
   * literal-space) gaps in edits[] inserted text. Word's `autoSpace`
   * already inserts the visual gap between CJK and Latin glyphs;
   * stacking a manual ASCII space on top renders too wide. */
  panguWarnings?: Array<{ source: string; snippet: string; hit: string }>
  /** Per-section before/after for any `pageSetup` mutation. Absent when
   * pageSetup was not declared. */
  pageSetup?: PageSetupReport
  /** Header/footer parts generated this run. Absent when headerFooter was
   *  not declared. */
  headerFooter?: HeaderFooterReport
  /** Per-sectPr binding summary — section count touched + whether
   *  `<w:titlePg/>` was set. Absent when headerFooter was not declared. */
  headerFooterBinding?: HeaderFooterBindingReport
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
  if (args.pageSetup && args.pageSetup.sections.length > 0) {
    const ps = args.pageSetup
    lines.push(
      `Page setup: applies to ${ps.sections.length} section(s); ${ps.touchedCount} mutated.`,
    )
    for (const sec of ps.sections) {
      if (!sec.changed) continue
      const diffs = pageSetupDiff(sec.before, sec.after)
      if (diffs.length === 0) continue
      lines.push(`  Section ${sec.index}: ${diffs.join("; ")}`)
    }
    lines.push("")
  }
  if (args.headerFooter && args.headerFooter.parts.length > 0) {
    const hf = args.headerFooter
    const bind = args.headerFooterBinding
    const flags: string[] = []
    if (hf.groups.some((g) => g.hasFirst)) flags.push("titlePg")
    if (hf.hasEven) flags.push("evenAndOddHeaders")
    const flagSuffix = flags.length > 0 ? ` (flags: ${flags.join(", ")})` : ""
    lines.push(
      `Header/footer: ${hf.parts.length} part(s) generated; bound to ${bind?.sectionCount ?? 0} section(s)${flagSuffix}.`,
    )
    for (const p of hf.parts) {
      const extras: string[] = [`${p.blockCount} block(s)`]
      if (p.hasHyperlinks) extras.push("hyperlink")
      lines.push(`  ${p.surface}.${p.variant} → ${p.partName} (rId=${p.rId}, ${extras.join(", ")})`)
    }
    for (const warning of hf.separatorWarnings) {
      lines.push(`  warning: ${warning}`)
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
  // Dry-run only: echo each excluded paragraph's leading text so the agent
  // can verify the indices still point at what they intended. exclude is a
  // bare number list and silently drifts when document order shifts.
  if (args.excludeSamples && args.excludeSamples.length > 0) {
    lines.push(`Excluded paragraphs (${args.excludeSamples.length}; will not be touched):`)
    const cap = 15
    for (const s of args.excludeSamples.slice(0, cap)) {
      lines.push(`  #${s.index}  "${s.snippet}"`)
    }
    if (args.excludeSamples.length > cap) {
      lines.push(`  … (${args.excludeSamples.length - cap} more)`)
    }
    lines.push("  → Verify these indices still match — if document order shifted,")
    lines.push("    exclude entries silently aim at the wrong paragraphs.")
    lines.push("")
  }
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
    const bound = [...args.manualNumberingDetected].filter(([, v]) => v.kind === "bound")
    const unbound = [...args.manualNumberingDetected].filter(([, v]) => v.kind === "unbound")
    lines.push("Manual numbering detected in edits[] paragraphs:")
    if (bound.length > 0) {
      for (const [styleId, info] of bound) {
        const samples = info.samples.map((s) => `"${s}"`).join(" / ")
        lines.push(`  ${styleId} (bound to numbering): ${info.count} paragraphs  e.g. ${samples}`)
      }
      lines.push(
        "    → Drop the typed prefix from `text` — the scheme's lvlText already emits the marker.",
      )
    }
    if (unbound.length > 0) {
      for (const [styleId, info] of unbound) {
        const samples = info.samples.map((s) => `"${s}"`).join(" / ")
        lines.push(`  ${styleId} (no numbering binding): ${info.count} paragraphs  e.g. ${samples}`)
      }
      lines.push("    → If these are list items, declare a list-bound style (e.g. ListNumber)")
      lines.push("      and let the numbering scheme emit the marker. Keep the typed form")
      lines.push("      only when the user prompt explicitly asked for typed prefixes.")
    }
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
    const hasAnySpec = args.styleResolutions.some((r) => r.userSpec !== null)
    if (hasAnySpec) {
      lines.push("  The script does not parse natural language. For styles with a")
      lines.push("  user spec, compare it to the agent-resolved fields by eye —")
      lines.push("  any mismatch means the agent's translation needs adjustment.")
      lines.push("  Styles without a spec are still listed so the resolved fields")
      lines.push("  are auditable.")
    } else {
      // No specs anywhere — one summary line beats restating "(none)" N times.
      lines.push(
        "  No `requirements` entries declared — Agent Resolved fields below reflect the structured `styles[]` declarations directly.",
      )
    }
    lines.push("")
    // The Δ-line legend is only useful when at least one styleResolution
    // has a prior state to diff against. All-fresh runs (every style is
    // newly installed) emit no Δ line, so the legend is dead load — skip.
    const hasAnyPrior = args.styleResolutions.some((r) => r.priorState !== null)
    if (hasAnyPrior) {
      lines.push("  Δ-line vs source (when styleId existed pre-apply):")
      lines.push("    A→B    field changed; source had A, declaration sets B")
      lines.push("    +field new declaration; source didn't have this field")
      lines.push("    ~field declaration matches source's cascade value (may be redundant)")
      lines.push("")
    }
    for (const r of args.styleResolutions) {
      const freshTag = r.priorState === null ? "  [fresh]" : ""
      lines.push(`  ${r.styleId}${freshTag}`)
      if (hasAnySpec) {
        if (r.userSpec !== null) {
          lines.push(`    User specified: "${r.userSpec}"`)
        } else {
          lines.push(`    User specified: (none — no requirements entry)`)
        }
      }
      lines.push(`    Agent resolved: ${formatResolvedFields(r.resolved)}`)
      if (r.priorState !== null) {
        const delta = formatDeltaLine(r.resolved, r.priorState)
        if (delta) lines.push(`    Δ vs source:    ${delta}`)
      }
      if (r.priorUsage && r.priorUsage > 0) {
        lines.push(
          `    Existing usage: ${r.priorUsage} paragraph(s) — will re-render with the new definition.`,
        )
      }
      if (r.warnings) {
        for (const wMsg of r.warnings) {
          lines.push(`    note: ${wMsg}`)
        }
      }
      if (r.vsDirect) {
        renderVsDirect(r.vsDirect, lines)
      }
    }
    lines.push("")
  }
  if (args.editsPreview.length > 0) {
    lines.push(
      "=== Edits Preview (locator-resolved; applied in memory, not written to disk in dry-run) ===",
    )
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
      lines.push(
        `  edits[${e.index}] ${e.op} → ${targetSpan}${containerNote}${replaceNote}${insertNote}`,
      )
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
  if (args.captionsPreview) {
    const cp = args.captionsPreview
    lines.push("=== Captions / Cross-Refs Preview (dry-run) ===")
    lines.push(
      `  Chapter SEQs injected: ${cp.chapterSeqsInjected}  (hidden auto-counters on outline paragraphs)`,
    )
    lines.push(
      `  Caption paragraphs re-emitted: ${cp.standardizeReemitted}  (existing captions rebuilt against current config)`,
    )
    lines.push(
      `  Caption paragraphs freshly emitted: ${cp.freshlyEmitted}  (new captions inserted by edits[])`,
    )
    if (cp.samples.length > 0) {
      lines.push(`  Predicted text (first ${cp.samples.length}):`)
      for (const s of cp.samples) lines.push(`    [${s.identifier}] "${s.text}"`)
    } else {
      lines.push(`  Predicted text: (no caption samples — config has captions but body emits none)`)
    }
    lines.push("")
  }
  if (args.panguWarnings && args.panguWarnings.length > 0) {
    const w = args.panguWarnings
    lines.push("=== Possible Pangu spacing in author-supplied text ===")
    lines.push(
      "  Word's autoSpace handles CJK ↔ Latin/digit gaps; typed ASCII spaces stack on top and render too wide.",
    )
    const shown = w.slice(0, 5)
    for (const entry of shown) {
      lines.push(`  ${entry.source}  ...${entry.snippet}...   (matched "${entry.hit}")`)
    }
    if (w.length > shown.length) {
      lines.push(`  (... ${w.length - shown.length} more not shown)`)
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

/** Per-section before/after summary lines. Returns only fields that actually
 *  changed so the report stays sparse. Twips → mm + cm for readability. */
function pageSetupDiff(
  before: PageSetupReport["sections"][number]["before"],
  after: PageSetupReport["sections"][number]["after"],
): string[] {
  const out: string[] = []
  if (before.paperSize !== after.paperSize) {
    out.push(`paper ${before.paperSize ?? "?"} → ${after.paperSize ?? "?"}`)
  }
  if (before.orientation !== after.orientation) {
    out.push(`orient ${before.orientation} → ${after.orientation}`)
  }
  const mEdges: Array<keyof typeof before.margins> = ["top", "bottom", "left", "right"]
  const mDiffs: string[] = []
  for (const e of mEdges) {
    const b = before.margins[e]
    const a = after.margins[e]
    if (b !== a) mDiffs.push(`${e} ${twipsToCmString(b)} → ${twipsToCmString(a)}`)
  }
  if (mDiffs.length > 0) out.push(`margins(${mDiffs.join(", ")})`)
  return out
}
