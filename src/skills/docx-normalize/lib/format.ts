/**
 * Display helpers shared across the inspector CLIs. These are the truly
 * identical formatters lifted from multiple tools — per-tool variants of
 * formatRPr/formatPPr stay local because each inspector orders fields and
 * picks fonts slightly differently for its specific use case, and the exact
 * output shape is part of the agent-facing contract.
 */

/** Pad a 1-based paragraph index to 3 digits — "001", "042", "123". */
export function pad(n: number): string {
  return n.toString().padStart(3, "0")
}

/** Collapse internal whitespace and clip with an ellipsis. */
export function truncate(s: string, n: number): string {
  const collapsed = s.replace(/\s+/g, " ").trim()
  return collapsed.length <= n ? collapsed : collapsed.slice(0, n) + "…"
}

/** Recognize standard paper sizes by their twip dimensions. */
export function paperName(width: number, height: number): string {
  // both portrait and landscape A4 are listed so either orientation matches
  const known: Array<[string, number, number]> = [
    ["A4", 11906, 16838],
    ["A4", 16838, 11906],
    ["A3", 16838, 23811],
    ["A5", 8392, 11906],
    ["Letter", 12240, 15840],
    ["Legal", 12240, 20160],
  ]
  for (const [name, w, h] of known) {
    if (Math.abs(width - w) < 50 && Math.abs(height - h) < 50) return name
  }
  return "Custom"
}

/** Render Word's spacing/@line + spacing/@lineRule pair as human text. */
export function formatLineSpacing(line: number, rule: string | undefined): string {
  const r = rule || "auto"
  if (r === "exact") return `${line / 20}pt fixed`
  if (r === "atLeast") return `${line / 20}pt atLeast`
  return `${parseFloat((line / 240).toFixed(2))}×`
}

/** Twips → millimeters, rounded to 1dp (Word page-setup-friendly). */
export function tw2mm(t: number): number {
  return +(t / 56.6929).toFixed(1)
}
