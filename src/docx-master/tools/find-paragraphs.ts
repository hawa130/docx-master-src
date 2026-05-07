/**
 * find_paragraphs <docx> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]
 *
 * Cross-document text search returning matching paragraphs with index,
 * fingerprint, and text preview. Closes the gap where the agent had to
 * grep the overview skeleton manually.
 *
 * Common uses:
 *   --regex "^图\\s*\\d+[-.]\\d+"      → all figure caption candidates
 *   --regex "^表\\s*\\d+[-.]\\d+"      → all table caption candidates
 *   --regex "^\\[\\d+\\]"               → bibliography entries
 *   --regex "^(关键词|Keywords?)\\s*[:：]"  → keyword lines
 *   --fingerprint A --regex "^\\d+\\."   → numbered list items in a specific style
 *
 * Filtering by fingerprint (--fingerprint) is useful when the same regex
 * pattern matches multiple roles and you want to scope to one visual class.
 */
import { loadDocx } from "@core/load.ts"
import { pad } from "@core/format.ts"

async function main() {
  const argv = process.argv.slice(2)
  let regex: string | undefined
  let flags: string | undefined
  let limit = 50
  let fingerprintFilter: string | undefined
  const positional: string[] = []
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!
    if (a === "--regex") {
      regex = argv[++i]
    } else if (a === "--flags") {
      flags = argv[++i]
    } else if (a === "--limit") {
      const v = argv[++i]
      if (!v || isNaN(parseInt(v, 10))) {
        console.error("--limit requires a positive integer")
        process.exit(1)
      }
      limit = parseInt(v, 10)
    } else if (a === "--fingerprint") {
      fingerprintFilter = argv[++i]
    } else if (a.startsWith("--")) {
      console.error(`unknown flag: ${a}`)
      process.exit(1)
    } else {
      positional.push(a)
    }
  }
  const file = positional[0]
  if (!file) {
    console.error(
      "Usage: node scripts/find_paragraphs.js <docx-path> --regex <pat> [--flags <flags>] [--limit N] [--fingerprint X]",
    )
    process.exit(1)
  }
  if (!regex) {
    console.error("--regex is required")
    process.exit(1)
  }

  let pattern: RegExp
  try {
    pattern = new RegExp(regex, flags ?? "")
  } catch (err) {
    console.error(`Invalid regex /${regex}/${flags ?? ""}: ${(err as Error).message}`)
    process.exit(1)
  }

  try {
    const doc = await loadDocx(file)
    const matches = doc.paragraphs.filter((p) => {
      if (fingerprintFilter && p.fingerprint !== fingerprintFilter) return false
      return pattern.test(p.text)
    })

    const lines: string[] = []
    lines.push(
      `Pattern: /${regex}/${flags ?? ""}` +
        (fingerprintFilter ? `  filter: fingerprint=${fingerprintFilter}` : ""),
    )
    lines.push(
      `Matches: ${matches.length}${
        matches.length > limit ? ` (showing first ${limit})` : ""
      }`,
    )
    lines.push("")
    for (const p of matches.slice(0, limit)) {
      const txt =
        p.text.length > 80 ? `${p.text.slice(0, 77)}…` : p.text
      lines.push(`  #${pad(p.index)} [${p.fingerprint}]  "${txt}"`)
    }
    console.log(lines.join("\n"))
  } catch (err) {
    console.error(`Error: ${(err as Error).message}`)
    process.exit(1)
  }
}

main()
