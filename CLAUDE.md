# Working in this repo

This repo builds **one** Word (.docx) automation skill, `docx-master`. The agent-facing contract is `skill/SKILL.md`; this file is for working *on* the project.

Surfaces today:

- **`apply`** — the unified writer. Single CLI, one config combining two shapes:
  - **standardize-shape** — role-based whole-doc reshape (paragraph classification, named-style injection, numbering migration, template import)
  - **edit-shape** — location-based surgical changes (replace / insert / delete paragraphs, table-cell content, image embedding, optional tracked changes)
- **`audit`** — read-only workflow producing a violation report

Future extensions add new config blocks inside `apply`, not new sub-commands.

**Keep this file in sync.** Tool names, build commands, file paths, and the lessons below are referenced concretely. When any of them changes, update here in the same commit — stale references mislead future maintainers and the next agent reviewing the design.

## Layout

- `skill/` — publishable skill bundle source: `SKILL.md` (agent-facing contract), `references/` (on-demand detail), `tools/` (TS source for agent-callable CLIs; one file = one `tsdown.config.ts` entry)
- `lib/` — non-tool TS modules grouped by concern (xml / parse / config / apply / edit / shared). Reachable via `@lib/*` alias. Imported by tools, never built as a script entry. `ls lib/` for the current breakdown.
- `test/fixtures/` — sample .docx files for manual testing
- `dist/` — build output (gitignored): `docx-master/` staged bundle + `docx-master.zip`
- `build-skill.ts` — packages the staged dir into the .skill zip

All `lib/` and `skill/tools/` imports go through `@lib/*` (declared in `tsconfig.json` paths + `tsdown.config.ts`), including lib-to-lib. No relative paths — group-prefixed imports survive moves between groups. `skill/tools/` is CLI entry-points only; anything imported but never invoked goes in `lib/`.

## Commands

| Task | Command |
|---|---|
| Build TypeScript → dist/ | `bun run build` |
| Build + stage bundle + zip | `bun run build:skill` |
| Watch | `bun run build:watch` |
| Type-check (tsc, no emit) | `bun run typecheck` |
| Lint (oxlint, type-aware, --deny-warnings) | `bun run lint` |
| Auto-fix lint where possible | `bun run lint:fix` |
| Format (oxfmt; markdown excluded) | `bun run fmt` |
| Format check (no write) | `bun run fmt:check` |

No automated tests — run scripts against `test/fixtures/*.docx` manually after changes. After edits to `skill/` or `lib/`, always rebuild and verify `dist/docx-master/` reflects the change before claiming done.

## Dependency policy

`bunfig.toml` sets `minimumReleaseAge = 259200` (3 days). New or upgraded packages whose latest matching version was published in the last 3 days are refused at resolution time — quarantine window against a compromised release. Doesn't affect `bun install --frozen-lockfile`; only fires when resolving fresh versions (`bun add`, `bun update`, unlocked install).

When `bun add` fails the age gate, check why the package was published so recently before reaching for an exclude. If you must take the version (security fix, blocking bug), add the specific package to `minimumReleaseAgeExcludes` in `bunfig.toml`, install, then remove the exclude. Don't let the exclude list grow as a default workaround.

## Adding a tool

1. Create `skill/tools/<file>.ts`. Use `@lib/...` for everything — OOXML primitives, skill engine helpers, config schema, CLI scaffolding all live under one alias.
2. Add a `<scriptName>: "<file>.ts"` entry to the `tools` map in `tsdown.config.ts`.
3. If the tool is a sub-command surface the agent should route to, add a row to the SKILL.md tool table.
4. `bun run build:skill` produces `dist/docx-master/` and `dist/docx-master.zip`.

## Periodic audits via `skill-creator`

After a multi-commit feature push or before a release, spawn a `general-purpose` subagent that invokes the `skill-creator` skill in **read-only** mode (no file edits), pointed at `skill/` + `dist/docx-master/`.

skill-creator's checklists (Anatomy of a Skill / Progressive Disclosure / Writing Patterns / Description Optimization) catch:

- Stale `references/` content
- Anti-pattern leakage in docs
- Outdated regex catalogs in rarely-edited directories
- Checklist items that human review skims past

Evaluate findings critically — skill-creator over-suggests. Act on the real ones, defer or decline the rest.

## Design principles

### Minimum tokens for maximum effect

Every line of agent-facing doc loads on every invocation. Examples are bloat unless they cover a non-obvious shape or a real footgun the rule alone won't convey. "What to avoid" lists and rebuttal-of-rationalization sections suffer the same way — each item must earn its keep by covering a class of mistakes agent would otherwise make. Defensive bloat compounds across turns.

Test for any line: could a reader derive it from the rule + surrounding context? Yes → cut. No → keep. Same when distilling a fix into a principle here: state the conclusion, drop the illustration. If a reader can't picture the failure without an example, the rule isn't general enough yet.

Reactive maintenance is the main vector for bloat: each external feedback adds a paragraph; over rounds the doc accretes restatements. When feedback comes in, ask first whether the agent's own judgment covers it. Periodically zoom out and audit holistically — patch-by-patch additions each looked justified; the sum often isn't.

### Address root cause; consolidate to canonical locations

When a class of failure recurs (each iteration breaks on a different specific case but the same axis — agent over-declaring fields, over-typing markers, mis-classifying chrome vs slot), patching each manifestation enumerates instances without removing the axis. The structural fix reframes the schema or workflow so the bad choice isn't reachable — e.g., making `fromParagraph` required for represented roles removed the agent's opportunity to over-declare typography, retiring several "warn against X" patches at once. Ask: is the next failure on this axis a different instance, or a different axis? Different instance → look for the structural reframe.

Within the doc bundle: SKILL.md is the router (decision + pointer); `references/*` carry detail. A rule lives in one canonical location; other places point. Cross-doc duplication accretes during patch cycles — when consolidating, push detail back to its owning ref doc and leave a one-line pointer in callers.

### Tools expose visible facts; agents make role judgments

Tools expose what a human reader of the artifact can see — font, size, weight, color, alignment, indent, numbering markers, text content. The agent classifies roles from those facts.

Don't pre-digest hidden metadata or natural language into role hints in default output. A "convenient pre-classification" gets accepted as ground truth and bypasses agent judgment — the same judgment that would have caught upstream errors. Hidden metadata stays available *on demand* (`inspect_range` / `inspect_style_def`), where the agent has already started reasoning and consumes it as one input among many.

Litmus test for any new "helpful hint": could a human reader derive this from the artifact's normal rendering? Yes → expose. No → on demand only, never as a role hint in default output.

### Examples illustrate, they're not triggers

When SKILL.md presents "intent → path / tool / option" mappings, the LLM pattern-matches surface phrasings instead of understanding intent. Write concept-first and mark example phrasings as illustrative. Avoid lookup-table forms (`If user says | Pick this`) and the literal words *Triggers* / *Keywords*.

Same with concrete Bad/Good code or text excerpts in skill docs: they freeze the rule to *this* document's terms, and the agent applies it as a literal match instead of recognizing the underlying category. State the rule in general terms, then illustrate only when the failure shape isn't derivable from the rule. If a reader can't picture the failure without the example, the rule isn't general enough yet — fix the rule, don't add another example.

### Verification must check against intent, not interpretation

If a check grades the system's output against the same system's interpretation of the input, it's a tautology and passes regardless of correctness. Real verification compares against ground truth: human-readable side-by-side (e.g. Style Resolution shows raw user text + resolved fields for visual review), or output re-parsed against an independent invariant (apply_styles validates by re-reading the produced docx).

### Tools follow the on-demand pattern

Each tool has one focused job. Default outputs stay scannable; deep info is one tool call away.

### Mechanical correctness is the script's job

LLMs are bad at byte-level work; scripts must guarantee these and never bend them under refactoring pressure:

- Paragraph walks go through one of two canonical helpers — `walkBodyParagraphs` (body + tbl/tr/tc; used by the cross-ref pipeline and inspection tools) or `walkIndexedParagraphs` (matches DocumentParser scope: body + layout-table cells, skips data/form). Reimplementing the walk inline is a bug factory — divergent scope between passes was the cause of multiple Phase 1 dogfood failures (chapter SEQ injected against a different set of paragraphs than the counter sim saw)
- XML-namespace-correct mutation of `styles.xml` / `numbering.xml` / `document.xml`
- Cross-run formatting preservation (smart-strip's uniform-vs-mixed rule)
- Character-indent semantics — `firstLineChars` round-trips as `"Nchar"`, fixed twips as `"Npt"`
- Dominant-run selection in `fromParagraph` skips numbering-prefix-only runs
- Leading-prefix strip accumulates across consecutive `<w:t>` runs; Word splits hand-edited paragraphs mid-prefix, so per-run regex testing silently misses some
- numId migration on template import uses fresh IDs to avoid collision
- The original file is never modified; every applying CLI writes a fresh copy and validates before keeping it
- Manual structural prefix detection is part of the standardize survey: pre-existing typed prefixes (`一、` / `（一）` / `1.1` / `第N章`) inside otherwise-styled headings are candidates for `stripPrefixPatterns` conversion to auto-numbering, not chrome to preserve
- **Edit side**: locators resolve to Element refs *before* any mutation, so subsequent ops survive DOM rearrangement; stale-element rejection guards against an op targeting a paragraph an earlier op removed
- **Edit side**: blocker scan refuses edits inside existing `<w:ins>` / `<w:del>` / `<w:fldChar>` regions / `<w:sdt>` controls — fields and revisions don't survive ad-hoc paragraph rewriting
- **Edit side**: track-changes mode snapshots previous `<w:rPr>` / `<w:pPr>` *before* mutating the live element, so `<w:rPrChange>` / `<w:pPrChange>` carry the genuine prior state (cloning after mutation would record the new state as the snapshot)
- **Edit side**: insertions into the body land before any trailing `<w:sectPr>` so the section descriptor stays last; cell-container insertions append plainly
- **Image asset path**: each image registration touches three coordinated parts (`word/media/`, `[Content_Types].xml`, `word/_rels/document.xml.rels`); all three are staged before write so a failure leaves no half-registered asset

When changing any of these, verify against `test/fixtures/` and inspect the output zip.

## Implementation workflow

Process patterns that converge faster on non-trivial changes.

### Ground claims in the artifact

Factual claims about codebase state come from grep / Read / running the tool — not memory, not inference. Applies equally to your own statements and to subagent feedback; both can be wrong, the code and the validator can't. A minute of verification saves an hour of revert.

### Audit existing conventions before designing new

When adding to an existing concept layer (schemas, helpers, file conventions, doc structure), read the nearest 2-3 siblings first. Catches helpers that already exist, prevents API-shape drift between siblings. When an existing convention is imperfect, weigh refactoring against blast radius — convenience-only consistency rarely clears the bar for breaking change.

### Spawn parallel review agents for non-trivial work

Dispatch two reviewers in parallel — one for agent-facing outputs (docs, API contracts), one for code (bugs, reuse, edges). Different agents have different blind spots; reusing one shares its blind spot. Iterate until both report no new issues — typically 2–3 rounds with monotonically decreasing severity. If round 3 still surfaces critical issues, the design has a structural problem worth reframing, not patching.

### Exercise new × existing combinations

Before declaring done, test the new feature combined with horizontal-cutting features already in the system. Pre-existing latent bugs hide in combinations; new-in-isolation tests miss them. Same at the doc level — grep for redundancy across the whole doc set after each phase, because canonical-location decay accumulates between phases.

## Commit style

- One line, title only (what changed). No body.
- Keep commits atomic — one logical change per commit; split unrelated work into separate commits.
- No `co-authored-by` tags.
