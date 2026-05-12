# Working in this repo

This repo builds **one** Word (.docx) automation skill, `docx-master`. Two writing sub-commands ship today — `standardize` (role-based whole-doc reshape: paragraph classification, named-style injection, numbering migration, template import) and `edit` (location-based surgical edits: replace/insert/delete paragraphs, table-cell content, image embedding, optional Word tracked changes) — plus the read-only `audit`. Future sub-commands (content authoring adapters, batch fragment imports) will add to the same SKILL.md as routed entries. The agent-facing contract is `skill/SKILL.md`; this file is for working *on* the project.

**Keep this file in sync.** Tool names, build commands, file paths, and the lessons below are referenced concretely. When any of them changes, update here in the same commit — stale references mislead future maintainers and the next agent reviewing the design.

## Layout

```
skill/                             the publishable skill bundle source.
skill/SKILL.md                       agent-facing contract (router + invariants)
skill/references/                    on-demand reference docs (progressive disclosure)
skill/tools/                         TS source for CLIs the agent invokes directly
                                       (each file = one entry in tsdown.config.ts)
lib/                               every non-tool TypeScript module, grouped
                                     by concern. Reachable via the `@lib/*`
                                     alias. Imported by tools, never built as
                                     a script entry.
  lib/xml/                           OOXML/zip primitives (xml-utils,
                                       xml-order, reader, load, docx-plumbing)
  lib/parse/                         read-side parsing (document-parser,
                                       style-resolver, fingerprint,
                                       table-classifier, format,
                                       manual-numbering-detect,
                                       section-metrics, types)
  lib/config/                        zod schemas + derived types for both
                                       sub-commands (config-schema,
                                       config-types, edit-config-schema,
                                       edit-types)
  lib/apply/                         standardize sub-command engine
                                       (apply-styles orchestrator,
                                       style/numbering/para/list mutation,
                                       template-import)
  lib/edit/                          edit sub-command engine (edit-engine,
                                       locator, text-search, blockers,
                                       fragment-emit, track-changes,
                                       image-asset, bookmark, field-ref,
                                       table-emit, math/)
  lib/shared/                        cross-engine helpers (cli-helpers,
                                       docx-validate, report)
                                     Two type files split by concern:
                                       - `lib/parse/types.ts` for OOXML /
                                         parser types (NS, ParsedParagraph,
                                         ...)
                                       - `lib/config/config-types.ts` for
                                         config-derived + internal data
                                         shapes (ApplyConfig, ApplyContext,
                                         ...)
test/fixtures/                     sample .docx files for manual testing
dist/docx-master/                  staged skill bundle (SKILL.md + references/ + scripts/)
dist/docx-master.zip               zipped bundle ready to publish
build-skill.ts                     packages staged dir into the .skill zip
```

Tools and `lib/` modules import internal code via the `@lib/*` alias (declared in `tsconfig.json` paths and `tsdown.config.ts` alias) — including lib-to-lib imports. Don't use relative paths (`./foo.ts`, `../xml/foo.ts`); keep imports group-prefixed (`@lib/xml/foo.ts`) so moving a file between groups only touches the import-site path, not its style. The `skill/tools/` directory is exclusively for files built as agent-callable CLIs; anything imported but never invoked goes in `lib/`.

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

## Adding a tool

1. Create `skill/tools/<file>.ts`. Use `@lib/...` for everything — OOXML primitives, skill engine helpers, config schema, CLI scaffolding all live under one alias.
2. Add a `<scriptName>: "<file>.ts"` entry to the `tools` map in `tsdown.config.ts`.
3. If the tool is a sub-command surface the agent should route to, add a row to the SKILL.md tool table.
4. `bun run build:skill` produces `dist/docx-master/` and `dist/docx-master.zip`.

## Periodic audits via `skill-creator`

After a multi-commit feature push or before a release, spawn a subagent that invokes the `skill-creator` skill to audit the bundle. Pattern: `Agent` tool with `general-purpose` subagent; prompt asks it to invoke `skill-creator` and audit `skill/` + `dist/docx-master/`, **read-only** (no file edits). The framework's checklists (Anatomy of a Skill / Progressive Disclosure / Writing Patterns / Description Optimization) catch stale `references/` content, anti-pattern leakage in docs, and checklist items that human review skims past — especially in directories that get edited rarely and accumulate wrong-tooling examples or outdated regex catalogs. Evaluate findings critically (skill-creator can over-suggest); act on real ones, defer or decline the rest.

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

## Commit style

Title: what changed. Body: why, in 1–3 short paragraphs max — enough that the next maintainer wouldn't accidentally revert the decision. The diff already shows what; don't restate it. No "Net: X → Y lines" stats, no per-bullet narration of each change, no recap of content already in CLAUDE.md / SKILL.md. Don't include `co-authored-by` tags.
