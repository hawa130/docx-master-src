# Working in this repo

This repo builds a family of Word (.docx) automation skills. The first one shipped is `docx-normalize` (paragraph classification + named-style injection); planned siblings are `docx-cleanup` (accept changes, strip metadata) and `docx-augment` (insert captions, manage cross-references). Each skill ships as its own bundle. Per-skill `SKILL.md` files are the agent-facing contracts; this file is for working *on* the project.

**Keep this file in sync.** Tool names, build commands, file paths, and the lessons below are referenced concretely. When any of them changes, update here in the same commit — stale references mislead future maintainers and the next agent reviewing the design.

## Layout

```
src/core/                          shared across all skills:
                                     parsing / style / fingerprint / template-import,
                                     plus display formatters (format.ts) for any
                                     skill's print tools
src/skills/<name>/                 one directory per skill, fully self-contained:
src/skills/<name>/SKILL.md           the skill spec the agent reads at runtime
src/skills/<name>/references/        on-demand reference docs (progressive disclosure)
src/skills/<name>/tools/             TS source for CLIs the agent invokes directly
                                       (each file = one entry in tsdown.config.ts)
src/skills/<name>/lib/               skill-internal modules (CLI scaffolding, the
                                       skill's engine, anything imported but never
                                       built as a script entry)
test/fixtures/                     sample .docx files for manual testing
dist/<name>/                       staged skill bundle (SKILL.md + references/ + scripts/)
dist/<name>.zip                    zipped bundle ready to publish
build-skill.ts                     packages staged dirs into .skill zips
```

Skill tools import shared modules via the `@core/*` alias (declared in `tsconfig.json` paths and `tsdown.config.ts` alias). Don't use relative `../../../core/...` paths.

## Commands

| Task | Command |
|---|---|
| Build TypeScript → dist/ (all skills) | `bun run build` |
| Build + stage skill bundle + zip (single skill, or `--all`) | `bun run build:skill [<name>] [--all]` |
| Watch | `bun run build:watch` |
| Type-check (tsc, no emit) | `bun run typecheck` |

No automated tests — run scripts against `test/fixtures/*.docx` manually after changes. After edits to `src/skills/<name>/` or shared `src/core/`, always rebuild and verify `dist/<name>/` reflects the change before claiming done.

## Adding a new skill

1. `mkdir -p src/skills/<name>/{tools,references}` and write `src/skills/<name>/SKILL.md`
2. Add the skill's tool entries to `tsdown.config.ts` (each entry maps a script name → its `tools/<file>.ts`)
3. Tools import shared modules via `@core/...`, skill-internal helpers via `../lib/...`
4. **`tools/` is exclusively for files that get built as agent-callable CLIs.** Anything that's only imported (CLI scaffolding, formatters) goes in `lib/` so the build entry list and the agent's mental model of "tools" stay aligned.
5. `bun run build:skill <name>` produces `dist/<name>/` and `dist/<name>.zip`

## Periodic audits via `skill-creator`

After a multi-commit feature push or before a release, spawn a subagent that invokes the `skill-creator` skill to audit a specific skill bundle. Pattern: `Agent` tool with `general-purpose` subagent; prompt asks it to invoke `skill-creator` and audit `src/skills/<name>/` + `dist/<name>/`, **read-only** (no file edits). The framework's checklists (Anatomy of a Skill / Progressive Disclosure / Writing Patterns / Description Optimization) catch stale `references/` content, anti-pattern leakage in docs, and checklist items that human review skims past — especially in directories that get edited rarely and accumulate wrong-tooling examples or outdated regex catalogs. Evaluate findings critically (skill-creator can over-suggest); act on real ones, defer or decline the rest.

## Design principles

### Tools expose visible facts; agents make role judgments

Tools should expose what a human reader of the artifact can see — in this skill: font, size, weight, color, alignment, indent, numbering markers, text content. The agent classifies roles from those visible facts.

**Don't pre-digest hidden metadata or natural language into role hints in default output.** When tooling presents a "convenient pre-classification" (a parsed translation, a metadata-derived label), the LLM accepts it as ground truth and bypasses its own judgment — the same judgment that would have caught upstream errors. This anti-pattern showed up twice in this project (a regex parser for Chinese typography, an `[LN]` outline-level hint on the fingerprint summary); both were removed for the same reason.

Hidden metadata can still be exposed — but as **raw data on demand**, when the agent explicitly calls `inspect_range` / `inspect_style_def`. On-demand exposure means the agent has already started reasoning and is using the metadata as one input among many. Pre-packaging it into a default summary means the tool is doing the reasoning for the agent.

**Litmus test for any new "helpful hint":** could a human reader derive this from the artifact's normal rendering, without opening the underlying file format? Yes → expose it. No → only on demand, never as a role hint in default output.

### Examples illustrate, they're not triggers

When SKILL.md presents "intent → path / tool / option" mappings, the LLM pattern-matches keywords instead of understanding intent — the same surface phrase routinely lands in different branches depending on what the user actually wants. Write concept-first and mark example phrasings as illustrative. Avoid lookup-table forms (`If user says | Pick this`) and the literal words *Triggers* / *Keywords* — both signal "stop thinking, start matching." Same anti-pattern as "Tools expose visible facts; agents make role judgments," applied to the doc itself.

### Verification must check against intent, not interpretation

If a check grades the system's output against the same system's interpretation of the input, it's a tautology and passes regardless of correctness. The removed `requirements-parser.ts` had this flaw — it parsed user text into fields, then verified the script wrote those same fields; the parser's own misreading was invisible to the check.

Real verification compares against ground truth: human-readable side-by-side display (Style Resolution shows raw user text + agent-resolved fields for visual review), or output re-parsed against an independent invariant (apply_styles validates by re-reading the produced docx).

### Capture lessons here as you learn them

When a fix or design decision teaches a principle the next maintainer should know, distill it as a brief rule above. Bar: would another maintainer save time reading it? — if no, commit history is enough. A root cause recurring in two different shapes is a strong signal a principle has earned a slot.

### SKILL.md is the agent's runtime context — keep it lean

Every line is loaded into the agent's context on every invocation. Encode only what the agent can't derive: technical invariants the scripts depend on, workflow anchors, tool references. Cut restated points, disambiguation tables for things an LLM infers from context, examples of output the tool prints itself. Detailed schema lives in `references/`, loaded on demand.

Reactive maintenance is a trap: each external feedback adds a paragraph; over multiple rounds the doc bloats with restatements. When feedback comes in, ask first: *can the agent's own judgment cover this?* If yes, don't add. Periodically zoom out and audit holistically — patch-by-patch additions all looked justified, but the sum may not be.

### Tools follow the on-demand pattern

Each tool has one focused job. Don't add always-on enrichment to default outputs — prefer a new focused script. Default outputs stay scannable; deep info is one tool call away.

### Mechanical correctness is the script's job

LLMs are bad at byte-level work; scripts must guarantee these and never bend them under refactoring pressure:

- XML-namespace-correct mutation of `styles.xml` / `numbering.xml` / `document.xml`
- Cross-run formatting preservation (smart-strip's uniform-vs-mixed rule)
- Character-indent semantics — `firstLineChars` round-trips as `"Nchar"`, fixed twips as `"Npt"`
- Dominant-run selection in `fromParagraph` skips numbering-prefix-only runs
- Leading-prefix strip accumulates across consecutive `<w:t>` runs — Word splits hand-edited paragraphs mid-prefix (e.g. `<w:t>五</w:t><w:t>、Title</w:t>`), and a per-run regex test silently misses these
- numId migration on template import uses fresh IDs to avoid collision
- The original file is never modified; `apply_styles` always writes a fresh copy and validates before keeping it

When changing any of these, verify against `test/fixtures/` and inspect the output zip.

## Commit style

Title: what changed. Body: why, in 1–3 short paragraphs max — enough that the next maintainer wouldn't accidentally revert the decision. The diff already shows what; don't restate it. No "Net: X → Y lines" stats, no per-bullet narration of each change, no recap of content already in CLAUDE.md / SKILL.md. Don't include `co-authored-by` tags.
