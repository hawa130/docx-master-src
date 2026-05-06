# Working in this repo

This repo builds the `docx-normalize` Claude skill — scripts that classify paragraph roles in a Word document and inject named styles. `SKILL.md` is the agent-facing contract; this file is for working *on* the project.

**Keep this file in sync.** Tool names, build commands, file paths, and the lessons below are referenced concretely. When any of them changes, update here in the same commit — stale references mislead future maintainers and the next agent reviewing the design.

## Layout

```
src/core/        shared parsing / style / fingerprint / template-import logic
src/tools/       one CLI entry per tool (each maps to a dist script)
references/     agent-facing reference docs (progressive disclosure from SKILL.md)
SKILL.md        the skill spec the agent reads at runtime
test/fixtures/  sample .docx files for manual testing
dist/           build output; dist/docx-normalize/ is the staged skill bundle
build-skill.ts  packages the staged dir into the .skill zip
```

## Commands

| Task | Command |
|---|---|
| Build TypeScript → dist/ | `bun run build` |
| Build + stage skill bundle + zip | `bun run build:skill` |
| Watch | `bun run build:watch` |

No automated tests — run scripts against `test/fixtures/*.docx` manually after changes. After edits to `src/` or `SKILL.md`, always `bun run build:skill` and verify `dist/docx-normalize/` reflects the change before claiming done.

## Design principles

### Tools expose visible facts; agents make role judgments

Tools should expose what a human reader of the artifact can see — in this skill: font, size, weight, color, alignment, indent, numbering markers, text content. The agent classifies roles from those visible facts.

**Don't pre-digest hidden metadata or natural language into role hints in default output.** When tooling presents a "convenient pre-classification" (a parsed translation, a metadata-derived label), the LLM accepts it as ground truth and bypasses its own judgment — the same judgment that would have caught upstream errors. This anti-pattern showed up twice in this project (a regex parser for Chinese typography, an `[LN]` outline-level hint on the fingerprint summary); both were removed for the same reason.

Hidden metadata can still be exposed — but as **raw data on demand**, when the agent explicitly calls `inspect_range` / `inspect_style_def`. On-demand exposure means the agent has already started reasoning and is using the metadata as one input among many. Pre-packaging it into a default summary means the tool is doing the reasoning for the agent.

**Litmus test for any new "helpful hint":** could a human reader derive this from the artifact's normal rendering, without opening the underlying file format? Yes → expose it. No → only on demand, never as a role hint in default output.

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
- numId migration on template import uses fresh IDs to avoid collision
- The original file is never modified; `apply_styles` always writes a fresh copy and validates before keeping it

When changing any of these, verify against `test/fixtures/` and inspect the output zip.

## Commit style

Title: what changed. Body: why, in 1–3 short paragraphs max — enough that the next maintainer wouldn't accidentally revert the decision. The diff already shows what; don't restate it. No "Net: X → Y lines" stats, no per-bullet narration of each change, no recap of content already in CLAUDE.md / SKILL.md. Don't include `co-authored-by` tags.
