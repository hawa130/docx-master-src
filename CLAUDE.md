# Working in this repo

This repo builds the `docx-normalize` Claude skill — scripts that read a Word document, classify paragraph roles, and inject named styles. The agent-facing contract lives in `SKILL.md`; this file is for working *on* the project.

## Layout

```
src/core/        shared parsing / style / fingerprint / template-import logic
src/tools/      CLI entry per tool (one file per dist script)
references/     agent-facing reference docs (progressive disclosure from SKILL.md)
SKILL.md        the agent-facing skill spec — what an agent reads at runtime
test/fixtures/  sample .docx files for manual testing
dist/           build output; dist/docx-normalize/ is the staged skill bundle
build-skill.ts  packages dist/docx-normalize/ + writes the .skill zip
```

## Commands

| Task | Command |
|---|---|
| Build (TypeScript → dist/) | `bun run build` |
| Build + stage skill bundle + zip | `bun run build:skill` |
| Watch | `bun run build:watch` |

There are no automated tests. Run scripts against `test/fixtures/*.docx` manually after changes.

## Editing rules

### SKILL.md is the agent's runtime context — keep it lean

Every line of SKILL.md is loaded into the agent's context on every invocation. Only put things there that meet at least one of:

- A technical fact the agent can't derive (e.g. `firstLineChars` vs `firstLine` semantics; smart-strip's uniform-vs-mixed rule; that `font` extracts as Arial because of theme defaults).
- A workflow phase or tool reference.
- An invariant that would silently produce wrong output if violated.

Things that do **not** belong in SKILL.md:

- Re-statements of points already made (cross-section repetition).
- "Don't reason X" anti-pattern paragraphs that just restate the positive principle.
- Disambiguation tables for phrasings an LLM agent infers naturally from context.
- Sales-copy / encouragement ("combine layers freely!").
- Examples of tool *output* (the tool prints them itself).
- Edge cases the tooling already handles automatically.

When external feedback comes in, ask first: *can the agent's own judgment cover this?* If yes, don't add to SKILL.md. Only encode things that are mechanically non-obvious or repeatedly missed across multiple agents.

Detailed schema docs go to `references/` (loaded on demand, not always-on). The agent reads them when needed; SKILL.md just points.

### Tools follow the "on-demand" pattern

Each tool has one focused job and is invoked when the agent specifically needs it. Don't add always-on enrichment fields to the default output (e.g. don't put neighbor info on every paragraph in `inspect_range` — that's what `inspect_neighbors` is for). The same applies for new features: prefer a new focused script over bolting onto an existing one.

### Tools expose visible facts; agents make role judgments

The agent classifies paragraphs by the same evidence a human reader has — font, size, weight, color, alignment, indent, numbering presence, position relative to images/tables, text content. Tools should expose these visible facts cleanly.

Tools should NOT pre-digest hidden metadata into role classifications. Two patterns this killed in the past:

- A regex parser that translated "小四宋体" into structured fields (`requirements-parser.ts`, removed). Looked like it was helping; actually let the agent stop reading the user's text carefully and silently mistranslated negation / synonyms.
- An `[LN]` hint on the fingerprint summary that surfaced the source's `outlineLevel` metadata as a heading-role label. Looked like a strong signal; actually let the agent skip visual classification reasoning and propagated source-metadata errors (POI-generated docs, conversion artifacts, residual outlineLvls from abandoned Heading-style usage) straight into the output.

The pattern in both cases: a "convenient" pre-classification that LLMs gladly accept as ground truth, bypassing their own judgment. The same judgment that catches errors when present.

Hidden metadata (`outlineLevel`, `pStyle`, `numId`, abstractNum definitions) can still be exposed by tools — but as **raw data on demand**, surfaced when the agent asks (`inspect_range` shows outlineLevel, `inspect_style_def` walks the inheritance chain). On-demand exposure means the agent has already started reasoning and is using the metadata as one input. Pre-packaging it into a default summary means the tool is doing the reasoning for the agent.

Litmus test for new features: **could a human reader of the document, without opening the XML, derive this information?** If yes, exposing it is fine. If no, it's hidden metadata; only expose on demand, never as a role hint.

### The script never does language judgment

The agent (LLM) handles all semantic work: classifying paragraph roles, translating natural-language requirements into structured fields, deciding whether two fingerprints should merge. The scripts present facts: computed styles, element positions, fingerprints, what's adjacent. **Don't add Chinese typography parsing back** — there used to be a `requirements-parser.ts` that translated "小四宋体" into config fields; it was removed because regex parsing of natural language gives false confidence (it can't tell "不要加粗" from "加粗"). The `requirements` field is annotation-only.

### Mechanical correctness is the script's job

Things the scripts must get right (because LLMs are bad at byte-level work):

- XML-namespace-correct mutation of styles.xml / numbering.xml / document.xml
- Cross-run formatting preservation (smart-strip uniform-vs-mixed rule in `apply-styles.ts`)
- Character-based indent semantics (`firstLineChars` round-trips as `"Nchar"`, fixed twips as `"Npt"`)
- Dominant-run selection in `fromParagraph` (skip numbering-prefix-only runs)
- numId migration on template import (fresh IDs to avoid collision)

When changing any of these, verify against `test/fixtures/` and check the output zip. The original file must never be modified — `apply_styles` always writes a fresh copy.

### Build before claiming a change works

The skill is shipped from `dist/docx-normalize/`, built by `tsdown` + `build-skill.ts`. After edits to `src/` or `SKILL.md`, run `bun run build:skill` and check `dist/docx-normalize/` reflects the change. The Node scripts in `dist/` are what an agent actually runs.

## Commit style

Detailed messages explaining *why*, not just *what*. The history is part of the design record — when a feature was removed (the natural-language parser, the always-on neighbor fields), the commit message captures the rationale so the same idea isn't reinvented.

Don't include "co-authored-by" tags.
