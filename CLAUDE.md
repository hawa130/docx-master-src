# Working in this repo

This repo builds the `docx-normalize` Claude skill — scripts that classify paragraph roles in a Word document and inject named styles. `SKILL.md` is the agent-facing contract; this file is for working *on* the project.

## Layout

```
src/core/        shared parsing / style / fingerprint / template-import logic
src/tools/       one CLI entry per tool (each maps to a dist script)
references/     agent-facing reference docs (progressive disclosure from SKILL.md)
SKILL.md        the skill spec the agent reads at runtime
test/fixtures/  sample .docx files for manual testing
dist/           build output; dist/docx-normalize/ is the staged skill bundle
```

## Commands

| Task | Command |
|---|---|
| Build TypeScript → dist/ | `bun run build` |
| Build + stage skill bundle + zip | `bun run build:skill` |
| Watch | `bun run build:watch` |

No automated tests — run scripts against `test/fixtures/*.docx` manually after changes. After edits to `src/` or `SKILL.md`, always `bun run build:skill` and verify `dist/docx-normalize/` reflects the change before claiming done.

## Design principles

### Tools expose visible facts; the agent makes role judgments

Tools should expose what a human reader of the document can see — font, size, weight, color, alignment, indent, numbering markers, text content. The agent classifies roles from these facts.

**Don't pre-digest hidden metadata or natural language into role hints in default output.** Two times this was tried — a regex parser for Chinese typography ("小四宋体" → fields), and an `[LN]` outline-level hint on the fingerprint summary — both backfired the same way: the agent accepts the convenient pre-classification as ground truth and bypasses its own judgment, the same judgment that would have caught upstream errors. Both removed.

Hidden metadata (`outlineLevel`, `pStyle`, `numId`, abstractNum) can still be exposed by tools — but as **raw data on demand**, when the agent calls `inspect_range` / `inspect_style_def`. Never pre-packaged into a default summary.

**Litmus test for any new "helpful hint":** could a human reader derive this without opening the XML? Yes → expose it. No → only on demand, never as a role hint.

### SKILL.md is the agent's runtime context — keep it lean

Every line of SKILL.md is loaded into the agent's context on every invocation. Encode only what the agent can't derive: technical invariants the script depends on, workflow anchors, tool references. Cut restated points, disambiguation tables for things an LLM infers from context, and examples of tool output (the tool prints them). Detailed schema lives in `references/`, loaded on demand.

When feedback comes in, ask first: *can the agent's own judgment cover this?* If yes, don't add. Encode only what's repeatedly missed across multiple agents or mechanically non-obvious.

### Tools follow the on-demand pattern

Each tool has one focused job. Don't add always-on enrichment to default outputs — prefer a new focused script (e.g. `inspect_neighbors` is separate rather than putting neighbor info on every paragraph in `inspect_range`).

### Mechanical correctness is the script's job

LLMs are bad at byte-level work. The scripts must guarantee:

- XML-namespace-correct mutation of `styles.xml` / `numbering.xml` / `document.xml`
- Cross-run formatting preservation (smart-strip's uniform-vs-mixed rule)
- Character-indent semantics — `firstLineChars` round-trips as `"Nchar"`, fixed twips as `"Npt"`
- Dominant-run selection in `fromParagraph` skips numbering-prefix-only runs
- numId migration on template import uses fresh IDs to avoid collision
- The original file is never modified; `apply_styles` always writes a fresh copy and validates before keeping it

When changing any of these, verify against `test/fixtures/` and inspect the output zip.

## Commit style

Detailed messages explaining *why*. The history is the design record — when a feature was removed (the natural-language parser, the always-on neighbor fields, the outline hint), the commit message documents the rationale so the same idea isn't reinvented. Don't include `co-authored-by` tags.
