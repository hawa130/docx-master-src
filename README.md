<!-- prettier-ignore -->
**English** · [简体中文](README.zh-CN.md)

# docx-master

The Word document automation agents have been waiting for. One skill, 15 tools, and a sparse-by-design config language that mutates `.docx` OOXML directly — restyle, renumber, edit, audit — without the fragile round-trip through Markdown or HTML.

> **Quick start:** `npx skills add hawa130/docx-master` (or grab the [latest release](https://github.com/hawa130/docx-master/releases) zip and drop it into your harness's skills directory).

## Peer Word skills

Most Word skills are LLM-facing documentation for an underlying library — python-docx, docx-js, OpenXML SDK. They hand the agent a set of primitives; what a good Word document looks like is left for the agent to figure out.

docx-master comes with its own convention for writing Word documents:

- Every paragraph binds to a named style (Heading, Body Text, List, etc.); formatting follows the style.
- Chapters, figures, tables, and equations all use Word's built-in auto-numbering.
- References like "see Figure 1.2" stay live; every number updates together when sections move.
- CJK and Latin runs in the same paragraph keep their own font sizes.

Each convention comes with its own tooling: `pattern_rules` / `bulk_rules` apply rules across the document in one pass, `migrate_captions` detects manually-numbered captions to convert, `audit` flags violations, and 12 inspect / find tools let the agent survey the document before mutating.

Side-by-side with peer skills:

- [**anthropics/docx**](https://skills.sh/anthropics/skills/docx) — Anthropic's official skill. Unpack + edit XML + docx-js for creation, general-purpose Word work.
- [**qodex-ai/word-document-processor**](https://skills.sh/qodex-ai/ai-agent-skills/word-document-processor) — A toolbox of pandoc + docx-js + python-docx + raw XML, focused on redlining workflows.
- [**minimax-ai/minimax-docx**](https://skills.sh/minimax-ai/skills/minimax-docx) — C# + OpenXML SDK (.NET), focused on structured editing.
- [**claude-office-skills/docx-manipulation**](https://skills.sh/claude-office-skills/skills/docx-manipulation) — python-docx wrapper, focused on template placeholder replacement.

Capability comparison:

| Scenario | anthropics | qodex-ai | minimax-ai | claude-office | docx-master |
|---|:---:|:---:|:---:|:---:|:---:|
| Fill an existing blank template with content | ~ | ~ | ✓ | ✓ | ✓ |
| Create a new Word document from scratch | ~ | ~ | ✓ | ~ | ✓ |
| Auto-numbered chapter / section headings (no typed "1.1.1") | — | — | ~ | — | ✓ |
| CJK and Latin fonts don't override each other in mixed paragraphs | — | — | ~ | — | ✓ |
| Figures, tables, equations numbered as `chapter.n` | — | — | — | — | ✓ |
| "See Figure 1.1" as a live field that updates on reorder | — | — | ~ | — | ✓ |
| LaTeX equations as centered display blocks + number | — | — | — | — | ✓ |
| Reviewer feedback flows back as tracked changes | ✓ | ✓ | ~ | — | ✓ |
| Convert typed "Figure 2.1" from a legacy draft to live fields | — | — | — | — | ✓ |
| Fill form blanks without breaking the underline | — | — | — | — | ✓ |
| Edit a specific cell in a table | — | — | ~ | ~ | ✓ |
| Page headers / footers with page numbers | ~ | — | ~ | ~ | ✓ |
| Different page layout per section (portrait / landscape, margins) | ~ | — | ~ | ~ | ✓ |
| Format conformance check | — | — | — | — | ✓ |
| Borrow styles from another document | — | — | ✓ | — | ✓ |

> ✓ built-in helper, declare-and-use · ~ doable but the agent assembles pieces · — no specific support

anthropics/docx covers everyday creation and one-off edits. qodex-ai fits more naturally when redlining workflows dominate, .NET projects tend toward minimax-ai, and claude-office-skills is the shortest route for template placeholder fills. docx-master sits in the long-form structural reshape space, especially for CJK documents.

The full sub-command surface, Block types, and reference docs are listed in the "What's Included" section below.

## What's Included

### The Skill: docx-master

A focused Word-automation skill with 11 on-demand reference files ([view skill](skill/SKILL.md)):

| Reference | Covers |
|-----------|--------|
| [standardize.md](skill/references/standardize.md) | Whole-doc reshape: `styles[]`, `numbering`, `pattern_rules`, `bulk_rules`, `assignments`, `exclude` |
| [edit.md](skill/references/edit.md) | Surgical edits: locators, ops (replace / insert / delete / set-run), MDF, tracked changes |
| [config-schema.md](skill/references/config-schema.md) | Full `apply` config field reference |
| [captions.md](skill/references/captions.md) | SEQ-based captions, chapter-prefixed numbering, REF cross-references |
| [cross-references.md](skill/references/cross-references.md) | `InlineRef` schema for figure / table / equation / section cites |
| [numbering-formats.md](skill/references/numbering-formats.md) | Multi-level numbering shapes — decimal, parenthesized, CJK 序号, bullet |
| [tables.md](skill/references/tables.md) | Table block schema for `edits[]` insertion |
| [header-footer.md](skill/references/header-footer.md) | Page headers / footers — fixed text, page numbers, current chapter title |
| [equations.md](skill/references/equations.md) | LaTeX → OMML inline + display math, numbered equation layout |
| [chinese-font-sizes.md](skill/references/chinese-font-sizes.md) | 小四 / 五号 / 三号 / … → pt |
| [audit.md](skill/references/audit.md) | Read-only conformance workflow + scanning axes |

### 15 Tools

All tools invoke via `node scripts/<name>.js <args>` and write to stdout. The skill's prompt routes the agent to the right one.

| Tool | When |
|------|------|
| `overview` | First call on any task. Metadata, page setup, theme, style defs, numbering schemes, fingerprint statistics, skeleton |
| `inspect_range` | Full text + computed styles for a paragraph range |
| `inspect_runs` | Per-run rPr dump for mixed-formatting or form-fill paragraphs |
| `inspect_neighbors` | What surrounds a paragraph — caption / first-after-heading classification |
| `inspect_style` | What role a fingerprint plays across the document |
| `inspect_style_def` | Pre-defined styles in `styles.xml` and `basedOn` chain |
| `inspect_section` | Page-setup differences between sections |
| `inspect_table` | Top-level tables with cell text + paragraph-index spans |
| `inspect_blockers` | Paragraphs the edit phase will refuse (tracked changes, fields, SDT) |
| `inspect_caption` | SEQ-based captions — list identifiers, per-occurrence dump |
| `migrate_captions` | Read-only detector for manually-numbered caption paragraphs |
| `find_paragraphs` | Cross-document regex search to validate `pattern_rules` coverage |
| `find_text` | Character-level locator — paragraph index, run index, char offset, context |
| `validate` | Schema-aware OOXML check on any `.docx` |
| `apply` | The unified writer. `--dry-run` for iteration, no flag to write |

### `apply` pipeline

```
install styles + numbering + theme + template
  → run edits (referencing just-installed styleIds)
  → re-fingerprint
  → run rules (pattern_rules / bulk_rules / assignments / exclude —
    match BOTH pre-existing chrome AND agent-inserted content uniformly)
  → validate, write
```

Sparse by design — only declared blocks apply.

## Design principles

Tools expose visible facts; classification and judgment are the agent's. Default output carries no pre-classification — the agent sees what a human reader would see, then consults hidden metadata on demand.

Mechanical correctness sits with the scripts. Paragraph walks, namespace-correct XML mutation, cross-run formatting preservation, `numId` collision avoidance, blocker detection, validation: all kept out of the agent's hands, and not loosened under refactoring pressure either.

What gets verified is intent, not the system's own reading of its input. A check that grades the system against its own interpretation is a tautology. Real verification is a human-readable side-by-side, or the output reparsed against an independent invariant.

The original file is never modified. Every write produces a fresh file; if validation fails, that file is discarded and surfaced. No silent retry.

## Installation

Easiest: the [`skills` CLI](https://skills.sh) routes the bundle into the right directory for your harness.

```bash
# Project-local (committed to your repo)
npx skills add hawa130/docx-master

# User-wide
npx skills add hawa130/docx-master -g

# Target a specific harness
npx skills add hawa130/docx-master -a claude-code
```

Or grab `docx-master.zip` from the [latest release](https://github.com/hawa130/docx-master/releases) and unzip into whichever skills directory your harness loads from:

```bash
# Claude Code, user-wide
unzip ~/Downloads/docx-master.zip -d ~/.claude/skills/

# Cursor, project-local
unzip ~/Downloads/docx-master.zip -d your-project/.cursor/skills/
```

Harness-specific gotchas:

- **Cursor** — switch to the Nightly channel (Settings → Beta) and enable Agent Skills (Settings → Rules). [Docs](https://cursor.com/docs/context/skills).
- **Gemini CLI** — `npm i -g @google/gemini-cli@preview`, `/settings` → enable "Skills", verify via `/skills list`. [Docs](https://geminicli.com/docs/cli/skills/).

### Runtime requirements

The skill ships its own scripts as bundled Node CJS. Any harness that can run `node` (Node 18+) on the user's machine can run docx-master. No Python, no pip install, no system dependencies — the OOXML schemas and xmllint-wasm validator are vendored into the bundle.

## Usage

The skill auto-triggers from the prompt. Some everyday shapes:

> Fill in this empty proposal template — headings auto-numbered, body in
> 11pt, figure and table cites as live fields. Body text is mixed Chinese
> and English; Songti for CJK, Times New Roman for Latin runs.

> Audit this draft. What fights the style system, what headings aren't
> tagged, which captions won't renumber on reorder.

> Convert every "Figure 2.1" / "Table 1.3" in this manuscript from typed
> numbers to live fields, and rewire the body references.

> Insert these three clauses after paragraph 42 with tracked changes on.

> This form has labeled blanks like "Name: ____". Fill the blank after
> "Project Title" with "Q3 Marketing Plan" without disturbing the label
> or the underline run.

> Render these LaTeX equations as numbered display blocks, with cross-refs
> in the prose pointing at the equation numbers.

> Lift the heading and caption styles from reference.docx, apply them here
> without touching my numbering.

Every write produces a fresh, schema-validated copy — the input file is never modified. Numbering, captions, and cross-references land as live fields, so reordering a section keeps everything consistent on the next "Update Fields".

## Repo layout

| Path | What it is |
|---|---|
| `skill/SKILL.md` | Agent-facing contract (top-level routing) |
| `skill/references/` | On-demand detail; loaded only when relevant |
| `skill/tools/` | TS source for the 15 CLIs |
| `lib/` | Non-tool TS modules (xml / parse / config / apply / edit / shared) |
| `tests/` | Math regression corpus + runner (`bun run test:math`) |
| `build-skill.ts` | Stages `dist/plugin/skills/docx-master/` (the publish-repo image) + zip + rendered `.claude-plugin/` manifest |
| `CLAUDE.md` | Working-on-the-project guide for contributors / future agents |

## Building from source

```bash
bun install
bun run build:skill   # → everything under dist/
bun run typecheck
bun run lint
bun run fmt:check
```

`bun run build:skill` produces, under `dist/`:

- `docx-master/` — the staged skill bundle (`SKILL.md` + `references/` + bundled `scripts/`)
- `docx-master.zip` — universal release artifact
- `<provider>/<configDir>/skills/docx-master/` — one per supported harness (`claude-code/.claude/`, `cursor/.cursor/`, `codex/.agents/`, `gemini/.gemini/`, `opencode/.opencode/`, `github/.github/`)
- `plugin/skills/docx-master/` — mirror that `.claude-plugin/marketplace.json` references for local Claude Code marketplace testing

All of `dist/` is gitignored — regenerated on each build. Per-harness zips for releases are produced by the [release workflow](.github/workflows/release.yml).

Math conversion has a regression corpus (`bun run test:math`). For other features verify changes against local .docx samples and inspect the produced bundle.

## Supported harnesses

- [Claude Code](https://claude.com/claude-code) — primary target, plugin-installable
- [Cursor](https://cursor.com)
- [Codex CLI](https://github.com/openai/codex)
- [Gemini CLI](https://github.com/google-gemini/gemini-cli)
- [OpenCode](https://opencode.ai)
- [GitHub Copilot](https://github.com/features/copilot)

Other harnesses that load Markdown skills with YAML frontmatter should work with the universal bundle — drop the `docx-master/` directory wherever the harness expects skills, no transformation needed.

## Contributing

See [CLAUDE.md](CLAUDE.md) for repo conventions, the design-principle ladder the skill content holds to, and the cross-command invariants the engine maintains. Pull requests should keep the agent-facing surface (`SKILL.md` + `references/`) within its existing token budget — every line loads on every invocation.

## License

Apache 2.0. See [LICENSE](LICENSE).
