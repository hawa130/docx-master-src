<!-- prettier-ignore -->
**English** · [简体中文](README.zh-CN.md)

# docx-master

The Word document automation agents have been waiting for. One skill, 18 tools, and a sparse-by-design config language that mutates `.docx` OOXML directly — restyle, renumber, edit, audit — without the fragile round-trip through Markdown or HTML.

> **Quick start:** `cp -r dist/docx-master ~/.claude/skills/` (or grab the [latest release](https://github.com/hawa130/docx-master/releases) and drop it into your harness of choice).

## Why docx-master?

Every LLM that touches a `.docx` does one of two things, and both are wrong:

- **Convert to Markdown, regenerate the document.** Loses styles, numbering, tracked changes, fields, sections, theme — everything that makes a Word document *a Word document*.
- **Hand the user a `python-docx` snippet.** Direct paragraph formatting on every line; no `styles.xml` discipline; auto-numbering replaced by typed `"1."` prefixes that desync the moment someone inserts a section.

docx-master takes a different stance: **the document is OOXML; mutate it in place, sparsely, through styles + numbering + sections** — the same way Microsoft Word, WebAIM, and ECMA-376 say it should be done.

docx-master adds:

- **Two CLIs, one mental model.** `apply` is the unified writer (install styles + numbering + theme + template, restyle by pattern or fingerprint, insert content via edits — all one config). `audit` is the read-only conformance check.
- **18 inspect / find / migrate / validate tools.** Each surfaces one slice of the document so the agent can classify before it writes. Default outputs are scannable; deep info is one tool call away.
- **A sparse config language.** Declare only what's wrong, missing, or what the user asked to change. Untouched styles, numbering, paragraphs, and theme stay as they are. No "regenerate the whole stylesheet" footguns.
- **Real OOXML safety.** Schema-aware validation runs on every write. The original file is never modified. Tracked-change / field / SDT regions are detected and refused for ad-hoc rewriting. Fresh `numId`s on template import to prevent collision. Cross-namespace XML correctness for `styles.xml`, `numbering.xml`, `document.xml`.
- **Caption + cross-reference primitives.** Figures, tables, equations, theorems — declared once in a `captions` table, referenced via `InlineRef` nodes in body text. Word emits the SEQ + STYLEREF fields and REF links automatically. Counters never go stale.
- **Locale-aware defaults.** Chinese font-size names (小四, 五号, …), 2-character first-line indent for CJK prose, autoSpace-respecting whitespace rules at CJK ↔ Latin boundaries, GB/T 15834 curly quotes. Out of the box.

## What's Included

### The Skill: docx-master

A focused Word-automation skill with 10 on-demand reference files ([view skill](skill/SKILL.md)):

| Reference | Covers |
|-----------|--------|
| [standardize.md](skill/references/standardize.md) | Whole-doc reshape: `styles[]`, `numbering`, `pattern_rules`, `bulk_rules`, `assignments`, `exclude` |
| [edit.md](skill/references/edit.md) | Surgical edits: locators, ops (replace / insert / delete / set-run), MDF, tracked changes |
| [config-schema.md](skill/references/config-schema.md) | Full `apply` config field reference |
| [captions.md](skill/references/captions.md) | SEQ-based captions, chapter-prefixed numbering, REF cross-references |
| [cross-references.md](skill/references/cross-references.md) | `InlineRef` schema for figure / table / equation / section cites |
| [numbering-formats.md](skill/references/numbering-formats.md) | Multi-level numbering shapes — decimal, parenthesized, CJK 序号, bullet |
| [tables.md](skill/references/tables.md) | Table block schema for `edits[]` insertion |
| [equations.md](skill/references/equations.md) | LaTeX → OMML inline + display math, numbered equation layout |
| [chinese-font-sizes.md](skill/references/chinese-font-sizes.md) | 小四 / 五号 / 三号 / … → half-points |
| [audit.md](skill/references/audit.md) | Read-only conformance workflow + scanning axes |

### 18 Tools

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
| `migrate_numbering` | Surveys legacy `numId`s for consolidation into one multi-level scheme |
| `find_paragraphs` | Cross-document regex search to validate `pattern_rules` coverage |
| `find_text` | Character-level locator — paragraph index, run index, char offset, context |
| `import_template` | Pull styles + numbering from a reference docx |
| `restyle` | Standalone restyle entry (legacy; `apply` is the unified writer) |
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

- **Tools expose visible facts; the agent makes role judgments.** No pre-classification baked into default output; the agent classifies from what a human reader would see, then optionally consults hidden metadata on demand.
- **Mechanical correctness is the script's job.** Paragraph walks, namespace-correct XML mutation, cross-run formatting preservation, `numId` collision avoidance, blocker detection, validation — never the agent's responsibility, never bent under refactoring pressure.
- **Verification checks intent, not interpretation.** A check that grades the system against its own interpretation of the input is a tautology. Real verification is human-readable side-by-side or output re-parsed against an independent invariant.
- **The original is never modified.** Every write produces a fresh file + validates before keeping it. Validation failure discards and surfaces — no silent retry.

See [CLAUDE.md](CLAUDE.md) for the full set.

## Installation

Grab the zip for your harness from the [latest release](https://github.com/hawa130/docx-master/releases) and unpack it into the directory the harness loads skills from. Every release publishes:

| Asset | Unpacks to | For |
|---|---|---|
| `docx-master.zip` | `docx-master/` | Universal — any harness that loads Markdown skills |
| `docx-master-claude-code.zip` | `.claude/skills/docx-master/` | Claude Code |
| `docx-master-cursor.zip` | `.cursor/skills/docx-master/` | Cursor |
| `docx-master-codex.zip` | `.agents/skills/docx-master/` | Codex CLI |
| `docx-master-gemini.zip` | `.gemini/skills/docx-master/` | Gemini CLI |
| `docx-master-opencode.zip` | `.opencode/skills/docx-master/` | OpenCode |
| `docx-master-github.zip` | `.github/skills/docx-master/` | GitHub Copilot |

Pick one of the per-harness zips and unzip it at the right scope:

```bash
# Claude Code, user-wide
cd ~ && unzip ~/Downloads/docx-master-claude-code.zip

# Claude Code, project-local
cd your-project && unzip ~/Downloads/docx-master-claude-code.zip

# Or the universal bundle, dropped into whichever skills directory applies
unzip ~/Downloads/docx-master.zip -d ~/.claude/skills/
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
| `skill/tools/` | TS source for the 18 CLIs |
| `lib/` | Non-tool TS modules (xml / parse / config / apply / edit / shared) |
| `test/fixtures/` | Sample `.docx` files for manual verification |
| `build-skill.ts` | Stages `dist/docx-master/` + zip + per-provider fan-out |
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

There are no automated tests yet — verify changes against `test/fixtures/*.docx` and inspect the produced bundle.

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
