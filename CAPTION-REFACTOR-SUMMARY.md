# Caption refactor — what landed

Branch `caption-refactor` (14 commits) implements the design in
[`CAPTION-REFACTOR-SPEC.md`](./CAPTION-REFACTOR-SPEC.md). This file is
the post-implementation reference; consult the spec for the *why* in
depth.

## 1. TL;DR

Caption-class enumerators (equation numbers, figure / table / theorem
captions) now render as Word-native `SEQ` + `STYLEREF` fields wrapped
in bookmarks, instead of riding `numPr` / `numbering.xml`. Outline
numbering (H1–H6) and lists are unchanged. A new top-level `captions`
table in the apply config carries the per-identifier shape (prefix /
chapter-prefix / format / styleId / subCounter); two new block types
(`CaptionBlock`, `CaptionCounterReset`) and one new edit op
(`edit-caption`) sit on the edit-config surface.

## 2. Motivation recap

Quoting the spec: *"`numPr` / `numbering.xml`: outline numbering (H1–H6),
bullet / ordered lists. … `SEQ` fields: caption-class enumerators."*
The previous toolchain repurposed `numPr` for captions because the
multi-level machinery was already there. Side-effects: chapter prefix
forced outline coupling (H1 and equation co-tenant a `numId`), each
caption type needed its own `abstractNum`, and Word's
References → Cross-reference dialog couldn't surface our captions (it
expects SEQ). Best-practice consensus across the Word community (cited
in spec §1) is SEQ + STYLEREF + bookmark + REF `\h`. This refactor
aligns.

## 3. Architecture

Apply pipeline ordering (spec §9.4, now implemented):

```
parse apply config
  └─ caption-resolver.ts ─────────► ResolvedCaptionConfig per identifier
                                     (styleId→styleName, outlineLvl+1,
                                      defaults applied)

standardize pass
  ├─ paragraph-style application
  └─ standardize-captions.ts ────► re-emit existing SEQ-bearing caption
                                     paragraphs in place (preserves
                                     bookmark id+name, paragraph identity,
                                     body text after primary bookmarkEnd)

edit ops
  └─ fragment-emit.ts dispatch
       ├─ "equation"  ───► equation-emit.ts
       │                     └─ if captionId: caption-emit.emitNumberedEquation
       │                        else: existing centered-paragraph path
       ├─ "caption"   ───► caption-emit.emitCaptionBlock
       └─ "caption-counter-reset" ─► caption-emit.emitCaptionReset

outline numbering simulator         (existing — produces lvlText renderings)
caption-counter.ts simulator        (new — consumes outline values for
                                      STYLEREF resolution; produces
                                      caption placeholder text per
                                      identifier; handles subGroup +
                                      CaptionCounterReset)

placeholder backfill                (REF results for caption-class targets
                                      get fullCaptionText; SEQ / STYLEREF
                                      result runs get fieldValues from sim)
```

Key modules:

| Path | Role |
|---|---|
| `lib/edit/fields/complex-field.ts` | Shared 5-run fldChar skeleton + `applyFieldFormat` (rPr replication + MERGEFORMAT) |
| `lib/edit/fields/ref-field.ts` | REF emitter (relocated + thinned from old `field-ref.ts`) |
| `lib/edit/fields/seq-field.ts` | SEQ emitter; `\* <FORMAT>` `\s` `\r` `\c` `\h` `\* MERGEFORMAT`; `\c`+`\r` mutex enforced |
| `lib/edit/fields/styleref-field.ts` | STYLEREF emitter; always quotes styleName |
| `lib/edit/fields/field-parse.ts` | fldChar group → structured `ParsedRun` (text / field with details) |
| `lib/edit/caption-emit.ts` | `emit{Numbered,Unnumbered}Equation`, `emitCaptionBlock`, `emitCaptionReset` |
| `lib/edit/caption-counter.ts` | Per-identifier counter sim: parent + sub + openSubGroup; chapter restart on last-chapter-style boundary |
| `lib/edit/edit-caption-op.ts` | `edit-caption` op: locate paragraph → find primary bookmarkEnd → swap body |
| `lib/apply/standardize-captions.ts` | Re-emit pre-body run sequence on already-SEQ caption paragraphs |
| `lib/parse/caption-resolver.ts` | `captions[*]` × `styles.xml` → `ResolvedCaptionConfig` |
| `lib/edit/bookmark.ts` | + `allocateRangeBookmark` / `bindRangeBookmark` for inline bookmarkStart/End emission |

## 4. Schema changes

Before:

```jsonc
// apply config
"numbering": [
  { "levels": [{ "styleId": "EquationNumber", "lvlText": "(%1.%2)",
                 "numFmt": "decimal", ... }] },
  { "levels": [{ "styleId": "FigureCaption",  "lvlText": "图 %1.%2", ... }] }
]
// edit block
{ "type": "equation", "latex": "E=mc^2",
  "numbering": { "numId": 7, "level": 0 }, "displayStyle": true }
{ "type": "paragraph", "styleId": "FigureCaption", "text": "..." }
```

After:

```jsonc
// apply config
"captions": {
  "Equation": { "prefix": "(", "suffix": ")", "format": "arabic",
                "chapterPrefix": ["Heading1New"], "chapterSeparator": ".",
                "styleId": "EquationNumber" },
  "Figure":   { "prefix": "图 ", "format": "arabic",
                "chapterPrefix": ["Heading1New"], "chapterSeparator": ".",
                "bodySeparator": "  ", "styleId": "FigureCaption" }
}
// edit blocks
{ "type": "equation", "latex": "E=mc^2",
  "styleId": "Equation", "captionId": "Equation", "anchor": "eq-mass-energy" }
{ "type": "caption", "captionId": "Figure", "text": "系统架构示意图",
  "anchor": "fig-arch" }
```

Cross-field invariants enforced at config validation (spec §3.6):
exactly one of `latex` / `omml` on `EquationBlock`; `subGroup` and
`anchor` each require `captionId`; `subGroup` requires
`captions[captionId].subCounter`; `display:"full"` on EquationBlock
throws; a styleId cannot appear in both `numbering[].levels[].styleId`
and `captions[<id>].styleId` (caught at apply start).

## 5. Files added / changed / removed

### Engine (`lib/`)

| File | Change |
|---|---|
| `lib/edit/fields/complex-field.ts` | added |
| `lib/edit/fields/ref-field.ts` | added (relocated from `field-ref.ts`) |
| `lib/edit/fields/seq-field.ts` | added |
| `lib/edit/fields/styleref-field.ts` | added |
| `lib/edit/fields/field-parse.ts` | added |
| `lib/edit/caption-emit.ts` | added |
| `lib/edit/caption-counter.ts` | added |
| `lib/edit/edit-caption-op.ts` | added |
| `lib/edit/field-ref.ts` | **removed** (split into `fields/` modules) |
| `lib/edit/edit-engine.ts` | + caption-target prewalk, ApplyDeps threading, edit-caption dispatch |
| `lib/edit/fragment-emit.ts` | + `caption` / `caption-counter-reset` dispatch; equation routes via captionId |
| `lib/edit/math/equation-emit.ts` | + `omml` escape hatch, captioned routing via caption-emit |
| `lib/edit/bookmark.ts` | + `allocateRangeBookmark` / `bindRangeBookmark` |

### Apply (`lib/apply/`, `lib/parse/`, `lib/config/`)

| File | Change |
|---|---|
| `lib/apply/standardize-captions.ts` | added |
| `lib/apply/apply-styles.ts` | + caption-resolver wiring, simulateCaptions backfill, numPr+SEQ conflict guard |
| `lib/parse/caption-resolver.ts` | added |
| `lib/config/config-schema.ts` | + `captions` / `CaptionEntry` / `CaptionFormat` / `subCounter` schemas |
| `lib/config/edit-config-schema.ts` | + `CaptionBlock`, `CaptionCounterReset`, `edit-caption` op; `EquationBlock` extended (`omml`, `captionId`, `subGroup`) |
| `lib/config/config-types.ts` | + ResolvedCaptionConfig surface types |

### Tools and docs (`skill/`)

| File | Change |
|---|---|
| `skill/tools/inspect-caption.ts` | added |
| `skill/tools/migrate-captions.ts` | added (read-only detector for manually-numbered candidates) |
| `skill/tools/overview.ts` | + Captions section (skip-if-empty) |
| `skill/references/captions.md` | added |
| `skill/references/equations.md` | rewritten for captions+EquationBlock pattern |
| `skill/references/cross-references.md` | + caption-class display routing |
| `skill/SKILL.md` | Match-content-shape updated |
| `CLAUDE.md` | layout description refreshed |
| `tsdown.config.ts` | + 2 entries (inspect-caption, migrate-captions) |

## 6. Verification

- Typecheck / lint / fmt / build clean at each commit.
- Existing fixtures (`_config_forward_ref.json`,
  `_config_thesis_proposal.json`) migrated to the new pattern and apply
  with no regression. Outputs (`_filled_forward_ref.docx`,
  `_filled_thesis_proposal.docx`) re-verified by re-parsing.
- New fixtures exercise the additions:
  - `_config_caption_basic.json` — Equation + Figure + chapter-restart
    + forward and backward refs.
  - `_config_edit_caption.json` — round-trip the `edit-caption` op on
    `_filled_caption_basic.docx`; bookmark preserved, body swapped.
- Byte-level XML inspection on output zips for SEQ / STYLEREF
  instrText, bookmark wrap, REF result backfill (per
  MEMORY.md guidance: typecheck can't catch wrong-parent mutations).
- Word round-trip on `_filled_caption_basic.docx`: F9 refresh produces
  the same numbers the counter sim emitted; References →
  Cross-reference dialog lists the captions under their identifiers.

No automated tests in this repo (per CLAUDE.md); verification is
fixture-driven plus byte-level inspection.

## 7. Migration notes

Existing apply configs need:

1. Drop `numbering[]` entries whose `styleId` was a caption-class style
   (`EquationNumber`, `FigureCaption`, `TableCaption`, ...). Apply now
   rejects `styleId` appearing in both `numbering` and `captions` with a
   migration hint.
2. Add a `captions` table entry for each caption identifier. `styleId`
   is required; `chapterPrefix` is an array of styleIds, outermost
   first.
3. Replace `{ type: "paragraph", styleId: "FigureCaption", ... }` plus
   `numbering[]` binding with `{ type: "caption", captionId: "Figure",
   text: "...", anchor?: "..." }`.
4. On `EquationBlock`: drop `displayStyle`, drop `numbering`, drop the
   `math: { latex }` wrapper (now flat `latex` or `omml`); add
   `captionId` when the equation should be numbered; `anchor` requires
   `captionId`.
5. Use `subGroup: "start"` / `"continue"` on `EquationBlock` for
   `(2a)` `(2b)` sequences (requires `captions[id].subCounter`).
6. `CaptionCounterReset` block restarts a counter mid-doc; emits a
   hidden SEQ paragraph.

`migrate-captions <doc>` detects manually-numbered caption-shaped
paragraphs in a source document and lists candidates (read-only — the
agent decides identifier mapping).

## 8. Known limitations / follow-up

- `inspect-range` / `inspect-runs` / `inspect-table` / `import-template`
  still emit raw fldChar XML rather than structured `field-parse`
  output. Deferred — non-blocking for caption pipeline correctness;
  current output remains valid for debugging. See commit `48c843b`
  body for the explicit defer note.
- `migrate-captions` is read-only in v1. Identifier mapping
  (which prefix → which identifier? which chapter prefix style?)
  needs agent context, so the tool detects and points; the agent
  writes the apply config.
- `edit-caption` blocker error message: the generic field-blocker
  scan already refuses `replace_paragraph` / `delete_paragraph_runs`
  on field-bearing paragraphs (captions contain SEQ / STYLEREF), so
  agents get the generic error rather than an "use edit_caption"
  hint. Documented in `references/captions.md`; a more pointed error
  would be a small docs/UX improvement, not a correctness fix.
- MathML as direct input is intentionally not accepted. Two sources
  (`latex` / `omml`) is enough surface; agents with MathML upstream
  convert client-side. See spec §9.9.

## 9. Commits

```
152bdc3 refactor(edit): extract complex-field skeleton; move REF emitter to fields/
0765d40 feat(edit): SEQ field emitter
2840b05 feat(edit): STYLEREF field emitter
523c3d7 feat(edit): field-parse module — REF / SEQ / STYLEREF
981da56 feat(edit): caption counter simulator
50e945c feat(edit): caption-class layout emitter
610001a feat(config): captions table + CaptionBlock + CaptionCounterReset; extend EquationBlock
18d7b35 feat(apply): integrate caption pipeline end-to-end
dd0cc9f feat(edit): edit-caption op handler
06e7309 feat(apply): standardize re-emit existing captions
c4f9a15 refactor(apply): migrate fixtures to captions table; reject numPr+SEQ conflict
3064c3a feat(tools): inspect-caption + migrate-captions
48c843b feat(tools): overview Captions section
24cef75 docs(skill): rewrite caption pattern for SEQ + STYLEREF model
```
