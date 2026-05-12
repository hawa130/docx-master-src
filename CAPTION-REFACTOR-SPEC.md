# Caption-class numbering refactor

Migrate caption-class enumerators (equation numbers, figure / table /
theorem captions) from `numPr` binding to Word-native `SEQ` + `STYLEREF`
fields. Outline / list numbering stays on `numPr`.

This doc is the implementation reference. Update it as decisions land.

## 1. Motivation

Word's data model splits numbered things into two categories:

- `numPr` / `numbering.xml`: outline numbering (H1–H6), bullet / ordered
  lists. Counters live as paragraph properties; Word renders `lvlText` at
  layout time.
- `SEQ` fields: caption-class enumerators. Counters live as document-local
  field identifiers; each `SEQ <id>` occurrence increments the identifier.

We picked `numPr` for caption-class blocks because the toolchain already
handled multi-level numbering. The result: caption-class blocks share a
mechanism designed for outline structure, with three structural
side-effects — chapter prefix forces outline coupling (H1 and equation
co-tenant the same `numId`), multi caption-type independence requires
one `abstractNum` per type, and Word's References → Cross-reference
dialog can't surface our captions because it expects SEQ.

Best-practice consensus across Word community sources (UWaterloo IST,
Microsoft Q&A, StataTex, FormulAI) is unanimous: SEQ + STYLEREF +
bookmark + REF `\h`. We're aligning.

## 2. Scope

**Changing**: equation number rendering, figure / table caption
rendering, custom caption types (theorem / lemma / ...), cross-reference
routing for caption-class targets, standardize handling of existing
captions, edit operations on captions.

**Not changing**: H1–H6 outline numbering, body lists (bullet /
ordered), multi-level lists. These stay on `numPr`.

**Design decisions** (these are positive choices, not deferred features):

- **`styleId` only in `chapterPrefix` entries**, no `styleName` alternate.
  Locale-stable, consistent with codebase indexing.
- **No implicit default `captions`**. Agent must declare. Ref doc
  provides copy-paste templates.
- **No manual single-caption number override**. Counter is authoritative.
  To break a sequence: use a separate identifier or `CaptionCounterReset`.
  To restart at a point: `CaptionCounterReset`.
- **Inline equations don't carry captions**. Only `EquationBlock` (display
  equations) accept `captionId`.
- **`display: "full"` on EquationBlock anchor throws**. No body text to
  return.

## 3. Schema

### 3.1 `captions` (top-level config)

```jsonc
"captions": {
  "<identifier>": {
    "prefix": "<string>",                   // default ""
    "suffix": "<string>",                   // default ""
    "format": "arabic" | "alphabetic" | "ALPHABETIC"
            | "roman" | "ROMAN" | "chinese" | "chinese-formal",
                                            // default "arabic"
    "chapterPrefix": ["<styleId>", ...],    // default []; any depth
    "chapterSeparator": "<string>",         // default "."
    "bodySeparator": "<string>",            // default " "
    "styleId": "<string>",                  // required
    "subCounter": {                         // optional, enables subequations
      "format": "arabic" | "alphabetic" | "ALPHABETIC"
              | "roman" | "ROMAN",          // default "alphabetic"
      "prefix": "<string>",                 // default ""
      "suffix": "<string>"                  // default ""
    }
  }
}
```

`<identifier>` is a free string. Block-level `captionId` references this
key. Conventional values: `"Equation"`, `"Figure"`, `"Table"`,
`"Theorem"`, etc. — convention lives in ref docs, not the schema.

`format` → SEQ `\*` switch:

| schema value | SEQ switch | Renders |
|---|---|---|
| `"arabic"` | `\* ARABIC` | 1, 2, 3 |
| `"alphabetic"` | `\* alphabetic` | a, b, c |
| `"ALPHABETIC"` | `\* ALPHABETIC` | A, B, C |
| `"roman"` | `\* roman` | i, ii, iii |
| `"ROMAN"` | `\* ROMAN` | I, II, III |
| `"chinese"` | `\* CHINESENUM2` | 一, 二, 三 |
| `"chinese-formal"` | `\* CHINESENUM3` | 壹, 贰, 叁 |

`chapterPrefix` is an ordered array of `styleId` strings, outermost to
innermost. Any depth supported (rendered as
`STYLEREF[0].STYLEREF[1].…STYLEREF[N].SEQ`). `[]` = global counter, no
restart. SEQ's `\s` switch derived from the last entry's outline level.

`subCounter` enables subequation grouping. When present, `EquationBlock`
can carry `subGroup: "start" | "continue"` to render (1a)(1b) sequences.
Without `subCounter`, declaring `subGroup` throws.

### 3.2 `EquationBlock`

```jsonc
{
  "type": "EquationBlock",
  "latex": "<string>",                     // exactly one of latex / omml required
  "omml": "<string>",                       // escape hatch when temml fails on the LaTeX
  "styleId": "<string>",                   // optional, default "Equation"
  "captionId": "<string>",                 // optional; omit = unnumbered
  "subGroup": "start" | "continue",        // optional; requires captionId + subCounter
  "anchor": "<string>"                     // optional; requires captionId
}
```

**Subequation semantics**:
- `subGroup: "start"`: parent counter advances; sub-counter resets to 1.
  Rendered: `(2a)` (with `Equation` captionId, chapter 2, sub at index 1).
- `subGroup: "continue"`: parent counter does NOT advance; sub-counter
  advances. Rendered: `(2b)`, `(2c)`, ...
- Omitted: standalone. Closes any open sub-group implicitly. Parent
  advances; no sub-counter rendered.

Layout:
- With `captionId` (any subGroup or none): 3-col borderless table. Middle
  cell = OMML paragraph (style: `EquationBlock.styleId`). Right cell =
  caption paragraph (style: `captions[captionId].styleId`).
- Without `captionId`: single centered paragraph
  (style: `EquationBlock.styleId`). No table.

### 3.3 `CaptionBlock`

Replaces `FigureCaption` / `TableCaption`.

```jsonc
{
  "type": "CaptionBlock",
  "captionId": "<string>",                 // required
  "text": "<string>",                       // required (may be "")
  "anchor": "<string>"                      // optional
}
```

Single paragraph; style from `captions[captionId].styleId`. Content
sequence: prefix + STYLEREFs (joined by chapterSeparator) + SEQ + suffix
+ bodySeparator + text (the last two omitted when text is empty).

### 3.4 `CaptionCounterReset`

```jsonc
{
  "type": "CaptionCounterReset",
  "captionId": "<string>",                 // required
  "newValue": <integer>                     // optional, default 1
}
```

Standalone marker block. Engine emits a hidden SEQ field at this
position: `{ SEQ <identifier> \r <newValue> \h }`. Counter sim resets
accordingly.

### 3.5 `edit_caption` op (apply_edits)

```jsonc
{
  "op": "edit_caption",
  "target": { "anchor": "<name>" } |
            { "captionId": "<id>", "index": <N> },
  "text": "<string>"                       // required
}
```

Replaces caption body text — runs after the primary anchor's
`bookmarkEnd`, preserving fields and bookmark. Target by name (when
anchor was declared) or by `(captionId, 1-based index)` (counts caption
paragraphs in body order matching the identifier).

Throws when target is `EquationBlock` (no body).

### 3.6 Cross-field rules

| Rule | When | Action |
|---|---|---|
| `EquationBlock` has neither `latex` nor `omml` | Schema validation | Throw |
| `EquationBlock` has both `latex` and `omml` | Schema validation | Throw — mutex |
| `EquationBlock.anchor` set, `captionId` omitted | Schema validation | Throw |
| `EquationBlock.subGroup` set, `captionId` omitted | Schema validation | Throw |
| `EquationBlock.subGroup` set, `captions[captionId]` has no `subCounter` | Pre-scan | Throw |
| `subGroup: "continue"` without preceding `subGroup: "start"` in same identifier scope | Pre-scan | Throw |
| `captionId` references undeclared `captions` key | Pre-scan | Throw |
| `chapterPrefix` references unknown `styleId` | Pre-scan against styles.xml | Throw |
| `chapterPrefix` references `styleId` not bound to outline numbering | Pre-scan | Warn (STYLEREF will render 0) |
| `InlineRef.display = "full"` targets caption-class anchor on EquationBlock | Pre-scan | Throw |
| `InlineRef.display = "full"` targets caption-class anchor on CaptionBlock | Pre-scan | Allocate paragraph-wide bookmark internally |
| `replace_paragraph` / `delete_paragraph_runs` targets a caption paragraph | Blocker scan | Block with "use edit_caption" hint |

## 4. Engine architecture

### 4.1 Complex-field skeleton (shared)

`lib/edit/fields/complex-field.ts`. The 5-run skeleton:

```
<w:r>[ rPr? ]<w:fldChar fldCharType="begin"/></w:r>
<w:r>[ rPr? ]<w:instrText xml:space="preserve"> <code> </w:instrText></w:r>
<w:r>[ rPr? ]<w:fldChar fldCharType="separate"/></w:r>
<w:r>[ rPr? ]<w:t>[result placeholder]</w:t></w:r>
<w:r>[ rPr? ]<w:fldChar fldCharType="end"/></w:r>
```

`applyFieldFormat` (rPr replication + MERGEFORMAT) shared. Field-specific
modules (REF / SEQ / STYLEREF) supply instrText body, initial result,
MERGEFORMAT suppression flag.

Hidden field: SEQ's `\h` switch is the canonical hide mechanism (Word
suppresses field render). No additional rPr decoration needed — the
switch handles it entirely.

### 4.2 SEQ field emitter

`lib/edit/fields/seq-field.ts`. `emitSeqField(ownerDoc, spec)`:

```ts
interface SeqFieldSpec {
  identifier: string
  format: SeqFormat
  restartAtOutlineLevel?: number    // → \s N
  resetTo?: number                   // → \r N
  repeat?: boolean                   // → \c (repeat last without increment)
  hidden?: boolean                   // → \h
  initialResult?: string             // default ""
  format_rpr?: RunFormat
}
```

instrText composition:

```
SEQ <id> \* <FORMAT> [\s <N>] [\r <N>] [\c] [\h] [\* MERGEFORMAT]
```

Switch order matches Word's UI output. `\c` and `\r` are mutually
exclusive (engine enforces).

### 4.3 STYLEREF field emitter

`lib/edit/fields/styleref-field.ts`. `emitStyleRefField(ownerDoc, spec)`:

```ts
interface StyleRefFieldSpec {
  styleName: string          // resolved from styleId upstream
  switches: string[]         // typically ["\\n"]
  initialResult?: string
  format_rpr?: RunFormat
}
```

instrText composition: `STYLEREF "<styleName>" \n [\* MERGEFORMAT]`.
styleName is always double-quoted.

### 4.4 Field-parse module

`lib/edit/fields/field-parse.ts`. Walks a run sequence and reduces
complete fldChar begin → end groups to structured form.

```ts
type ParsedRun =
  | { kind: "text"; text: string; rPr?: ParsedRPr }
  | { kind: "field"; fieldType: "REF" | "SEQ" | "STYLEREF" | "OTHER";
      instrText: string; result: string; rPr?: ParsedRPr;
      details: FieldDetails }

interface FieldDetails {
  bookmarkName?: string                              // REF
  switches?: string[]                                // REF
  identifier?: string                                // SEQ
  format?: SeqFormat                                 // SEQ
  restartAtOutlineLevel?: number                     // SEQ \s
  resetTo?: number                                   // SEQ \r
  repeat?: boolean                                   // SEQ \c
  hidden?: boolean                                   // SEQ \h
  styleName?: string                                 // STYLEREF
}
```

Consumed by `inspect-range`, `inspect-runs`, `migrate-captions`,
standardize re-emit, and `inspect-caption`.

### 4.5 Caption layout emitter

`lib/edit/caption-emit.ts`:

```ts
// EquationBlock with captionId
emitNumberedEquation(spec: {
  mathSource: { latex: string } | { omml: string }
  equationStyleId: string
  captionConfig: ResolvedCaptionConfig
  subGroup?: "start" | "continue"
  anchor?: string
  emitCtx: EmitContext
}): Element  // <w:tbl>

// EquationBlock without captionId
emitUnnumberedEquation(spec: {
  mathSource: { latex: string } | { omml: string }
  equationStyleId: string
  emitCtx: EmitContext
}): Element  // <w:p>

// CaptionBlock
emitCaptionBlock(spec: {
  captionConfig: ResolvedCaptionConfig
  text: string
  anchor?: string
  emitCtx: EmitContext
}): Element  // <w:p>

// CaptionCounterReset
emitCaptionReset(spec: {
  captionConfig: ResolvedCaptionConfig
  newValue: number
  emitCtx: EmitContext
}): Element  // <w:p> with single hidden SEQ run
```

`ResolvedCaptionConfig`:

```ts
interface ResolvedCaptionConfig {
  identifier: string
  prefix: string
  suffix: string
  format: SeqFormat
  chapterPrefix: Array<{ styleName: string; outlineLevel: number }>
  chapterSeparator: string
  bodySeparator: string
  paragraphStyleId: string
  restartAtOutlineLevel: number | undefined   // last entry's outline level
  subCounter: { format: SeqFormat; prefix: string; suffix: string } | undefined
}
```

Resolution happens once at apply start; emitters consume the resolved
form. Avoids re-parsing styles.xml per emit.

### 4.6 Caption counter simulator

`lib/edit/caption-counter.ts`. Runs after outline numbering simulator
(consumes outline counters for STYLEREF resolution).

State per identifier: parent counter (integer) + sub-counter (integer,
when subCounter declared) + `openSubGroup` flag.

Walk body in order:
- On outline-numbered heading: update "latest heading by styleName" map
  (for STYLEREF resolution)
- On caption paragraph (CaptionBlock / EquationBlock with captionId):
  - If `subGroup` omitted: parent++, sub=0, close open subgroup
  - If `subGroup: "start"`: parent++, sub=1, open subgroup
  - If `subGroup: "continue"`: parent unchanged, sub++ (requires open)
- On CaptionCounterReset paragraph: parent = newValue, sub=0, close
  any open subgroup
- Render: prefix + chapter STYLEREFs (joined by chapterSeparator) +
  formatted-parent + (subCounter.prefix + formatted-sub + subCounter.suffix
  when sub > 0) + suffix

Output: `Map<Element, string>` mapping caption paragraph elements to
rendered text. Drives `pendingBackfills` placeholder text.

### 4.7 REF routing for caption-class targets

| `display` | Outline target (H1–H6) | Caption-class target |
|---|---|---|
| `"label"` | REF `\n` | REF `\h` on primary bookmark (returns prefix+chapter+SEQ+suffix) |
| `"number"` | REF `\r` | REF `\h` on primary bookmark (same as label) |
| `"full"` | REF `\h` on paragraph bookmark | REF `\h` on **paragraph-wide** bookmark (allocated on demand) |

For caption-class targets, `"label"` and `"number"` collapse — both
return the SEQ result with full decoration. Document the equivalence in
cross-references.md.

Bookmark allocation per anchored caption block:
- **Primary bookmark** (named after agent's `anchor`): wraps prefix run
  through suffix run (the number with decoration).
- **Secondary bookmark** (auto-named `_Ref<8hex>`): wraps the whole
  paragraph. Allocated only when an `InlineRef` with `display: "full"`
  targets the anchor.

### 4.8 Standardize re-emit existing captions

When apply runs on a doc that already contains SEQ-based caption
paragraphs (from prior apply or Word UI Insert Caption):

1. **Detect**: scan body for paragraphs where (a) paragraph style
   matches some `captions.<id>.styleId` AND (b) paragraph contains a
   SEQ field whose identifier matches `<id>`.
2. **Re-emit**: replace the pre-body run sequence (from paragraph start
   to bodySeparator inclusive) with a freshly-emitted version using
   current `captions[id]` config. Preserves:
   - Bookmark (name and pair structure) wrapping the new number runs
   - Body text after bodySeparator
   - Paragraph style
   - SEQ identifier (Word's running counter state continues)
3. **Identifier mismatch** (SEQ exists but identifier doesn't match any
   declared caption): pass through unchanged — don't disturb the
   existing SEQ structure. Warn so the misconfiguration is visible. To
   bring the caption under standardize's purview, agent adds the
   identifier to `captions` config in next apply; to remove a residual
   caption entirely, use `delete_paragraph`.

`lib/apply/standardize-captions.ts`. Runs after style application,
before counter sim.

### 4.9 `edit_caption` op

`lib/edit/edit-caption-op.ts`. Target resolution:
- `{ anchor: name }`: lookup via BookmarkAllocator
- `{ captionId, index }`: walk body, find Nth caption paragraph with
  matching identifier (1-based)

Engine logic: locate paragraph → find primary `bookmarkEnd` → remove
subsequent runs (the body) → insert new runs (bodySeparator + text in
caption paragraph's style cascade) → leave everything before bookmarkEnd
intact.

### 4.10 Blocker enforcement for caption paragraphs

`lib/edit/blockers.ts` adds detection:
- Paragraph contains SEQ field
- Paragraph style ∈ {captions.<id>.styleId | id ∈ apply config}

Blocked operations: `replace_paragraph`, `delete_paragraph_runs`
covering field runs. Error message points agent to `edit_caption` (for
body changes) or `delete_paragraph` (for whole removal).

## 5. Tooling

### 5.1 New tools

**`inspect-caption <doc> <identifier>`** — per-identifier view; parallel
to `inspect-style-def`. Output:

```
Caption: Equation
  Resolved config:
    style:            EquationNumber
    prefix / suffix:  "(" / ")"
    format:           arabic
    chapter prefix:   Heading 1 ("标题 1", outline level 1)
    chapter sep:      "."
    sub-counter:      alphabetic
    SEQ field:        SEQ Equation \* ARABIC \s 1

  Occurrences (12):
    para 23   counter (1.1)   anchor: eq_intro_1    refs: 2 (paras 45, 51)
    para 28   counter (1.2)   anchor: eq_intro_2    refs: 0
    para 67   counter (2.1a)  anchor: eq_method_1a  refs: 0   [subGroup: start]
    para 68   counter (2.1b)  anchor: (none)        refs: 0   [subGroup: continue]
    ...

  Cross-reference summary:
    targets cited:    8 of 12 occurrences carry anchors
    citations total:  17 InlineRef pointing into this identifier
```

**`migrate-captions <doc>`** — detect manually-numbered caption-shaped
paragraphs (e.g. paragraph starts with `(N)` / `(N.M)` / `图 N.M` /
`Figure N.M` in a caption-class style) and convert to SEQ-based.

```
migrate-captions <doc> [--captionId <prefix-pattern>=<id> ...] [--dry-run]
```

Dry-run prints candidates. Default mode applies in place. The
`--captionId` flag maps detected text patterns to identifiers (e.g.
`--captionId "图"=Figure`). Without explicit mapping, tool warns and
skips.

### 5.2 Enhanced tools

**`overview`** — adds Captions section listing SEQ identifiers found in
body (skip-if-empty). Per identifier: style, format, chapter prefix,
occurrence count. Does not list `prefix` / `suffix` / `bodySeparator` —
agent runs `inspect-caption` for those details.

**`inspect-range`** / **`inspect-runs`** — output structured field
descriptions (via field-parse module) instead of raw fldChar XML when
paragraph contains complex fields.

**`inspect-table`** — adds `caption` classification: 3-col borderless +
middle cell contains OMML paragraph + right cell contains SEQ field.

**`import-template`** — scans template body for SEQ identifiers;
suggests `captions` table entries (style inferred from paragraph
style, chapter prefix inferred from co-occurring STYLEREF, format from
SEQ `\*` switch).

## 6. Migration

### 6.1 Removed code paths

- `lib/apply/numbering-apply.ts`: drop emission of `abstractNum` entries
  for caption-class styleIds; remove caption-role recognition in
  standardize
- `lib/config/config-schema.ts`: `FigureCaption` / `TableCaption` block
  types removed
- `EquationBlock.math` wrapper removed (flat `latex` / `omml`)
- `EquationBlock.displayStyle` removed
- `EquationBlock.numbering` field removed
- `lib/edit/field-ref.ts` relocates into `lib/edit/fields/`

### 6.2 Added code paths

- `lib/edit/fields/complex-field.ts` — shared 5-run skeleton
- `lib/edit/fields/ref-field.ts` — REF emitter (moved + thinned)
- `lib/edit/fields/seq-field.ts` — SEQ emitter
- `lib/edit/fields/styleref-field.ts` — STYLEREF emitter
- `lib/edit/fields/field-parse.ts` — field parser
- `lib/edit/caption-emit.ts` — equation / caption / reset layout emitters
- `lib/edit/caption-counter.ts` — caption counter simulator
- `lib/edit/edit-caption-op.ts` — edit_caption op handler
- `lib/apply/standardize-captions.ts` — re-emit existing captions
- `lib/config/captions-schema.ts` — zod schema for captions table +
  CaptionBlock + CaptionCounterReset + edit_caption op + EquationBlock
  shape update
- `lib/parse/caption-resolver.ts` — captionId → ResolvedCaptionConfig
- `skill/tools/inspect-caption.ts`
- `skill/tools/migrate-captions.ts`

### 6.3 Test fixtures

All `_config_*.json` with caption-class blocks need migration:
- Replace `FigureCaption` / `TableCaption` types → `CaptionBlock` +
  `captionId`
- Drop `numbering: { numId, level }` on caption blocks
- Add top-level `captions` config matching the prior lvlText pattern
- EquationBlock: `math: { latex }` → flat `latex` (or `omml` escape
  hatch); drop `displayStyle` / `numbering`; add `captionId`

Migration is one-shot — no dual-path period.

## 7. Phases

### Phase 1: Engine primitives

Field emitters, complex-field skeleton, field-parse module. Not yet
integrated.

Deliverables:
- `lib/edit/fields/{complex-field, ref-field, seq-field, styleref-field, field-parse}.ts`
- Unit tests: byte-level XML comparison per emitter; round-trip parse
  for field-parse

Verification: typecheck / lint / fmt:check / build clean; new tests
pass; no regression on existing REF-using fixtures.

Suggested commit: `refactor(edit): extract complex-field skeleton; add SEQ/STYLEREF emitters + field parser`

### Phase 2: Caption pipeline

Counter simulator (with subequation + reset support), caption layout
emitter.

Deliverables:
- `lib/edit/caption-counter.ts` (parent / sub counters, chapter
  resolution, reset handling)
- `lib/edit/caption-emit.ts` (3-col table for numbered equation; single
  paragraph for unnumbered equation and CaptionBlock; hidden-SEQ
  paragraph for CaptionCounterReset)
- `pendingBackfills` extended for caption placeholders

Verification: counter sim regression tests (linear, chapter restart,
multi-identifier, subequation, CaptionCounterReset); placeholder text
matches Word F9 output for representative cases.

Suggested commit: `feat(edit): caption counter simulator + caption-class layout emitter`

### Phase 3: Schema integration + edit ops + standardize re-emit

Switch schema, integrate engine paths, add edit_caption op + blocker
enforcement, add standardize re-emit, remove old code, migrate fixtures.

Deliverables:
- Schema: `captions` table, `CaptionBlock`, `CaptionCounterReset`,
  `edit_caption` op, `EquationBlock` shape update (latex / omml +
  subGroup + captionId)
- Engine: edit-engine and apply-styles route through new emitters;
  `lib/edit/edit-caption-op.ts` implements body replacement
- Blocker scan: caption-paragraph detection blocks paragraph-level
  destructive ops
- Standardize re-emit: `lib/apply/standardize-captions.ts`
- All `_config_*.json` fixtures migrated
- Old caption-numbering paths deleted

Verification: all fixtures apply cleanly; subequation + reset fixtures
produce correct counter sequence; Word round-trip on 2+ fixtures.

Suggested commits:
1. `feat(config): captions + CaptionBlock + CaptionCounterReset; refactor EquationBlock`
2. `feat(edit): edit_caption op + caption-paragraph blocker scan`
3. `feat(apply): standardize re-emit existing captions`
4. `refactor(apply): retire numPr binding for caption-class roles; migrate fixtures`

### Phase 4: Tools + skill docs

New tools, enhanced tools, ref doc rewrite.

Deliverables:
- `skill/tools/inspect-caption.ts`
- `skill/tools/migrate-captions.ts`
- `overview.ts` — Captions section
- `inspect-range.ts` / `inspect-runs.ts` — field-parse structured output
- `inspect-table.ts` — caption classification
- `import-template.ts` — SEQ identifier detection in template body
- `skill/references/equations.md` rewritten
- `skill/references/cross-references.md` caption-refs section updated
- `skill/references/captions.md` new
- `skill/SKILL.md` Match-content-shape section reflects new caption
  pattern
- `CLAUDE.md` updated (new module references)

Verification: skill-creator audit (read-only subagent) flags
inconsistencies; manual scan for any residual `numbering[] +
EquationNumber` pattern in docs; `dist/docx-master/SKILL.md` inspection.

Suggested commits:
1. `feat(tools): inspect-caption + migrate-captions; field-parse output in inspect-range/runs`
2. `feat(tools): caption classification in inspect-table; SEQ detection in import-template; overview Captions section`
3. `docs(skill): rewrite caption pattern for SEQ + STYLEREF model`

## 8. Test plan

### 8.1 Unit tests (Phase 1)

- SEQ emit per `format` value: assert instrText byte content
- SEQ switches: `\s N`, `\r N`, `\c`, `\h`; mutually-exclusive flags
  (`\c` + `\r`) throw
- STYLEREF emit with styleName containing space (quoted)
- Format-bearing field: rPr replicated; `\* MERGEFORMAT` appended iff
  rPr non-empty
- Field-parse: round-trip each field type back from XML, verify
  structured details

### 8.2 Counter simulator tests (Phase 2)

- Linear increment (no chapter): 1, 2, 3, ...
- H1 chapter restart: 1.1, 1.2, 2.1, 2.2
- Multi-level chapter (depth 2): 1.1.1, 1.1.2, 1.2.1
- Multi-identifier independence: Figure 1, Figure 2, Table 1, Figure 3
- Subequation grouping: (1) (2a) (2b) (3a) (3b) (3c) (4)
- CaptionCounterReset: counter resets to specified value mid-doc
- chinese / chinese-formal / roman / ALPHABETIC format renderings

### 8.3 E2E tests (Phase 3)

- Migrated `_config_forward_ref.json`: forward ref resolves
- Migrated `_config_thesis_proposal.json`: chapter-prefixed equations
- New `_config_multi_caption.json`: Equation + Figure + Table + Theorem,
  cross-refs between
- New `_config_subequation.json`: subequation groups + refs targeting
  parent and sub members
- New `_config_caption_reset.json`: CaptionCounterReset mid-doc
- New `_config_edit_caption.json`: edit_caption op modifying existing
  captions; verify number + bookmark preserved

### 8.4 Word round-trip (Phase 3 spot-check)

- Open ≥ 2 e2e outputs in Word
- Ctrl+A → F9 to refresh fields
- Verify rendered numbers match counter sim placeholders
- Verify References → Cross-reference dialog lists captions under their
  identifiers
- Insert a new caption via Word UI on the same identifier; F9; verify
  it joins our sequence

### 8.5 Tool tests (Phase 4)

- `inspect-caption`: output matches expected structured form for a doc
  with multiple identifiers, subequations, refs
- `migrate-captions`: dry-run lists expected candidates; full run
  produces valid SEQ-based captions; cross-refs to migrated targets
  resolve
- `inspect-range` / `inspect-runs`: field-parse output matches expected
  for paragraphs with REF / SEQ / STYLEREF
- `inspect-table`: classifies caption table; doesn't misclassify
  data / form tables

## 9. Implementation notes

### 9.1 SEQ `\s N` derivation

Word's `\s` switch takes outline level (1-9, 1-indexed). `styles.xml`
stores `<w:outlineLvl w:val="..."/>` 0-indexed. Conversion at resolution
time: `styleId` → outlineLvl val → `+1` for SEQ.

Multi-entry `chapterPrefix`: SEQ `\s` uses the **last** entry's outline
level (deepest chapter level controls restart).

### 9.2 STYLEREF styleName resolution

`STYLEREF "<name>" \n` takes the style's display name. Engine resolves
`styleId` → `<w:name w:val="..."/>` from styles.xml at apply start
(cached in `ResolvedCaptionConfig`). Missing style name → throw.

### 9.3 Bookmark scope

Primary bookmark wraps just the number + decoration (prefix run through
suffix run). For CaptionBlock with body, this stops before
bodySeparator. For EquationBlock, this is the entire right-cell content.

Secondary bookmark (auto-named) wraps the full paragraph when a
`display: "full"` reference targets the anchor; allocated on demand
during pre-scan.

### 9.4 Apply-pipeline ordering

1. Parse apply config; resolve `captions[*]` against styles.xml into
   `ResolvedCaptionConfig` (styleName, outlineLevel)
2. Standardize: paragraph-style application + existing-caption re-emit
   (just rebuilds XML structure; counter values come later)
3. New caption emit: agent's edit ops produce new caption paragraphs
4. Outline numbering simulator (existing, unchanged) — produces outline
   counter renderings
5. Caption counter simulator (new) — consumes outline counter outputs
   for STYLEREF chapter prefix resolution; produces caption placeholder
   text
6. Placeholder backfill — writes computed text into REF result runs
   (both outline-target and caption-target pendings)

### 9.5 Subequation counter logic

Counter sim tracks `{ parent: number, sub: number, openSubGroup: bool }`
per identifier with rules:

- `subGroup` omitted: `parent++; sub = 0; openSubGroup = false`
- `subGroup: "start"`: `parent++; sub = 1; openSubGroup = true`
- `subGroup: "continue"`: requires `openSubGroup == true`; else throws
  at pre-scan. `sub++`. Parent unchanged.

Field emission produces SEQ switches that yield the same sequence under
Word's F9 — exact `\c` / `\r` switch combination determined during
phase 2 and verified by Word round-trip (counter sim placeholder must
match Word's rendered value after F9).

### 9.6 CaptionCounterReset emission

Generated paragraph: single hidden-SEQ run. Paragraph style: `Normal`
(invisible content, no visual presence). Counter sim consumes the
reset event; downstream captions render with restarted counter.

### 9.7 Standardize re-emit identifier matching

When standardize encounters a paragraph that:
- Has paragraph style ∈ declared captions' styleIds AND
- Contains a SEQ field whose identifier matches one declared in
  `captions`

→ re-emit. The detection uses field-parse module on the paragraph's
runs.

When identifier mismatches all declared captions: pass through
unchanged and warn. Agent's options: add the identifier to `captions`
config to bring it under management, or remove the paragraph via
`delete_paragraph` if it's residual.

### 9.8 chapterPrefix style without outline numbering

If `captions[id].chapterPrefix` references a styleId whose style has no
`numPr`, Word's STYLEREF returns 0. Counter sim mirrors this (returns
"0"). Pre-scan warns at apply start so misconfiguration is visible
before output is produced.

### 9.9 EquationBlock omml escape hatch

When `omml` is provided directly, engine bypasses temml entirely;
embeds the provided OMML in the middle cell as-is, wrapped in
`<m:oMathPara>`. Schema validation ensures exactly one of `latex` /
`omml` is set.

This is the escape hatch for LaTeX expressions temml fails on. Agents
with MathML upstream convert to OMML themselves (via mathml2omml or
equivalent) and pass `omml`. MathML is not accepted as a direct input
— two sources is enough surface area, and the MathML→OMML conversion
isn't differentiated from what an agent can do client-side.

## 10. Overview surfacing

After this refactor, `overview` adds a Captions section listing
SEQ-based caption identifiers detected in the source document body.
Skip-if-empty: no SEQ fields → no section.

Per identifier listed: style, format, chapter prefix, occurrence count.
Decoration details (`prefix` / `suffix` / `bodySeparator`) require
`inspect-caption` for full view.

Identifiers used outside caption context (page-internal counters,
custom enumerations) appear here too — agent judges their nature from
the identifier name + paragraph style context.
