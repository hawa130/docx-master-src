# Header / Footer

Generates header and footer parts and binds them to every section. Sparse-by-design: omit the `headerFooter` block entirely to leave existing HF references untouched. When the block IS declared, every sectPr is rebound and any pre-existing HF parts become orphans in the archive (Word ignores unreferenced parts).

## Shape

```jsonc
headerFooter: {
  header: {
    default: [ ...Block[] ],   // applies to pages not covered by first / even
    first:   [ ...Block[] ],   // optional. First page of each section.
                               //   Auto-sets <w:titlePg/> on every sectPr.
    even:    [ ...Block[] ],   // optional. Even-numbered pages.
                               //   Auto-sets <w:evenAndOddHeaders/> in settings.
  },
  footer: { default, first?, even? },  // same shape
}
```

At least one of `header` / `footer` is required, and each declared surface needs at least one variant. Empty array `[]` is legal — means "this variant exists for the trigger flag but renders nothing" (used to blank a cover page's header while keeping `titlePg` set).

To blank only the first page, declare both `first: []` AND `default: [...]` — omitting `default` leaves non-first, non-even pages with no header/footer at all.

## Block subset

Inside header / footer content, only these block types are allowed (rejected at config parse, including any nesting depth — table cells too):

- `paragraph` — text, optionally styleId-bound. Inline nodes (text, fields, hyperlinks) work the same as in body edits.
- `image`
- `table` — common pattern for split layouts via a single-row borderless 3-column table.
- `horizontal-rule`

Disallowed: `caption`, `caption-counter-reset`, `equation`, `page-break`. The first three bind to body-side counters; emitting them inside HF double-increments and breaks the counter sim. Page breaks inside HF are no-ops (Word ignores).

## Paragraph styleId restrictions

HF paragraphs cannot use heading or body-text styleIds (`Heading1..9`, `Title`, `Subtitle`, `BodyText`) — also rejected at any nesting depth. These carry outline-level / numbering bindings that misbehave outside the body.

Use the built-in `Header` / `Footer` styleId (engine auto-injects them when missing) with a `paraFormat` / `runFormat` override for typography. Custom non-heading styleIds declared in top-level `styles[]` also work.

## Variant semantics

ECMA-376 sectPr supports three reference types per surface; engine emits whichever variants the config declares:

- **default** — applies to every page not covered by `first` or `even`.
- **first** — section's first page only. Engine sets `<w:titlePg/>` on every sectPr when this variant is declared anywhere; clears the flag when the variant is dropped on a re-run.
- **even** — even-numbered pages. Engine sets `<w:evenAndOddHeaders/>` in settings.xml when declared anywhere; clears the flag when dropped.

HF config is the source of truth for both flags — re-applying with a smaller variant set cleans up stale flags from prior runs.

v1 binds the same HF set to every section. Per-section variation is not yet supported.

## Combinations

- **HF + `edits[]`** — both share the body asset registry; image / hyperlink rIds allocate linearly across body and HF emissions. Registering the same `png` in both adds exactly one `<Default Extension="png">` entry to `[Content_Types].xml`.
- **HF + `trackChanges: true`** — track-changes wraps body edits only. HF mutation runs as a separate pass producing pristine `<w:hdr>` / `<w:ftr>` content (no `<w:ins>` / `<w:del>` wrappers). Use this combo when iterating on body content under change tracking but wanting header/footer rewritten cleanly.

## Not supported

- `ParagraphBlock.anchor` inside HF — bookmark semantics on a part rendered across multiple pages are ambiguous; the engine throws at emit. For page-number cross-references, use the `field: "page"` inline node — shapes in [`edit.md`](edit.md).
- Per-section HF overrides (every section gets the same HF set in v1).
- HF cleanup (orphan strategy leaves old parts in archive; future phase may GC).
