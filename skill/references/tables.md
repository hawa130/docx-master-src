# Tables (`{ "type": "table", ... }`)

A Block type used inside `edits[]` to insert a new table. For simple data tables a two-line declaration is enough; complex layouts (merged cells, custom borders, mixed cell content) build progressively from the same schema.

## Quick start

```jsonc
{ "type": "table",
  "headerRows": 1,
  "rows": [
    ["Model", "Accuracy", "F1"],
    ["LLM-A", "0.87",     "0.85"],
    ["LLM-B", "0.84",     "0.82"]
  ]}
```

Default: `borders: "all"` (thin black grid), `alignment` omitted (left), `layout: "autofit"`. `headerRows: 1` makes the first row repeat as a header on page breaks — does NOT auto-bold the text; bind a styled paragraph via `headerStyle` (below) or format header cells explicitly.

## Schema reference

```ts
TableBlock {
  type: "table"
  rows: TableCell[][]
  headerRows?: number          // top N rows get <w:tblHeader/>; default 0
  headerStyle?: string         // styleId applied to cell paragraphs in header rows (default; explicit Block[] styleId wins)
  cols?: { width: TableWidth }[]   // length = effective column count; if omitted, engine auto-generates <w:gridCol w:w="0"/>
  borders?: BordersPreset | BordersCustom    // default "all"
  alignment?: "left" | "center" | "right"    // whole-table horizontal alignment on page
  vAlign?: "top" | "center" | "bottom"       // default cell vertical alignment; skill default "center"
  layout?: "fixed" | "autofit"               // default "autofit"
}

TableWidth = "auto" | <number, pt>

TableCell = string                                  // plain text, single paragraph
          | InlineNode[]                            // formatted text / cross-refs, single paragraph
          | Block[]                                 // multi-paragraph / images
          | { content: <string | InlineNode[] | Block[]>,
              colspan?: number,
              rowspan?: number,
              vAlign?: "top" | "center" | "bottom",
              borders?: BordersCustom,
              shading?: string                      // hex RGB
            }

BordersPreset = "all" | "none" | "outer" | "three-line"
BordersCustom = { top?, bottom?, left?, right?, insideH?, insideV?: BorderEdge }
BorderEdge    = "none" | "single" | "thick" | "double" | "dotted" | "dashed"
              | { style, size?: pt, color?: hex | "auto" }
```

## Cell content — four progressive forms

| Form | When to reach for it |
|---|---|
| `"text"` (plain string) | The 90% case. Single paragraph, no formatting, no spans. |
| `InlineNode[]` | Single paragraph with mixed run formatting (`{ text, format }`) or inline cross-refs (`{ refTo, display }`). |
| `Block[]` | Multiple paragraphs in one cell, images inside cells, or cell paragraphs needing their own `styleId`. |
| `{ content: ..., colspan, rowspan, vAlign, borders, shading }` | Anything above wrapped in an object that also carries spans or per-cell properties. |

Empty content (`""`, `[]`, or `{ content: [] }`) emits an empty `<w:p/>` — Word requires every cell to contain at least one block-level child.

## Cell merging

The agent declares only the "restart" cell (the one with content). The engine computes a grid occupancy map and injects vMerge continuation cells in subsequent rows at the same column. Continuation cells inherit gridSpan from the restart cell when both axes merge.

```jsonc
"rows": [
  [
    { "content": "Day", "rowspan": 2, "vAlign": "center" },
    { "content": "Performance", "colspan": 3 }
  ],
  ["Accuracy", "F1", "Latency"],
  [
    { "content": "Mon", "rowspan": 3 },
    "0.87", "0.85", "120ms"
  ],
  ["0.88", "0.86", "118ms"],
  ["0.86", "0.84", "122ms"]
]
```

Row 2 omits a cell at column 0 — claimed by `Day` from row 1. Rows 4 and 5 omit a cell at column 0 — claimed by `Mon`'s `rowspan: 3` from row 3.

Each row must total: `declared cells + cells claimed by ongoing rowspans + colspan-expanded width` = the table's column count. The engine throws with a clear error if a row's effective width disagrees.

## Borders

### Presets

| Preset | Effect |
|---|---|
| `"all"` | Thin single border on every edge (table-level + inside). Default. |
| `"none"` | All edges suppressed. |
| `"outer"` | Top / bottom / left / right thin; no inside lines. |
| `"three-line"` | Top thick + bottom thick + thin line under last header row. Sides + inside suppressed. Requires `headerRows ≥ 1` for the middle line; degrades silently to "top + bottom only" when `headerRows: 0`. |

### Custom

```jsonc
"borders": {
  "top":    { "style": "single", "size": 1.5 },
  "bottom": { "style": "single", "size": 1.5 },
  "insideH": "single",
  "left": "none",
  "right": "none",
  "insideV": "none"
}
```

Field `size` is in pt; engine converts to OOXML's 1/8 pt units (`w:sz`). Omitting a side defaults it to `"none"`. `"thick"` is a 1.5 pt shortcut.

### Cell-level override (escape hatch)

A cell in object form may carry its own `borders` overriding the table-level for that cell only. Used for "this row has a thick top line" or manual three-line construction beyond what the preset gives. Footgun: **adjacent cells do not auto-coordinate.** Setting cell A's right border to `"none"` does NOT remove the visible line between A and B — B is still drawing its own left border per the table-level cascade. To remove the line you must set both sides to `"none"`.

The `"three-line"` preset handles this coordination internally (engine injects per-cell `tcBorders` on the last header row). Outside of preset paths, agents writing manual coordination must do it explicitly.

## Headers

`headerRows: N` marks the top N rows as table headers — emits `<w:tblHeader/>` on each, making them repeat at the top of every page the table spans. Two important non-effects:

- **No auto-bold.** Header text renders identically to body cells unless agent specifies formatting.
- **No styling cascade.** Header rows don't get any default cell shading or border treatment beyond what the table-wide borders specify (except `"three-line"`'s middle line).

Bind a paragraph style to all header-row cells with `headerStyle`:

```jsonc
{ "type": "table",
  "headerRows": 1,
  "headerStyle": "TableHeader",
  "rows": [
    ["Header A", "Header B"],
    ["data",     "data"]
  ]}
```

`TableHeader` should be installed in `styles[]` with the desired font / weight / alignment. Cells with explicit `Block[]` containing their own `styleId` win over `headerStyle`.

## Column widths and layout

`cols` declares per-column widths; length must equal the effective column count (declared cells + rowspan claims, expanded by colspans). When omitted, engine auto-generates `<w:gridCol w:w="0"/>` × N — Word treats this as autofit.

```jsonc
"cols": [{ "width": "auto" }, { "width": 100 }, { "width": 200 }]
```

`layout` interpretation:
- `"autofit"` (default) — Word reflows column widths to fit content and page.
- `"fixed"` — Declared widths are honored even if total exceeds page width; content may overflow horizontally. Pair with explicit pt widths for predictable column sizing.

Percentage widths (e.g. `{ pct: 30 }`) are deferred out of v1 — they require coordinated `tblW` + per-cell `tcW` emission with OOXML's fiftiethPercent units. Use fixed pt widths instead.

## Alignment axes

Three distinct alignment fields apply at different scopes:

| Field | Scope | Affects |
|---|---|---|
| `alignment` on TableBlock | whole table on page | Horizontal position of the table (left / center / right). Academic / formal documents typically center. |
| `vAlign` on TableBlock (table-level) | every cell that doesn't carry its own `vAlign` | Skill default `"center"` — matches the academic / formal typography norm. Set `"top"` for form-style layouts where labels should hug the top of each cell. |
| `vAlign` on cell object form | within one cell | Overrides the table-level default for that cell only. |
| `paraFormat.alignment` on a Paragraph block inside a cell | within paragraph | Horizontal text alignment within the cell — separate from `vAlign`. |

Resolution for any given cell: `cell.vAlign ?? block.vAlign ?? "center"`. The engine ALWAYS emits `<w:vAlign>` on each `<w:tc>`, overriding Word's native top default. To restore Word's top-default behavior on the whole table, set `vAlign: "top"` at the table level.

## Edge cases the engine handles

- **Trailing paragraph after a table.** Word rejects a `<w:tbl>` as the last child of a body or cell. Engine appends `<w:p/>` after a trailing table in either container.
- **Adjacent tables.** Two `<w:tbl>` siblings need a `<w:p/>` between. Engine inserts one.
- **Empty cell.** `""` / `[]` / `{ content: [] }` becomes `<w:p/>`.
- **Mismatched row widths.** Engine throws at emit time with the row index and column counts.
- **Out-of-bounds spans.** `rowspan` past the last row or `colspan` past the row's remaining width throws at emit time.

## Caption placement

Place `TableCaption`-bound paragraphs **above** the table (mirror of figures which place captions below). Cross-refs from body text anchor on the caption paragraph, never on the table itself:

```jsonc
{ "type": "paragraph", "styleId": "TableCaption", "anchor": "tab-results", "text": "评测结果汇总" },
{ "type": "table", "headerRows": 1, "rows": [...] }
```

Then in body text:
```jsonc
[{ "text": "见 " }, { "refTo": { "type": "anchor", "name": "tab-results" }, "display": "label" }]
```

## Integration with the style system

| What | How |
|---|---|
| Cell paragraph styleId | `Block[]` form: `[{ "type": "paragraph", "styleId": "TableCellBody", "text": "..." }]`. Cascade resolves through styles.xml normally. |
| Auto-numbered captions inside cells | A cell paragraph bound to `FigureCaption` / `TableCaption` participates in the doc-wide counter — useful for image-gallery cells. |
| Theme fonts inside cells | rPr cascade applies; no special wiring needed. |
| Bookmark anchors inside cells | Cell paragraph blocks can declare `anchor` — same BookmarkAllocator path as body paragraphs. Cross-refs from any position resolve. |
| `pattern_rules` / `bulk_rules` coverage | **Data tables: NO.** Cell paragraphs aren't indexed by the parser. **Layout tables: YES.** Bind cell paragraphs to explicit styleId via `Block[]` when you need predictable styling, instead of relying on post-emit rules. |
| MDF (match destination formatting) | TableBlock is non-paragraph; engine doesn't propagate anchor pPr into it. The cells' own pPr is what you declared. |
| Table-level styles (Word's `<w:tblStyle>`) | **Not supported.** Reuse visual specs across many tables via JSON-fragment reuse in your config, not via styles.xml. |

## What's not supported (v1)

- **Nested tables via cell `Block[]`** — schema is acyclic. To put a table inside an existing cell, run a separate `apply` with a `cell` locator + insert op.
- **Deleting / replacing an existing table** — no locator selects a `<w:tbl>` directly. `paragraph` / `range` / `cell` / `heading` all target paragraphs.
- **trackChanges on TableBlock insertion** — engine throws at emit; OOXML has no clean "table inserted" tracked-change wrapper. Run table insertion in a separate apply without trackChanges, then enable trackChanges for cell-content edits afterward.
- **Table-level styles** (`<w:tblStyle>` reference + conditional region formatting like banded rows / firstRow auto-bold). OOXML's conditional regions are a separate sub-spec; this skill expects per-`TableBlock` declarations of borders / shading / alignment / headerStyle. Reuse across tables via JSON-fragment reuse.
- **Cell margins (`tcMar`), text rotation (`textDirection`), row no-break (`cantSplit`), zebra striping** — fall outside v1 scope. textDirection workaround: rephrase headers horizontally. cantSplit workaround: user fixes in Word after open.
- **Word built-in styles like "Grid Table 4 - Accent 1"** — not referenced; declare equivalent borders + shading manually.
- **Same-apply edit of a freshly inserted table's cells** — paragraph indices and cell locators reference pre-edit state. Insert the table in one apply, edit cells in a second.
- **Hyperlinks in cells** — hyperlink blocks are out of scope for this skill globally; not specific to tables.
