# `audit`

When the user wants to check whether a document conforms to a spec **without modifying it**. Read-only — produces a violation report, not a new docx.

*Illustrative phrasings: "看看这份合不合学校规范", "对照这个标准检查一下". Same surface words can land in `apply` instead when the user actually wants the issues fixed; ask one focused question if intent is genuinely ambiguous.*

## Workflow

1. **Read the spec.** What rules need verifying? Body font / size / line spacing? Heading hierarchy and outline levels? Image-caption format? Page margins? Reference style?

2. **Gather facts.** `overview` for the global picture; `inspect_style`, `inspect_style_def`, `inspect_section`, `find_paragraphs` for specifics. Same inspect tools `apply` uses — just no write.

3. **Compare and report.** Produce a structured violation list:
   - What the spec required
   - What the document actually has
   - Which paragraphs / sections are affected (use letter labels and `#NNN` indices so the user can navigate to issues directly)

   If the user wants to fix the violations after seeing the audit, that becomes a separate `apply` call — don't auto-fix without permission.

## Tip

If the spec is long (a 20-page school formatting standard), don't try to audit every clause in one pass. Group violations by area (cover page / headings / body / figures / references) and let the user prioritize.
