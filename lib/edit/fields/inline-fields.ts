/**
 * Inline field nodes — thin wrappers over `emitComplexField`.
 *
 *   `{ field: "page"      }` → PAGE
 *   `{ field: "numPages"  }` → NUMPAGES
 *   `{ field: "date"      }` → DATE
 *   `{ styleRef: "..."    }` → STYLEREF "<name>"
 *   `{ styleRef: "...", numberOnly: true }` → STYLEREF "<name>" \n
 *
 * The placeholder text shown before Word's field-update pass is a sensible
 * default per field type — Word replaces it on next open thanks to the
 * `updateFields` flag set in settings.xml.
 *
 * STYLEREF requires the referenced styleId to exist when Word resolves the
 * field. Validation is the engine's job (apply reads stylesDoc and throws
 * on missing styleId); zod doesn't have that context.
 */

import { emitComplexField } from "@lib/edit/fields/complex-field.ts"
import type { RunFormat } from "@lib/config/edit-types.ts"
import type { StyleInfo } from "@lib/parse/style-names.ts"

type FieldKind = "page" | "numPages" | "date"

interface FieldDescriptor {
  instr: string
  placeholder: string
}

const FIELD_DESCRIPTORS: Record<FieldKind, FieldDescriptor> = {
  page: { instr: "PAGE", placeholder: "1" },
  numPages: { instr: "NUMPAGES", placeholder: "1" },
  // Empty placeholder for DATE — Word fills the current date on open;
  // typing a plausible-looking date here would mislead on read-without-update.
  date: { instr: "DATE", placeholder: "" },
}

export function emitInlineField(
  ownerDoc: Document,
  field: FieldKind,
  format: RunFormat | undefined,
): Element[] {
  const { instr, placeholder } = FIELD_DESCRIPTORS[field]
  const { runs } = emitComplexField(ownerDoc, {
    instrCode: instr,
    initialResult: placeholder,
    format,
  })
  return runs
}

export function emitInlineStyleRef(
  ownerDoc: Document,
  info: StyleInfo,
  numberOnly: boolean,
  format: RunFormat | undefined,
): Element[] {
  // Three-way dispatch driven by the resolved style's identity:
  //
  //   1. outline-bound (`<w:outlineLvl>` present) → emit `STYLEREF N`.
  //      Locale-neutral — Word resolves by outline level, not name.
  //      Built-in headings 1-9 + any custom style the agent bound to
  //      an outline level both land here.
  //
  //   2. custom (non-built-in) name → emit `STYLEREF "<name>"`.
  //      Custom style names aren't translated by Word's locale
  //      mapping; the quoted form matches reliably across all UI
  //      languages.
  //
  //   3. built-in localizable non-outline (Title / Caption / Subtitle /
  //      etc.) → throw. `STYLEREF "Title"` would silently fail in
  //      non-EN Word UIs (the UI shows "标题" / "Titre" / "Titel" and
  //      STYLEREF resolves against that localized form, not the
  //      stored English canonical name). No `STYLEREF N` equivalent
  //      exists for non-outline built-ins — agent must either bind
  //      the style to an outline level or rename to a custom name.
  const switches = numberOnly ? " \\n" : ""
  let argument: string
  if (info.outlineLevel !== undefined) {
    argument = String(info.outlineLevel)
  } else if (!info.isBuiltInLocalizable) {
    argument = `"${info.name}"`
  } else {
    throw new Error(
      `InlineStyleRef: cannot reference built-in style "${info.name}" via STYLEREF. ` +
        `Word translates built-in style names per UI locale (e.g. "Title" ↔ "标题", ` +
        `"Caption" ↔ "题注"), so STYLEREF "${info.name}" silently fails in non-English Word. ` +
        `Fix by either: ` +
        `(a) binding this style to an outline level (set styles[].outlineLevel = 0-8) so the ` +
        `engine can emit STYLEREF N (locale-neutral); or ` +
        `(b) renaming the style to a custom name (e.g. "Doc${info.name}") — custom names ` +
        `aren't translated.`,
    )
  }
  const { runs } = emitComplexField(ownerDoc, {
    instrCode: `STYLEREF ${argument}${switches}`,
    initialResult: "",
    format,
  })
  return runs
}
