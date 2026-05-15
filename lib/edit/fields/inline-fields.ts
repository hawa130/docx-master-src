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
  const { runs } = emitComplexField(ownerDoc, { instrCode: instr, initialResult: placeholder, format })
  return runs
}

export function emitInlineStyleRef(
  ownerDoc: Document,
  styleName: string,
  numberOnly: boolean,
  format: RunFormat | undefined,
): Element[] {
  // STYLEREF's argument for built-in heading styles is locale-sensitive
  // in Word: a doc whose styles.xml carries `<w:name w:val="heading 1"/>`
  // resolves as "标题 1" in a Chinese Word UI, "Überschrift 1" in
  // German, etc. — emitting `STYLEREF "heading 1"` fails the localized
  // built-in identity match and Word renders the "Error! Use Home tab
  // to apply ..." placeholder. The documented cross-locale form is
  // `STYLEREF N` (bare numeric outline level 1-9) — Word treats it as
  // "the heading at level N", works in every language version. Custom
  // style names aren't translated so the quoted form is fine there.
  const headingLevel = headingLevelFromName(styleName)
  const argument = headingLevel !== null ? String(headingLevel) : `"${styleName}"`
  const switches = numberOnly ? " \\n" : ""
  const { runs } = emitComplexField(ownerDoc, {
    instrCode: `STYLEREF ${argument}${switches}`,
    initialResult: "",
    format,
  })
  return runs
}

/** Returns 1-9 when `name` is a built-in heading display name in the
 *  `<w:name w:val="heading N"/>` form (case-insensitive, trims), null
 *  otherwise. */
function headingLevelFromName(name: string): number | null {
  const m = /^heading\s+([1-9])$/i.exec(name.trim())
  return m ? parseInt(m[1]!, 10) : null
}
