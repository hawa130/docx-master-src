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
  styleRef: string,
  numberOnly: boolean,
  format: RunFormat | undefined,
): Element[] {
  // Per ECMA-376 17.16.5.61, STYLEREF's style-name argument is quoted with
  // straight ASCII double quotes. Internal " is not valid in OOXML styleIds,
  // so escaping isn't needed — strictObject's NonEmptyString and the apply-
  // time styleId-existence check catch anything weird earlier.
  const switches = numberOnly ? " \\n" : ""
  const { runs } = emitComplexField(ownerDoc, {
    instrCode: `STYLEREF "${styleRef}"${switches}`,
    initialResult: "",
    format,
  })
  return runs
}
