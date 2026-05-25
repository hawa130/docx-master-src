/**
 * Type system for the `edit` command — surgical edits on an existing docx.
 *
 * Design rules (preserve when extending):
 *   - All unions discriminate on `type` (Locator / Block) or `op` (EditOp).
 *     Adding a variant forces every consumer's switch to update — protected by
 *     `assertNever` defaults.
 *   - Format types (RunFormat / ParagraphFormat) are open interfaces with
 *     optional fields. New format properties = one new field; emitter switches
 *     pick them up automatically when the field is present.
 *   - This file declares only types. Zod runtime validation lives in
 *     edit-config-schema.ts (single source of truth for the JSON the agent
 *     writes); these TS types are inferred from those schemas, kept in sync
 *     by import.
 *
 * Extension points (where to add new variants):
 *   - new Locator kind    → edit-types.ts (add to union)
 *                         + edit-config-schema.ts (add to LocatorSchema)
 *                         + locator.ts (add resolver case)
 *   - new Block type      → edit-types.ts + edit-config-schema.ts + fragment-emit.ts
 *   - new RunFormat field → edit-types.ts + edit-config-schema.ts
 *                         + fragment-emit.ts (serialize)
 *                         + style-mutation.ts equivalent for format-op writes
 *   - new EditOp          → edit-types.ts + edit-config-schema.ts + edit-engine.ts
 */

import type * as z from "zod/mini"
import type { WIdAllocator } from "@lib/edit/wid-allocator.ts"
import type {
  BlockSchema,
  BorderEdgeSchema,
  EditConfigSchema,
  EditOpSchema,
  FragmentSchema,
  LocatorSchema,
  ParagraphFormatSchema,
  RichTextSchema,
  RunFormatSchema,
  RunLocatorSchema,
} from "@lib/config/edit-config-schema.ts"

/* ------------- Public config types (inferred from zod schemas) ------------- */

export type Locator = z.infer<typeof LocatorSchema>
export type RunLocator = z.infer<typeof RunLocatorSchema>
export type RunFormat = z.infer<typeof RunFormatSchema>
export type ParagraphFormat = z.infer<typeof ParagraphFormatSchema>
export type RichText = z.infer<typeof RichTextSchema>
export type Block = z.infer<typeof BlockSchema>
export type BorderEdge = z.infer<typeof BorderEdgeSchema>
export type Fragment = z.infer<typeof FragmentSchema>
export type EditOp = z.infer<typeof EditOpSchema>
export type EditConfig = z.infer<typeof EditConfigSchema>

/* ------------- Internal model (post-resolution) ------------- */

/**
 * After locator resolution, every edit op carries a concrete target —
 * a list of <w:p> Element references plus the container they live in.
 * Op execution switches on `op.op`, never on the original locator kind:
 * the resolver collapses all locator forms into this same shape.
 *
 * `paragraphs` may be empty for `whole-body` insert at the end (the
 * caller infers position from `container` + `containerEndIndex`); replace /
 * delete / format require at least one resolved paragraph and report an
 * error otherwise (caught earlier by the resolver).
 */
export interface ResolvedTarget {
  /** Matched <w:p> elements in document order. */
  paragraphs: Element[]
  /** Either the body element or a <w:tc> (for cell locators). Insertions
   * land here as new children. */
  container: Element
}

export interface ResolvedEdit {
  op: EditOp
  target: ResolvedTarget
  /** Set only for `set-run`: the resolved run inside the target paragraph.
   * Other ops leave this undefined. */
  runRef?: Element
}

/* ------------- Track-changes context ------------- */

/**
 * Wraps id allocation and the fixed metadata that every <w:ins> / <w:del> /
 * <w:rPrChange> / <w:pPrChange> needs. Author is intentionally empty (we
 * don't fabricate identity); date is fixed per-run for a deterministic
 * change-set fingerprint.
 *
 * `nextId` delegates to a `WIdAllocator` shared with the apply's
 * BookmarkAllocator — revision IDs share the document-wide `w:id` space
 * with bookmarks, comments, moves and pPrChange/rPrChange, so allocating
 * from a private counter collides with source IDs and with the
 * bookmark allocator.
 */
export interface TrackContext {
  enabled: boolean
  nextId(): number
  author: string
  date: string
}

export function makeTrackContext(
  enabled: boolean,
  idAllocator: WIdAllocator,
  options?: { author?: string; isoDate?: string },
): TrackContext {
  return {
    enabled,
    nextId: () => idAllocator.next(),
    // Empty string fallback: ECMA-376 §17.13 requires `w:author` to be
    // present on every revision element, but accepts empty value. We never
    // synthesize a tool-brand default — unattributed beats fabricated.
    author: options?.author ?? "",
    date: options?.isoDate ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z"),
  }
}

/* ------------- Exhaustiveness helper ------------- */

/**
 * Compile-time enforcement that every union variant is handled. Place in the
 * `default` of a switch over a discriminated union — TypeScript narrows the
 * residual type to `never`, so adding a new variant without updating the
 * switch becomes a type error here.
 */
export function assertNever(x: never): never {
  throw new Error(`unhandled variant: ${JSON.stringify(x)}`)
}
