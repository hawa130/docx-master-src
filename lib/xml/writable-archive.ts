/**
 * Writer-bound replacement map with an enforced single-writer invariant
 * for the OOXML package's coordination parts.
 *
 * Background: `[Content_Types].xml` and `word/_rels/document.xml.rels`
 * are shared parts that multiple apply subsystems would naively want to
 * mutate — numbering registration, settings fabrication, header / footer
 * Override entries, image rels, hyperlink rels, etc. Letting each
 * subsystem read-modify-write the part text independently produces a
 * silent last-writer-wins race: subsystem A reads the source bytes,
 * appends entry X, writes the result back; subsystem B reads the SAME
 * source bytes, appends entry Y, writes the result back, **dropping X**.
 * This actually shipped (the chineseCounting numbering scheme being
 * silently dropped when the body asset registry's flushTo ran after
 * `ensureNumberingContentType` — discovered during the blank-source demo).
 *
 * The fix is structural: both parts have exactly one writer — the
 * accumulator owned by `bodyAssetRegistry` (`ContentTypes` for the
 * content-types file, body `PartRels` for the doc rels file). Every
 * subsystem that wants to register a new Default / Override / Relationship
 * routes through those accumulators; one final `flushTo` at end of apply
 * serialises the accumulated state.
 *
 * This wrapper enforces the invariant at the API boundary: a runtime
 * `throw` on any attempt to `.set()` either of the two coordination
 * paths from outside the accumulator. The accumulators themselves use
 * `setFromAccumulator()` (intentionally underscore-prefixed so grep for
 * `replacements.set(` doesn't false-positive on legitimate flushTo
 * calls).
 */

const FORBIDDEN_DIRECT_WRITES: ReadonlySet<string> = new Set([
  "[Content_Types].xml",
  "word/_rels/document.xml.rels",
])

export class WritableArchive {
  private inner = new Map<string, string | Uint8Array>()

  /** Stage `content` at `path` in the archive. Throws on the two
   *  coordination paths — those must go through the shared accumulators
   *  owned by `bodyAssetRegistry`. */
  set(path: string, content: string | Uint8Array): void {
    if (FORBIDDEN_DIRECT_WRITES.has(path)) {
      throw new Error(
        `WritableArchive.set: direct writes to "${path}" are forbidden — ` +
          `multiple subsystems otherwise race-overwrite each other. Route ` +
          `the entry through bodyAssetRegistry.getContentTypes() (for ` +
          `Override / Default entries) or bodyAssetRegistry.getPartRels() ` +
          `(for body Relationships). See lib/xml/writable-archive.ts.`,
      )
    }
    this.inner.set(path, content)
  }

  /** Escape hatch for the two legitimate writers of forbidden paths:
   *  `ContentTypes.flushTo` and `PartRels.flushTo` (when its target path
   *  is `word/_rels/document.xml.rels`). The underscore prefix flags
   *  this as internal so it stays grep-visible during reviews. */
  setFromAccumulator(path: string, content: string | Uint8Array): void {
    this.inner.set(path, content)
  }

  get(path: string): string | Uint8Array | undefined {
    return this.inner.get(path)
  }

  has(path: string): boolean {
    return this.inner.has(path)
  }

  get size(): number {
    return this.inner.size
  }

  /** Underlying Map handed to `DocxReader.copyAndModify`. Read-only by
   *  convention — callers should not mutate it directly. */
  toMap(): Map<string, string | Uint8Array> {
    return this.inner
  }

  [Symbol.iterator](): IterableIterator<[string, string | Uint8Array]> {
    return this.inner.entries()
  }
}
