/**
 * A small most-recently-used cache for the units a Worker session keeps
 * parsed (slides, sheets). Unbounded, a session that visits every slide of a
 * 300-slide deck kept every model — and each model's part text — for as long
 * as the document stayed open. Bounded, the units a viewer flips between stay
 * warm and an evicted one simply re-parses. DOM-free; its own module because
 * `worker.ts` has side effects on import.
 */

/** The viewed slide and the neighbours a viewer goes back to. */
export const PPTX_SLIDE_CACHE = 4;
/** The viewed sheet (range reads slice it) and one tab switch back. */
export const XLSX_SHEET_CACHE = 2;

export class UnitCache<V> {
  readonly #entries = new Map<number, V>();

  constructor(readonly capacity: number) {}

  /** A hit becomes the most recent. */
  get(key: number): V | undefined {
    const value = this.#entries.get(key);
    if (value !== undefined) {
      this.#entries.delete(key);
      this.#entries.set(key, value);
    }
    return value;
  }

  /** A hit that leaves recency alone (reads that never paint, like search). */
  peek(key: number): V | undefined {
    return this.#entries.get(key);
  }

  set(key: number, value: V): void {
    this.#entries.delete(key);
    this.#entries.set(key, value);
    while (this.#entries.size > this.capacity) this.#entries.delete(this.#entries.keys().next().value!);
  }

  get size(): number {
    return this.#entries.size;
  }
}
