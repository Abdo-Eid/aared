/**
 * Which words go in which slot, when a line is split around a picture. Pure
 * numbers, no DOM: `split-wrap.ts` measures the words and paints the result.
 */

/** One place on a line where text may go, in reading order. */
export interface Slot {
  /** Width in pixels. */
  widthPx: number;
  /**
   * A slot that must take a word even when the word is wider than the slot
   * (a full-width slot: the alternative is a line that never advances).
   */
  force?: boolean;
}

/** Slots narrower than this many pixels take no text: not even one short word fits sensibly. */
export const MIN_SLOT_PX = 24;

/**
 * Fill `slots` in order from `widths` (each entry one unbreakable word unit),
 * greedily, a space between neighbours in the same slot. Returns how many
 * units each slot took. A slot too narrow for the next unit stays empty when a
 * later slot may take it; only a `force` slot takes an over-wide unit.
 */
export function fitUnits(widths: readonly number[], spacePx: number, slots: readonly Slot[], start = 0): number[] {
  const taken: number[] = [];
  let next = start;
  for (const slot of slots) {
    let count = 0;
    let used = 0;
    if (slot.widthPx >= MIN_SLOT_PX) {
      while (next < widths.length) {
        const width = widths[next]!;
        const needed = count === 0 ? width : used + spacePx + width;
        if (needed > slot.widthPx && !(count === 0 && slot.force === true)) break;
        used = needed;
        count += 1;
        next += 1;
      }
    }
    taken.push(count);
  }
  return taken;
}
