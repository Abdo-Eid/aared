/**
 * A table cell's `gridSpan`/`rowSpan` as the file states it, cut to the room
 * the table has left (the columns or rows from the cell to the table's end).
 * Both table formats read spans through this: a span past the edge covers
 * nothing more, and the renderers loop over a cell's span, so an unclamped
 * claim of millions would freeze the page. Returns `undefined` for no span and
 * never less than 1.
 */
export function clampSpan(span: number | undefined, room: number): number | undefined {
  if (span === undefined || !Number.isFinite(span)) return undefined;
  return Math.max(1, Math.min(Math.floor(span), Math.max(1, room)));
}
