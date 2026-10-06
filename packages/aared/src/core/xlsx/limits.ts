/**
 * Excel's worksheet grid: XFD1048576. Counts, so the last zero-based index of
 * each axis is one less. The parser, the A1 reference reader, the main-thread
 * range check and the Worker's window check all read these.
 */

export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;

/**
 * Merged cells the viewer indexes one entry each. Excel lets a merge cover the
 * whole sheet (17 billion cells); a preview needs only what the sheet shows, so
 * a merge is first cut to the sheet's extent and refused past these budgets.
 */
export const MAX_MERGED_CELLS = 250_000;
export const MAX_MERGED_CELLS_PER_SHEET = 1_000_000;

/** Whether a zero-based cell lies on the worksheet grid. */
export function isInsideGrid(row: number, col: number): boolean {
  return Number.isInteger(row) && Number.isInteger(col) && row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLUMNS;
}
