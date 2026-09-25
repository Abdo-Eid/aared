/**
 * Excel's worksheet grid: XFD1048576. Counts, so the last zero-based index of
 * each axis is one less. The parser, the A1 reference reader, the main-thread
 * range check and the Worker's window check all read these.
 */

export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;

/** Whether a zero-based cell lies on the worksheet grid. */
export function isInsideGrid(row: number, col: number): boolean {
  return Number.isInteger(row) && Number.isInteger(col) && row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLUMNS;
}
