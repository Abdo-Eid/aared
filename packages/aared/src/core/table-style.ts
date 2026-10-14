/**
 * Which conditional-formatting regions of a table style apply to one cell, in
 * the order they layer.
 *
 * PowerPoint (`a:tblStyle`) and Word (`w:tblStyle`) name the same thirteen
 * regions differently, but decide *which* region covers *which* cell by the
 * same rule, and that rule is what this module holds: the band arithmetic
 * especially, which excludes the header row from the band sequence (probe
 * `docx-table-style`: with `w:tblLook/@firstRow="1"` the painted band rows were
 * 2, 4, 6 - the first body row is band 1 - and with `firstRow="0"` they were
 * 1, 3, 5, 7). A vertical merge does not shift the count; the merged cell takes
 * its origin row's band.
 *
 * The two formats spell the regions differently (`wholeTbl`/`wholeTable`,
 * `band1H`/`band1Horz`, `band1V`/`band1Vert`), so the caller maps a returned
 * {@link TableStyleRegionName} onto its own spelling. The order is weakest
 * first, and the caller applies direct formatting after all of it.
 */

/** A region, named as `a:tblStyle` names it. `w:tblStylePr` maps onto these. */
export const TABLE_STYLE_REGIONS = [
  "wholeTbl", "band1H", "band2H", "band1V", "band2V",
  "lastCol", "firstCol", "lastRow", "firstRow",
  "seCell", "swCell", "neCell", "nwCell",
] as const;

export type TableStyleRegionName = typeof TABLE_STYLE_REGIONS[number];

/**
 * How a table's `w:tblLook` / `a:tbl` switches read. `bandRow` is "horizontal
 * banding is on", which is the absence of `noHBand`, not the presence of a flag.
 */
export interface TableStyleFlags {
  rows: number;
  columns: number;
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  bandRow: boolean;
  bandColumn: boolean;
}

/** One applicable region and the grid extent its own edges are drawn against. */
export interface TableStyleRegionUse {
  name: TableStyleRegionName;
  rows: readonly [number, number];
  columns: readonly [number, number];
}

/**
 * The regions covering the cell that spans rows `r0..r1` and columns `c0..c1`,
 * weakest first.
 *
 * An edge is the region's *outer* edge when the cell sits at that extent's
 * boundary and its *inside* edge otherwise, which is why the extent travels
 * with the region: a `band1Horz` fill spans the whole table, so the middle row
 * of that band draws `insideH` where a single-row region draws `top`.
 */
export function tableCellRegions(
  flags: TableStyleFlags,
  r0: number, r1: number, c0: number, c1: number,
): TableStyleRegionUse[] {
  const all: [number, number] = [0, flags.rows - 1];
  const allColumns: [number, number] = [0, flags.columns - 1];
  const regions: TableStyleRegionUse[] = [{ name: "wholeTbl", rows: all, columns: allColumns }];
  const isHeaderRow = flags.firstRow && r0 === 0;
  const isTotalRow = flags.lastRow && r1 === flags.rows - 1;
  const isFirstColumn = flags.firstColumn && c0 === 0;
  const isLastColumn = flags.lastColumn && c1 === flags.columns - 1;
  // A first/last row or column is not also a band: Word paints the header row's
  // own fill there, not a band tint over it.
  if (flags.bandColumn && !isFirstColumn && !isLastColumn) {
    const band = (c0 - (flags.firstColumn ? 1 : 0)) % 2 === 0 ? "band1V" : "band2V";
    regions.push({ name: band, rows: all, columns: [c0, c1] });
  }
  if (flags.bandRow && !isHeaderRow && !isTotalRow) {
    // The header row is excluded from the sequence, so the first body row is
    // band 1 (probe `docx-table-style`).
    const band = (r0 - (flags.firstRow ? 1 : 0)) % 2 === 0 ? "band1H" : "band2H";
    regions.push({ name: band, rows: [r0, r1], columns: allColumns });
  }
  if (isLastColumn) regions.push({ name: "lastCol", rows: all, columns: [flags.columns - 1, flags.columns - 1] });
  if (isFirstColumn) regions.push({ name: "firstCol", rows: all, columns: [0, 0] });
  if (isTotalRow) regions.push({ name: "lastRow", rows: [flags.rows - 1, flags.rows - 1], columns: allColumns });
  if (isHeaderRow) regions.push({ name: "firstRow", rows: [0, 0], columns: allColumns });
  const corner = (name: TableStyleRegionName, row: number, column: number): void => {
    regions.push({ name, rows: [row, row], columns: [column, column] });
  };
  if (isTotalRow && isLastColumn) corner("seCell", flags.rows - 1, flags.columns - 1);
  if (isTotalRow && isFirstColumn) corner("swCell", flags.rows - 1, 0);
  if (isHeaderRow && isLastColumn) corner("neCell", 0, flags.columns - 1);
  if (isHeaderRow && isFirstColumn) corner("nwCell", 0, 0);
  return regions;
}

/**
 * `w:tblLook`'s legacy `w:val` bitmask, for files that carry no explicit
 * attributes (Word writes both; `fixtures/docx-coverage.docx` writes only
 * `val="0000"`). ECMA-376 17.4.40.
 */
export function tableLookFromVal(val: string | undefined): {
  firstRow: boolean; lastRow: boolean; firstColumn: boolean; lastColumn: boolean;
  bandRow: boolean; bandColumn: boolean;
} | undefined {
  if (val === undefined || !/^[0-9A-Fa-f]{4}$/.test(val)) return undefined;
  const bits = Number.parseInt(val, 16);
  return {
    firstRow: (bits & 0x0020) !== 0,
    lastRow: (bits & 0x0040) !== 0,
    firstColumn: (bits & 0x0080) !== 0,
    lastColumn: (bits & 0x0100) !== 0,
    // The bit is "no banding", so the flag is its absence.
    bandRow: (bits & 0x0200) === 0,
    bandColumn: (bits & 0x0400) === 0,
  };
}