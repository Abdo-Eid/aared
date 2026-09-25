/**
 * Unit constants and conversions shared by every parser and renderer.
 *
 * The resolved model speaks inches for geometry and points for type, Office's
 * own conventions. The packages store EMU (DrawingML), twips (WordprocessingML)
 * and character widths (SpreadsheetML); renderers paint CSS pixels at the
 * 96-dpi baseline PowerPoint's PNG export uses. Pure numbers plus one string
 * formatter, no DOM, so the Worker and the main thread share one definition —
 * a leaf module every other module may import.
 */

/** CSS pixels per inch: PowerPoint's PNG export and the layout baseline share it. */
export const PX_PER_INCH = 96;
export const PT_PER_INCH = 72;
export const EMU_PER_INCH = 914400;
export const EMU_PER_PT = EMU_PER_INCH / PT_PER_INCH;
export const TWIPS_PER_INCH = 1440;
/** A twip is a twentieth of a point. */
export const TWIPS_PER_PT = 20;

export function emuToInches(emu: number): number {
  return emu / EMU_PER_INCH;
}

export function inchesToEmu(inches: number): number {
  return inches * EMU_PER_INCH;
}

export function twipsToInches(twips: number): number {
  return twips / TWIPS_PER_INCH;
}

export function ptToInches(points: number): number {
  return points / PT_PER_INCH;
}

export function inchesToPt(inches: number): number {
  return inches * PT_PER_INCH;
}

export function emuToPt(emu: number): number {
  return inchesToPt(emuToInches(emu));
}

export function ptToPx(points: number): number {
  return (points / PT_PER_INCH) * PX_PER_INCH;
}

export function twipsToPx(twips: number): number {
  return (twips / TWIPS_PER_INCH) * PX_PER_INCH;
}

/** A model distance in inches as a CSS pixel length. */
export function inchesToCss(inches: number): string {
  return `${inches * PX_PER_INCH}px`;
}
