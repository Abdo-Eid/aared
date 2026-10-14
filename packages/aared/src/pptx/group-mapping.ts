import type { BoundsModel } from "../core/model.js";
import { PX_PER_INCH } from "../core/units.js";

/**
 * The CSS mapping from a group's authored child space (`chOff`/`chExt`) onto
 * its placed bounds: translate the child origin out, then scale. Missing or
 * degenerate bounds/child space map to `""` (identity) rather than a guessed
 * scale. Pure so the scale math is unit-testable without a DOM. Internal —
 * not part of the `@aared/viewer/pptx` public surface.
 */
export function groupChildMapping(bounds: BoundsModel | undefined, childSpace: BoundsModel | undefined): string {
  if (bounds === undefined || childSpace === undefined) return "";
  if (bounds.width <= 0 || bounds.height <= 0 || childSpace.width <= 0 || childSpace.height <= 0) return "";
  const sx = bounds.width / childSpace.width;
  const sy = bounds.height / childSpace.height;
  const tx = -childSpace.x * sx * PX_PER_INCH;
  const ty = -childSpace.y * sy * PX_PER_INCH;
  return `translate(${tx}px, ${ty}px) scale(${sx}, ${sy})`;
}
