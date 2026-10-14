/**
 * The page half of the chart engine: lay a resolved `ChartSpecModel` out in
 * its frame and return the primitives to paint. `chart-svg.ts` loads this
 * module with a dynamic `import()` the first time a chart shows, so a file
 * with no charts never downloads the layout code, and the Worker (which only
 * reads charts, `core/chart.ts`) never contains it (Notes §71).
 */

import { PT_PER_INCH } from "../units.js";
import { drawChart } from "./engine.js";
import { arabicFamily, arabicScaleKey, hasArabicScript, setFontScales, textWidth } from "./metrics.js";
import type { ChartPrimitive } from "./scene.js";
import type { ChartSpecModel, TextLook } from "./spec.js";

/** Office's default inserted-chart size, for a chart whose frame has no size. */
const DEFAULT_FRAME_IN = { width: 5, height: 3 };

export interface ChartDrawing {
  /** The frame in points: the view box the primitives are drawn in. */
  width: number;
  height: number;
  primitives: ChartPrimitive[];
}

/**
 * Lay `chart` out in a frame of `frame` inches. Charts lay out differently at
 * different sizes, exactly as in Office (more room, more axis steps), so pass
 * the real frame.
 *
 * `fontScale` widens or narrows the measured text per face: Office's widths
 * are the default, and a reader without Office's faces paints a wider
 * substitute (see `chartFontProbes`). A key is a `TextLook.family` for Latin
 * text, or `arabicScaleKey(face)` for Arabic glyphs; a missing key is 1.
 */
export function layoutChart(
  chart: ChartSpecModel,
  frame: { width: number; height: number } | undefined,
  fontScale?: Readonly<Record<string, number>>,
): ChartDrawing {
  const size = frame !== undefined && frame.width > 0 && frame.height > 0 ? frame : DEFAULT_FRAME_IN;
  const width = Math.round(size.width * PT_PER_INCH * 100) / 100;
  const height = Math.round(size.height * PT_PER_INCH * 100) / 100;
  setFontScales(fontScale);
  try {
    return { width, height, primitives: drawChart(chart, width, height).primitives };
  } finally {
    setFontScales(undefined);
  }
}

/** One face to measure on the page: paint `text` in `family`, compare with `expected`, scale the layout by the ratio. */
export interface FontProbe {
  /** Where the scale goes in `layoutChart`'s `fontScale`. */
  key: string;
  family: string;
  text: string;
  /** What the engine's tables say `text` measures, in em. */
  expected: number;
}

const LATIN_PROBE = "Quarterly Revenue 0123456789";

/** Every text style in the chart, walking only objects (number arrays, the bulk of a big chart, are skipped at once). */
function textLooks(value: unknown, out: TextLook[], arabic: { seen: boolean }): void {
  if (value === null || typeof value !== "object") return;
  if (Array.isArray(value)) {
    const first = value[0];
    if (first === undefined || typeof first === "number" || first === null) return;
    if (typeof first === "string") {
      // Category and title lines: enough to know whether the chart carries Arabic.
      if (!arabic.seen) for (const item of value) if (typeof item === "string" && hasArabicScript(item)) { arabic.seen = true; break; }
      return;
    }
    for (const item of value) textLooks(item, out, arabic);
    return;
  }
  const record = value as Record<string, unknown>;
  if (typeof record.family === "string" && typeof record.size === "number") out.push(record as unknown as TextLook);
  for (const name in record) {
    const child = record[name];
    if (typeof child === "string") {
      if (!arabic.seen && name === "name" && hasArabicScript(child)) arabic.seen = true;
    } else if (child !== null && typeof child === "object") {
      textLooks(child, out, arabic);
    }
  }
}

/**
 * The faces `chart` is set in, as measurements for the page to take. The
 * layout runs on the engine's tables (Office's own widths), so on a device
 * without the face the chart is drawn in a wider or narrower substitute and
 * its labels no longer fit the boxes laid out for them. The painter measures
 * each probe's `text` once with the browser's own text shaping (DOM, so not
 * here), divides by `expected`, and passes the ratios to `layoutChart`.
 * Arabic is probed only when the chart has Arabic text, so a Latin chart does
 * not pull in an Arabic face; `arabicText` is the sample it paints (the
 * caller's `ARABIC_PROBE`: this module imports nothing the page entry also
 * imports, or a bundler can put that shared code in the entry chunk and an
 * entry with top-level await then never finishes loading this one).
 */
export function chartFontProbes(chart: ChartSpecModel, arabicText: string): FontProbe[] {
  const looks: TextLook[] = [];
  const arabic = { seen: false };
  textLooks(chart, looks, arabic);
  const probes = new Map<string, FontProbe>();
  for (const look of looks) {
    if (!probes.has(look.family)) {
      probes.set(look.family, {
        key: look.family, family: look.family, text: LATIN_PROBE,
        expected: textWidth(LATIN_PROBE, { ...look, size: 1, bold: false, italic: false }),
      });
    }
    if (arabic.seen) {
      const face = arabicFamily(look);
      const key = arabicScaleKey(face);
      if (!probes.has(key)) {
        probes.set(key, {
          key, family: face, text: arabicText,
          expected: textWidth(arabicText, { ...look, family: face, csFamily: face, size: 1, bold: false, italic: false }),
        });
      }
    }
  }
  return [...probes.values()];
}
