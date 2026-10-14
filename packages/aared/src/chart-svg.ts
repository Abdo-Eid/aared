/**
 * Paints charts on the page, for PPTX, XLSX and DOCX alike: one painter, or the
 * same chart would draw differently per format.
 *
 * The unit model carries the resolved chart (`ChartSpecModel`, read in the
 * Worker); laying it out is the bigger half of the engine, so it is a separate
 * file (`core/chart/draw.ts`) fetched the first time a chart shows. A file
 * with no charts never downloads it (Notes §71). The same holds for the
 * fill painter (`fill-paint.ts`): it is fetched only when a chart has a
 * gradient, pattern or picture fill, so a chart of plain colors never pays for
 * it.
 *
 * Every painted chart is an image to assistive technology: `role="img"`, the
 * author's alt text (else the chart's title, else "Chart") as its name, and a
 * `<desc>` listing the plot type, series and categories.
 */

import { cssClass } from "./class-sheet.js";
import { resourceId } from "./resource-id.js";
import { ARABIC_PROBE } from "./core/chars.js";
import { MARK_FALLBACKS } from "./diagnostics.js";
import type { ChartDrawing, FontProbe } from "./core/chart/draw.js";
import type { ChartPathPrimitive, ChartPrimitive } from "./core/chart/scene.js";
import { chartDescription, chartTitleText } from "./core/chart/text.js";
import type { ChartSpecModel, Fill } from "./core/chart/spec.js";
import { cssFontFamily } from "./core/typography.js";
import type { MediaUrls } from "./media-url.js";
import { loadFillPainter, loadedFillPainter, type FillPainter } from "./fill-loader.js";

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";

/**
 * The layout places every label for left-to-right anchoring, as Office does
 * whatever the language; `text-anchor` follows the inherited direction, so a
 * chart in a right-to-left slide, sheet or host page would flip its labels to
 * the other side of their anchors. Pinned here; Arabic text still runs right
 * to left inside its own label. Office lets a label overhang the chart area
 * slightly: the frame clips, not the SVG.
 */
const CHART_SVG = cssClass("direction:ltr;overflow:visible");

interface Layout {
  layoutChart: (spec: ChartSpecModel, frame: { width: number; height: number } | undefined, fontScale?: Record<string, number>) => ChartDrawing;
  chartFontProbes: (spec: ChartSpecModel, arabicText: string) => FontProbe[];
}

let layout: Layout | undefined;
let loading: Promise<Layout> | undefined;
/** Paints still waiting on a file or on a picture, for `chartsSettled()`. */
const pending = new Set<Promise<unknown>>();

function loadLayout(): Promise<Layout> {
  loading ??= import("./core/chart/draw.js").then(
    (module) => (layout = { layoutChart: module.layoutChart, chartFontProbes: module.chartFontProbes }),
    (error: unknown) => {
      // Offline, or the file was not deployed: a later chart tries again.
      loading = undefined;
      throw error;
    },
  );
  return loading;
}

function track(work: Promise<unknown>): void {
  const settled = work.then(() => undefined, () => undefined);
  pending.add(settled);
  void settled.then(() => pending.delete(settled));
}

/** What the painter needs beyond the primitive: the fill painter once loaded, and the media a picture fill names. */
export interface ChartPaintContext {
  fills?: FillPainter | undefined;
  urls?: MediaUrls | undefined;
}

/** A path whose fill is a paint server: anything but a plain color. */
type ServerPath = ChartPathPrimitive & { fill: Exclude<Fill, string> };

function hasPaintServer(primitive: ChartPrimitive): primitive is ServerPath {
  return primitive.kind === "path" && typeof primitive.fill === "object";
}

/**
 * A chart fill the SVG cannot take as a color, through the shared fill painter,
 * in the box of the shape it paints. Returns the `fill` attribute value, or
 * `undefined` when the painter cannot paint it (an unreadable gradient,
 * pattern or picture; the picture has no resolver): the caller marks the
 * shape and leaves it unfilled, never in a substitute color.
 */
function paintServerFill(svg: SVGSVGElement, primitive: ServerPath, fills: FillPainter, urls: MediaUrls | undefined): string | undefined {
  const box = primitive.box!;
  // The view box is the frame in points and the SVG is drawn at the frame's size in
  // CSS pixels, so a point is 96/72 of a pixel. Patterns tile from the chart's origin.
  const unitsPerPx = 72 / 96;
  return fills.fillPaint(svg, resourceId("chart-fill"), primitive.fill, {
    width: box.width,
    height: box.height,
    unitsPerPxX: unitsPerPx,
    unitsPerPxY: unitsPerPx,
    slideX: 0,
    slideY: 0,
    originX: box.x,
    originY: box.y,
  }, { urls });
}

/** Append one primitive to `svg`. An unfilled path stays unfilled (SVG's own default fill is black). */
export function appendChartPrimitive(svg: SVGSVGElement, primitive: ChartPrimitive, context: ChartPaintContext = {}): void {
  const node = document.createElementNS(SVG_NAMESPACE, primitive.kind);
  const set = (name: string, value: string | number | undefined): void => {
    if (value !== undefined) node.setAttribute(name, String(value));
  };
  if (primitive.kind === "path") {
    set("d", primitive.path);
    const fill = primitive.fill;
    let paint = "none";
    let marked = false;
    if (typeof fill === "string") paint = fill;
    else if (hasPaintServer(primitive) && primitive.box !== undefined && primitive.box.width > 0 && primitive.box.height > 0) {
      const fills = context.fills;
      const server = fills === undefined ? undefined : paintServerFill(svg, primitive, fills, context.urls);
      if (server === undefined) {
        // The shape's outline is right but its fill cannot be drawn faithfully:
        // it stays unfilled and says why. Development builds also dash the outline.
        marked = true;
        const reason = fills === undefined ? "painter-unavailable" : fills.fillFallbackReason(primitive.fill, context.urls);
        node.setAttribute("data-aared-fill-fallback", reason);
        svg.dataset.aaredFillFallback = reason;
      } else {
        paint = server;
      }
    }
    set("fill", paint);
    const stroke = primitive.stroke;
    if (stroke !== undefined) {
      set("stroke", stroke.color);
      set("stroke-width", stroke.width);
      if (stroke.dash !== undefined) set("stroke-dasharray", stroke.dash.join(" "));
      set("stroke-linecap", stroke.cap);
      set("stroke-linejoin", stroke.join);
    } else if (marked && MARK_FALLBACKS) {
      set("stroke", "#9CA3AF");
      set("stroke-width", 1);
      set("stroke-dasharray", "4 3");
    }
  } else {
    set("x", primitive.x);
    set("y", primitive.y);
    set("font-size", primitive.fontSize);
    // The same stack every run paints with: the family, a delivered
    // substitute's alias, then a generic. A bare family the reader lacks
    // (Aptos, the default theme's) fell to the SVG default serif.
    set("font-family", cssFontFamily(primitive.fontFamily));
    if (primitive.runs === undefined) {
      node.textContent = primitive.text;
    } else {
      // Arabic mixed with Latin: each part in its own face. A face that has both scripts' glyphs (Calibri
      // and Arial both do) would paint the Arabic in the Latin face from a font list alone, where Office
      // paints it in the complex-script face; tspans of one text element still lay out as one bidi line.
      for (const run of primitive.runs) {
        const part = document.createElementNS(SVG_NAMESPACE, "tspan");
        part.setAttribute("font-family", cssFontFamily(run.fontFamily));
        part.textContent = run.text;
        node.append(part);
      }
    }
    set("font-weight", primitive.fontWeight);
    set("font-style", primitive.fontStyle);
    if (primitive.direction === "rtl") {
      // Under direction=rtl `start` is the right edge and `end` the left. The layout anchors in
      // left-to-right terms, so swap them: the text stays where the engine placed it, and only the
      // order of its own neutral characters (digits, punctuation) follows the right-to-left base.
      set("direction", "rtl");
      set("text-anchor", primitive.textAnchor === "start" ? "end" : primitive.textAnchor === "end" ? "start" : "middle");
    } else {
      set("text-anchor", primitive.textAnchor);
    }
    set("fill", primitive.fill);
    if (primitive.rotate !== undefined) set("transform", `rotate(${primitive.rotate} ${primitive.x} ${primitive.y})`);
  }
  svg.append(node);
}

// ---------------------------------------------------------------- painted-font measurement

/**
 * The layout runs on Office's advance widths, so on a device without Office's
 * faces the browser paints a wider or narrower substitute and labels overflow
 * the boxes laid out for them. Each face a chart uses is measured once per
 * page: a hidden SVG `<text>` of a reference string, `getComputedTextLength`,
 * against the table's width for the same string. The ratio goes into the layout.
 * DOM work, so it lives here and not in core.
 */
const PROBE_SIZE = 100;
/** A face within this of its table is the face the table describes: browser kerning alone moves a width this far. */
const PROBE_TOLERANCE = 0.025;
const FONT_WAIT_MS = 2000;
const probeScales = new Map<string, number>();

function fontsReady(): Promise<void> {
  const fonts = typeof document === "undefined" ? undefined : document.fonts;
  if (fonts === undefined || fonts.ready === undefined) return Promise.resolve();
  return Promise.race([
    fonts.ready.then(() => undefined, () => undefined),
    new Promise<void>((resolve) => setTimeout(resolve, FONT_WAIT_MS)),
  ]);
}

/**
 * Measure `probes` and remember each ratio. Returns false, remembering nothing, when the page's
 * fonts are still loading (what was measured may be a fallback) unless `force` is set. A probe the
 * browser cannot measure (no layout, no glyphs) is left unscaled.
 */
function measureProbes(probes: readonly FontProbe[], force: boolean): boolean {
  const host = typeof document === "undefined" ? null : document.body;
  if (host === null || host === undefined) return true;
  const probe = document.createElementNS(SVG_NAMESPACE, "svg");
  probe.setAttribute("width", "10");
  probe.setAttribute("height", "10");
  probe.setAttribute("aria-hidden", "true");
  probe.style.cssText = "position:absolute;left:-9999px;top:0;visibility:hidden;pointer-events:none;direction:ltr";
  const nodes = probes.map((entry) => {
    const node = document.createElementNS(SVG_NAMESPACE, "text");
    node.setAttribute("font-size", String(PROBE_SIZE));
    node.setAttribute("font-family", cssFontFamily(entry.family));
    node.textContent = entry.text;
    probe.append(node);
    return node;
  });
  host.append(probe);
  const widths = nodes.map((node) => (typeof node.getComputedTextLength === "function" ? node.getComputedTextLength() : 0));
  probe.remove();
  if (!force && document.fonts !== undefined && document.fonts.status === "loading") return false;
  probes.forEach((entry, i) => {
    const ratio = widths[i]! / (PROBE_SIZE * entry.expected);
    if (!Number.isFinite(ratio) || ratio <= 0) return;
    probeScales.set(entry.key, Math.abs(ratio - 1) < PROBE_TOLERANCE ? 1 : Math.min(Math.max(ratio, 0.5), 2));
  });
  return true;
}

/** The measured scales for `probes` that differ from 1, or undefined when none does. */
function scalesFor(probes: readonly FontProbe[]): Record<string, number> | undefined {
  let scales: Record<string, number> | undefined;
  for (const entry of probes) {
    const scale = probeScales.get(entry.key);
    if (scale !== undefined && scale !== 1) (scales ??= {})[entry.key] = scale;
  }
  return scales;
}

// ---------------------------------------------------------------- painting

/** What `paintChart` is told about where the chart sits and how it is reached. */
export interface PaintChartOptions {
  /** The frame in inches; absent lays out at Office's default inserted-chart size. */
  frame: { width: number; height: number } | undefined;
  /** Media for picture fills; absent when the embedder gave no media resolver (such a fill is marked). */
  urls?: MediaUrls | undefined;
  /** The author's alt text, the chart's accessible name; absent falls back to the chart title, then "Chart". */
  alt?: string | undefined;
  /** Runs instead of painting when the layout code cannot be fetched. */
  unavailable: () => void;
}

/** Name and describe the chart to assistive technology; independent of layout, so it holds even when painting fails. */
function describeChart(svg: SVGSVGElement, spec: ChartSpecModel, alt: string | undefined): void {
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", alt !== undefined && alt.trim() !== "" ? alt : chartTitleText(spec) ?? "Chart");
  const id = resourceId("chart-desc");
  const desc = document.createElementNS(SVG_NAMESPACE, "desc");
  desc.setAttribute("id", id);
  desc.textContent = chartDescription(spec);
  svg.prepend(desc);
  svg.setAttribute("aria-describedby", id);
}

/**
 * Lay `spec` out at the frame and paint it into `svg`, whose view box becomes
 * the frame in points so the drawing maps 1:1 onto the box the caller sized.
 * Paints now when the layout code is loaded (and the fill painter too, when
 * the chart needs one), else when it arrives; `unavailable` runs instead when
 * a file cannot be fetched. `render()` waits for either through
 * `chartsSettled()`.
 */
export function paintChart(svg: SVGSVGElement, spec: ChartSpecModel, options: PaintChartOptions): void {
  svg.setAttribute("preserveAspectRatio", "none");
  svg.classList.add(CHART_SVG);
  describeChart(svg, spec, options.alt);
  const context: ChartPaintContext = { urls: options.urls };
  const paint = (drawing: ChartDrawing, fills: FillPainter | undefined): void => {
    svg.setAttribute("viewBox", `0 0 ${drawing.width} ${drawing.height}`);
    context.fills = fills;
    for (const primitive of drawing.primitives) appendChartPrimitive(svg, primitive, context);
    svg.dataset.aaredChartPrimitives = String(drawing.primitives.length);
  };
  // Lay out with the measured faces, then paint, fetching the fill painter first when a fill needs it.
  const layOut = (lay: Layout, probes: readonly FontProbe[]): void | Promise<void> => {
    const drawing = lay.layoutChart(spec, options.frame, scalesFor(probes));
    if (!drawing.primitives.some(hasPaintServer)) return paint(drawing, undefined);
    const fillPainter = loadedFillPainter();
    if (fillPainter !== undefined) return paint(drawing, fillPainter);
    return loadFillPainter().then((fills) => paint(drawing, fills), options.unavailable);
  };
  const draw = (lay: Layout): void | Promise<void> => {
    const probes = typeof document === "undefined" ? [] : lay.chartFontProbes(spec, ARABIC_PROBE);
    const todo = probes.filter((probe) => !probeScales.has(probe.key));
    // Fonts already settled (an installed face): measured and drawn in this call. A face still
    // loading is measured once the fonts are ready; the returned promise keeps `chartsSettled` waiting.
    if (todo.length > 0 && !measureProbes(todo, false)) {
      return fontsReady().then(() => {
        measureProbes(todo, true);
        // A face still loading after the wait was measured as its fallback: lay out with that
        // (the layout reads the scales at once), but measure again for the next chart.
        const provisional = document.fonts?.status === "loading";
        const drawn = layOut(lay, probes);
        if (provisional) for (const probe of todo) probeScales.delete(probe.key);
        return drawn;
      });
    }
    return layOut(lay, probes);
  };
  if (layout !== undefined) {
    const done = draw(layout);
    if (done !== undefined) track(done);
    return;
  }
  track(loadLayout().then(draw, options.unavailable));
}

/**
 * Fetch the layout code ahead of painting, for a render that waits before it
 * paints (DOCX): its charts then lay out with everything else. The fill
 * painter stays demand-loaded: a chart with gradient, pattern or picture
 * fills fetches it when its laid-out primitives need one, and the render's
 * `chartsSettled()` wait covers that fetch. Resolves either way; a chart
 * painted after a failed fetch tries again.
 */
export function chartLayoutReady(): Promise<void> {
  if (layout !== undefined) return Promise.resolve();
  return loadLayout().then(() => undefined, () => undefined);
}

/**
 * Resolves once every chart painted so far is drawn, or has given up on the
 * layout code. Charts already drawn (the code was loaded) resolve at once.
 */
export function chartsSettled(): Promise<void> {
  return pending.size === 0 ? Promise.resolve() : Promise.allSettled([...pending]).then(() => undefined);
}
