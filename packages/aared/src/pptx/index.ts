/**
 * PPTX renderer: paints a resolved SlideModel to real DOM/SVG — one
 * positioned <div> per shape, inline <svg> for geometry, HTML for text.
 * Never canvas. See AGENTS.md § Rules #1 and docs/ARCHITECTURE.md.
 *
 * Current focused support: text shapes and DrawingML tables from the
 * Arabic gate fixture. Geometry/pictures/fills/rotation join this same
 * DOM model incrementally; no text path ever falls back to canvas.
 */

import type { BoundsModel, ChartModel, GroupModel, PictureModel, ShapeFillModel, ShapeModel, SlideModel, StrokeModel, TableCellModel, TableModel, TextBodyModel, TextParagraphModel } from "../core/model.js";
import type {
  Appearance,
  Collection,
  DocumentInfo,
  MediaCollection as MediaCollectionShape,
  MediaOptions,
  MediaResolver,
  OpenOptions,
  ReadOptions,
  Source,
  UnitRef,
  ZoomValue,
} from "../core/index.js";
import { BULLET } from "../core/chars.js";
import { clipSnippet, foldQuery, foldText, runsText } from "../core/search.js";
import { SessionMedia, WorkerSession } from "../core/session.js";
import { presetPrograms } from "../core/presets.js";
import { itemAt, throwIfAborted } from "../core/errors.js";
import { resolveTheme } from "../appearance.js";
import { appendChartPrimitive } from "../chart-svg.js";
import { MediaUrls } from "../media-url.js";
import { EMU_PER_INCH, PX_PER_INCH, inchesToCss } from "../core/units.js";
import { parseResolvedColor } from "../core/color.js";
import { RenderHandleBase, fitScale, measureFitRoom, safeScale } from "../render-handle.js";
import { createTableCell, createTableGrid } from "../table-dom.js";
import { groupChildMapping } from "./group-mapping.js";
import { applyParagraphBox, applyRunBase, paragraphAlign, pt } from "../text-style.js";
import { lineFontSizePt } from "../core/typography.js";

/**
 * The line pitch PowerPoint gives 100% (`a:lnSpc/a:spcPct val="100000"`)
 * spacing, as a multiple of the font size; percentages scale it. Ground
 * truth: `pptx-coverage.pptx` slide 2's table rows of 28pt text at 100%
 * measure 54.5px in PowerPoint's export — 0.1in of cell insets plus a
 * 44.9px line, 1.20 × 37.3px. Exact point spacing (`a:spcPts`) is
 * unaffected and matched already.
 */
const DRAWINGML_SINGLE_LINE = 1.2;
import { MARK_FALLBACKS } from "../diagnostics.js";
import { fillFallbackReason, fillPaint, officeGradientStops, type FillBox } from "./fill-paint.js";
import { hollowTextFilter } from "./text-outline.js";

/** Document metadata for a deck; `format` is narrowed to `"pptx"`. */
export interface PptxInfo extends DocumentInfo {
  format: "pptx";
}

/**
 * Lightweight slide identity from `open()` indexing — no slide is parsed yet.
 * `title` is `"Slide N"`: `open()` stays as cheap for 200 slides as for 2 by
 * reading no slide part, so a real title (the title placeholder's text) comes
 * from `slide.read()`.
 */
export interface SlideInfo extends UnitRef {
  size: { width: number; height: number };
}

/** One slide result from `deck.search()`. */
export interface PptxSearchMatch {
  /** The slide holding the match; `slide.index` is its `goTo()` key. */
  slide: UnitRef;
  /** First matching paragraph text on the slide. */
  text: string;
}

/**
 * A single slide: a lightweight handle, not a parsed model. `read()`
 * parses the slide in the Worker on demand (index first, content second).
 */
export class SlideHandle {
  constructor(
    readonly info: SlideInfo,
    private readonly session: WorkerSession,
  ) {}

  /**
   * Parse this slide. When its layout has shared graphics (logos, header
   * text), `layout` is attached: the same object for every slide on that
   * layout, fetched once per document. Treat resolved models as read-only.
   */
  async read(options?: ReadOptions): Promise<SlideModel> {
    const signal = options?.signal;
    const requestOptions = signal === undefined ? undefined : { signal };
    const slide = await this.session.unit<SlideModel>(this.info.index, requestOptions);
    if (slide.layoutId !== undefined) slide.layout = await this.session.layout(slide.layoutId, requestOptions);
    return slide;
  }

  /** Picture bytes for this slide's `mediaId`s, via the owning session. */
  media(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.media(mediaId, options);
  }

  /** @internal Bytes for media shared by many slides (layout pictures): fetched once per document, read-only. */
  sharedMedia(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.sharedMedia(mediaId, options);
  }
}

/** `deck.slides`: addressable slides in deck order. */
export class SlideCollection implements Collection<SlideHandle, number> {
  private readonly handles: SlideHandle[];

  constructor(slides: readonly SlideInfo[], session: WorkerSession) {
    this.handles = slides.map((info) => new SlideHandle(info, session));
  }

  list(): readonly SlideHandle[] {
    return this.handles;
  }

  get(index: number): SlideHandle {
    return itemAt(
      this.handles,
      index,
      () => `aared/pptx: slide index ${String(index)} is outside this ${this.handles.length}-slide deck`,
    );
  }
}

function slideText(model: SlideModel): string[] {
  const paragraphs: string[] = [];
  const bodyText = (body: TextBodyModel): void => {
    for (const paragraph of body.paragraphs) {
      const text = runsText(paragraph.runs).trim();
      if (text !== "") paragraphs.push(text);
    }
  };
  const visit = (element: SlideModel["elements"][number]): void => {
    if (element.kind === "group") {
      for (const child of element.children) visit(child);
      return;
    }
    if (element.kind === "shape") {
      if (element.text !== undefined) bodyText(element.text);
    } else if (element.kind === "table") {
      for (const row of element.rows) {
        for (const cell of row.cells) bodyText(cell.text);
      }
    }
  };
  for (const element of model.elements) visit(element);
  return paragraphs;
}

/** One slide paragraph as `search()` scans it: original text plus its folded form. */
interface SlideParagraph {
  text: string;
  lower: string;
}

/** An opened presentation: collections, search, media, and lifecycle. */
export class PptxDocument {
  readonly info: PptxInfo;
  readonly slides: SlideCollection;
  readonly media: MediaCollectionShape;
  /** Per-slide paragraph text, filled by the first `search()` that reaches each slide. */
  private readonly searchIndex = new Map<number, SlideParagraph[]>();

  constructor(private readonly session: WorkerSession) {
    const info = session.info;
    this.info = { ...info, format: "pptx" };
    const sizes = info.sizes ?? [];
    this.slides = new SlideCollection(
      info.titles.map((title, index) => ({
        index,
        title,
        size: { ...(sizes[index] ?? { width: 0, height: 0 }) },
      })),
      session,
    );
    this.media = new SessionMedia(session);
  }

  /**
   * Case-insensitive substring search across slides. The first search reads
   * each slide in order (aborting stops at the next slide) and keeps its
   * paragraph text; later searches scan that text without touching the
   * Worker. One match per slide, the first matching paragraph.
   */
  async search(query: string, options?: ReadOptions): Promise<PptxSearchMatch[]> {
    const needle = foldQuery(query);
    if (needle === null) return [];
    const signal = options?.signal;
    const matches: PptxSearchMatch[] = [];
    for (const handle of this.slides.list()) {
      throwIfAborted(signal);
      const paragraphs = await this.paragraphsOf(handle.info.index, signal);
      const hit = paragraphs.find(({ lower }) => lower.includes(needle));
      if (hit !== undefined) {
        matches.push({ slide: { index: handle.info.index, title: handle.info.title }, text: clipSnippet(hit.text) });
      }
    }
    return matches;
  }

  /** Paragraph text of one slide, read once (search never paints, so no font work). */
  private async paragraphsOf(index: number, signal: AbortSignal | undefined): Promise<SlideParagraph[]> {
    const cached = this.searchIndex.get(index);
    if (cached !== undefined) return cached;
    const model = await this.session.unit<SlideModel>(
      index,
      signal === undefined ? { fonts: false, geometry: false, retain: false } : { signal, fonts: false, geometry: false, retain: false },
    );
    const paragraphs = slideText(model).map((text) => ({ text, lower: foldText(text) }));
    this.searchIndex.set(index, paragraphs);
    return paragraphs;
  }

  /**
   * End the Worker session and reject pending reads. The Worker itself is
   * kept for the next `open()` and stops about 15 s after the last document
   * closes.
   */
  close(): void {
    this.searchIndex.clear();
    this.session.close();
  }
}

/**
 * Open a `.pptx` file. Fails with an `unsupported-format` `AaredError`
 * when the file is not actually a presentation.
 */
export async function open(source: Source, options?: OpenOptions): Promise<PptxDocument> {
  // The preset table ships in this entry only; the session evaluates
  // deferred geometry against it.
  return new PptxDocument(await WorkerSession.openAs("pptx", source, { ...options, geometryPrograms: presetPrograms }));
}

export interface PptxRenderHandle extends EventTarget {
  /** Disconnect layout observers after this slide is no longer displayed. Idempotent. */
  destroy(): void;
  /** `true` once destroyed; every method is then a no-op that fires nothing. */
  readonly destroyed: boolean;
  /**
   * Change zoom: `fit-page` (the default) keeps the whole slide in view,
   * `fit-width` matches the container's width, a number scales the 96-DPI
   * scene. Fires `zoomchange` only when the value changes.
   */
  setZoom(zoom: ZoomValue): void;
  /** The requested zoom. */
  readonly zoom: ZoomValue;
  /** The scale actually painted — the number, or what the fit mode computed. */
  readonly scale: number;
  /** The slide this handle shows. */
  readonly slide: SlideInfo;
  /**
   * Replace the viewer appearance in place without re-reading the slide.
   * Slides are source-faithful by design, so this only refreshes the
   * `data-aared-theme` marker — never recolors authored content.
   */
  setAppearance(appearance: Appearance): void;
}

/** `render()` input: a slide handle (read on demand) or an already-resolved model. */
export interface PptxRenderInput {
  slide: SlideHandle | SlideModel;
  target: HTMLElement;
  appearance?: Appearance;
  /** Starting zoom, applied without a `zoomchange`. Default `"fit-page"`. */
  zoom?: ZoomValue;
  /** Override the default session-backed media resolution. */
  mediaResolver?: MediaResolver;
}

/** `paintSlide()` input: an already-resolved model painted synchronously. */
export interface PaintSlideInput {
  model: SlideModel;
  target: HTMLElement;
  appearance?: Appearance;
  /** Starting zoom, applied without a `zoomchange`. Default `"fit-page"`. */
  zoom?: ZoomValue;
  /** Resolves `mediaId`s to bytes; without one, pictures paint as placeholders. */
  mediaResolver?: MediaResolver;
}

function isSlideHandle(slide: SlideHandle | SlideModel): slide is SlideHandle {
  return slide instanceof SlideHandle;
}

class PptxViewer extends RenderHandleBase<Appearance> implements PptxRenderHandle {
  constructor(initialZoom: ZoomValue, readonly slide: SlideInfo) {
    super(initialZoom);
  }

  setAppearance(appearance: Appearance): void {
    if (this.destroyed) return;
    this.applyAppearance(appearance);
  }
}

/** A bare model's identity, as a handle would report it. */
function slideInfoOf(model: SlideModel): SlideInfo {
  const number = model.number ?? 1;
  return { index: number - 1, title: `Slide ${number}`, size: { ...model.size } };
}

/** Media ids painted under many slides: pictures and picture fills in shared layout graphics, and an inherited picture background. */
function layoutMediaIds(model: SlideModel): ReadonlySet<string> {
  const ids = new Set<string>();
  const visit = (element: SlideModel["elements"][number]): void => {
    if (element.kind === "group") {
      for (const child of element.children) visit(child);
      return;
    }
    if (element.kind === "picture") ids.add(element.mediaId);
    if (element.kind === "shape" && element.fill?.kind === "picture") ids.add(element.fill.mediaId);
    if (element.kind === "table") {
      for (const fill of element.cellFills.flat()) if (fill?.kind === "picture") ids.add(fill.mediaId);
    }
  };
  for (const element of model.layout?.graphics ?? []) visit(element);
  if (model.background?.kind === "picture" && model.background.shared === true) ids.add(model.background.mediaId);
  return ids;
}

/**
 * Paint one slide. Accepts a `SlideHandle` (read from the Worker on
 * demand) or an already-resolved `SlideModel`; when the handle carries
 * its document session, picture media resolves through it unless an
 * explicit `mediaResolver` is given. The caller owns virtualization.
 */
export async function render(input: PptxRenderInput): Promise<PptxRenderHandle> {
  const requested = input.slide;
  let mediaResolver = input.mediaResolver;
  let model: SlideModel;
  let info: SlideInfo;
  if (isSlideHandle(requested)) {
    model = await requested.read();
    info = { ...requested.info, size: { ...requested.info.size } };
    // Pictures the layout paints under every slide are fetched once per
    // document; slide-owned pictures are read per slide.
    const shared = layoutMediaIds(model);
    mediaResolver ??= (mediaId: string) => (shared.has(mediaId) ? requested.sharedMedia(mediaId) : requested.media(mediaId));
  } else {
    model = requested;
    info = slideInfoOf(model);
  }
  return paintSlideAs(info, {
    model,
    target: input.target,
    ...(input.appearance === undefined ? {} : { appearance: input.appearance }),
    ...(input.zoom === undefined ? {} : { zoom: input.zoom }),
    ...(mediaResolver === undefined ? {} : { mediaResolver }),
  });
}

/** Bullet marker font size: absolute points win, else a percentage of the run size, else follow-text. */
function bulletFontSize(paragraph: TextParagraphModel, bullet: NonNullable<TextParagraphModel["bullet"]>): string | undefined {
  if (bullet.sizePts !== undefined) return pt(bullet.sizePts);
  const runSize = paragraph.runs.find((run) => run.fontSizePt !== undefined)?.fontSizePt;
  if (bullet.sizePct !== undefined && runSize !== undefined) return pt((runSize * bullet.sizePct) / 100);
  if (runSize !== undefined) return pt(runSize);
  return undefined;
}

/**
 * Paint one bullet marker. Explicit `buClr`/`buSz` win; otherwise the marker
 * follows its paragraph (`buClrTx`/`buSzTx` default). Picture bullets load
 * through the render's media URLs with the character fallback when fetch
 * fails or no resolver is wired.
 */
function appendBulletMarker(
  line: HTMLElement,
  text: HTMLElement,
  paragraph: TextParagraphModel,
  bullet: NonNullable<TextParagraphModel["bullet"]>,
  urls?: MediaUrls,
): void {
  const marker = document.createElement("span");
  marker.dataset.aaredElement = "bullet-marker";
  marker.setAttribute("aria-hidden", "true");
  const color = bullet.color ?? paragraph.runs.find((run) => run.color !== undefined)?.color;
  if (color !== undefined) marker.style.color = color;
  const size = bulletFontSize(paragraph, bullet);
  if (size !== undefined) marker.style.fontSize = size;
  if (bullet.typeface !== undefined) marker.style.fontFamily = bullet.typeface;
  if (paragraph.direction !== undefined) {
    marker.dataset.aaredBulletSide = paragraph.direction === "rtl" ? "right" : "left";
  }
  const mediaId = bullet.pictureMediaId;
  if (mediaId !== undefined && urls !== undefined) {
    marker.dataset.aaredBulletPicture = mediaId;
    const image = document.createElement("img");
    image.dataset.aaredElement = "bullet-picture";
    image.alt = "";
    image.draggable = false;
    image.style.height = "1em";
    image.style.width = "auto";
    image.style.verticalAlign = "baseline";
    marker.append(image);
    urls
      .url(mediaId, bullet.pictureContentType ?? "image/svg+xml")
      .then((url) => {
        if (url === undefined || !image.isConnected) return;
        image.src = url;
        image.dataset.aaredMediaLoaded = "true";
      })
      .catch(() => {
        if (urls.closed || !marker.isConnected) return;
        image.remove();
        marker.textContent = bullet.char ?? BULLET;
      });
  } else {
    if (mediaId !== undefined) marker.dataset.aaredBulletPictureFallback = "missing-resolver";
    marker.textContent = bullet.char ?? BULLET;
  }
  line.append(marker, text);
}

/**
 * Map a DrawingML underline token (`a:rPr/@u`) to its CSS
 * `text-decoration-style`. Thickness variants collapse to their base style;
 * `none` (and unknown tokens) mean no underline.
 */
function underlineStyle(token: string | undefined): string | undefined {
  if (token === undefined || token === "none") return undefined;
  const normalized = token.toLowerCase();
  if (normalized.includes("dbl") && !normalized.includes("wavy")) return "double";
  if (normalized.includes("dotted")) return "dotted";
  if (normalized.includes("dash")) return "dashed";
  if (normalized.includes("wavy")) return "wavy";
  return "solid";
}

/**
 * CSS `background-image` for a text-run gradient fill (`background-clip:
 * text`), or `undefined` when CSS cannot draw it. A `circle` path is a true
 * circle whose edge reaches the extent's half-diagonal; CSS's `circle
 * farthest-corner` puts it in the same place whenever the focus sits at the
 * extent's centre — PowerPoint's centred and from-a-corner presets alike.
 * `rect` and `shape` paths have no CSS counterpart for text.
 */
function textGradientCss(gradient: ShapeFillModel): string | undefined {
  if (gradient.kind !== "linear-gradient" && !(gradient.kind === "path-gradient" && gradient.path === "circle")) return undefined;
  // PowerPoint's blend (linear light, eased per segment), as with shapes.
  const stops = officeGradientStops(gradient.stops).map((stop) => `${stop.color} ${Math.round(stop.position * 10000) / 100}%`).join(", ");
  if (gradient.kind === "path-gradient") {
    const clamp = (value: number): number => Math.round(Math.min(Math.max(value, 0), 1) * 10000) / 100;
    const x0 = gradient.extent.left;
    const x1 = 1 - gradient.extent.right;
    const y0 = gradient.extent.top;
    const y1 = 1 - gradient.extent.bottom;
    const cx = ((x0 + gradient.focus.left) + (x1 - gradient.focus.right)) / 2;
    const cy = ((y0 + gradient.focus.top) + (y1 - gradient.focus.bottom)) / 2;
    return `radial-gradient(circle farthest-corner at ${clamp(cx)}% ${clamp(cy)}%, ${stops})`;
  }
  // DrawingML 0° points right (clockwise, y down); CSS 0° points up.
  const cssAngle = (((gradient.angle + 90) % 360) + 360) % 360;
  return `linear-gradient(${Math.round(cssAngle * 1000) / 1000}deg, ${stops})`;
}

/**
 * Office's `a:glow` paints a near-solid halo hugging the glyph that fades
 * out to `rad` — a single CSS `text-shadow` blur is one uniform Gaussian
 * spread and reads as a faint, evenly diffuse haze instead (ground truth:
 * coverage slide 6's glow text against a PowerPoint render). Stacking
 * shadows at increasing blur radii approximates that near-solid-to-faded
 * falloff: alpha starts near opaque close to the glyph and eases down to
 * the authored alpha exactly at the authored radius, so the glow's outer
 * edge still matches the source.
 */
function glowTextShadow(glow: { color: string; radiusPt: number }): string {
  const { r, g, b, a } = parseResolvedColor(glow.color) ?? { r: 0, g: 0, b: 0, a: 1 };
  const layers = 4;
  const shadows: string[] = [];
  for (let layer = 1; layer <= layers; layer++) {
    const fraction = layer / layers;
    const alpha = a + (1 - a) * (1 - fraction);
    shadows.push(`0 0 ${pt(glow.radiusPt * fraction)} rgb(${r} ${g} ${b} / ${Math.round(alpha * 1000) / 1000})`);
  }
  return shadows.join(", ");
}

function appendParagraph(
  host: HTMLElement,
  paragraph: TextParagraphModel,
  textWidth?: number,
  wrapping?: TextBodyModel["wrapping"],
  slideNumber?: number,
  urls?: MediaUrls,
): void {
  const line = document.createElement("p");
  line.style.margin = "0";
  line.style.minWidth = "0";
  line.style.whiteSpace = wrapping === "none" ? "nowrap" : "pre-wrap";
  // The `<p>` is sized for its line box, which follows the runs on it (see
  // `lineFontSizePt`); the paragraph default reaches sizeless runs below.
  const defaultSize = paragraph.fontSizePt;
  const lineSize = lineFontSizePt(paragraph.runs, defaultSize);
  if (lineSize !== undefined) line.style.fontSize = `${lineSize}pt`;
  // Indents apply in shapes and table cells alike; the bullet grid already
  // supplies hanging marker/text separation for bulleted paragraphs.
  applyParagraphBox(line, paragraph, { singleLine: DRAWINGML_SINGLE_LINE, firstLineIndent: paragraph.bullet === undefined });
  // Unauthored spacing is PowerPoint's 100%, the same single line.
  if (paragraph.lineSpacing === undefined && paragraph.lineSpacingPt === undefined) {
    line.style.lineHeight = String(DRAWINGML_SINGLE_LINE);
  }

  const text = document.createElement("span");
  text.dataset.aaredElement = "paragraph-text";
  text.style.minWidth = "0";
  for (const run of paragraph.runs) {
    // Hyperlinked runs paint as real anchors; clicks never navigate (view-only)
    // and are surfaced as `linkactivate` by the viewport wiring in `paintSlide`.
    const href = run.hyperlink;
    const span = href === undefined ? document.createElement("span") : document.createElement("a");
    if (span instanceof HTMLAnchorElement && href !== undefined) {
      span.dataset.aaredElement = "run-link";
      span.href = href;
    }
    // A slide-number field shows this slide's number; a master/layout run that
    // carries the same field is shared, so the number is applied at paint time.
    span.textContent = run.field === "slidenum" && slideNumber !== undefined ? String(slideNumber) : run.text;
    applyRunBase(span, run);
    if (run.fontSizePt === undefined && defaultSize !== undefined && defaultSize !== lineSize) {
      span.style.fontSize = `${defaultSize}pt`;
    }
    if (run.noFill === true && run.outline !== undefined) {
      // Hollow WordArt (no fill, an outline): trace the run's merged glyph
      // silhouette instead of stroking each glyph — see text-outline.ts.
      // The fill must stay opaque for the filter to have a real shape to
      // trace; the filter itself discards it and repaints only the ring, so
      // the colour never actually reaches the page.
      span.style.color = run.outline.color;
      span.style.filter = hollowTextFilter(run.outline.widthPt, run.outline.color);
      span.dataset.aaredTextOutline = "true";
      span.dataset.aaredTextNofill = "true";
    } else if (run.noFill === true) {
      // Outline-only text with no outline authored: nothing paints.
      // Transparent keeps the layout box without inventing ink.
      span.style.color = "transparent";
      span.dataset.aaredTextNofill = "true";
    } else if (run.outline !== undefined) {
      // Filled WordArt text with a stroke: keep the fill, add the stroke on
      // top (`paint-order`). PowerPoint decks with this combination are
      // untested; the merged-silhouette filter above only applies to the
      // proven hollow case, where there is no fill for it to discard.
      (span.style as unknown as Record<string, string>).webkitTextStroke = `${pt(run.outline.widthPt)} ${run.outline.color}`;
      span.style.setProperty("paint-order", "stroke fill");
      span.dataset.aaredTextOutline = "true";
    }
    if (run.groupFill === true) span.dataset.aaredTextGroupfill = "true";
    const decorationStyle = underlineStyle(run.underlineStyle);
    if (decorationStyle !== undefined) span.style.textDecorationStyle = decorationStyle;
    if (run.underlineColor !== undefined) span.style.textDecorationColor = run.underlineColor;
    if (run.baselinePct !== undefined) span.style.verticalAlign = `${run.baselinePct}%`;
    // WordArt text gradient: linear and circle-path gradients paint via
    // background-clip; anything CSS cannot draw is marked and keeps the
    // resolved run color.
    if (run.gradient !== undefined) {
      const css = textGradientCss(run.gradient);
      if (css !== undefined) {
        span.style.backgroundImage = css;
        span.style.color = "transparent";
        (span.style as unknown as Record<string, string>).webkitBackgroundClip = "text";
        span.style.backgroundClip = "text";
        span.dataset.aaredTextGradient = run.gradient.kind === "linear-gradient" ? "linear" : "radial";
      } else {
        span.dataset.aaredTextGradient = run.gradient.kind === "unsupported"
          ? run.gradient.source
          : run.gradient.kind === "path-gradient" ? `path-${run.gradient.path}` : run.gradient.kind;
      }
    }
    // Glow and shadow approximate as text-shadow (blur/offset in points).
    const shadows: string[] = [];
    if (run.glow !== undefined) shadows.push(glowTextShadow(run.glow));
    if (run.shadow !== undefined) {
      shadows.push(`${pt(run.shadow.dxPt)} ${pt(run.shadow.dyPt)} ${pt(run.shadow.blurPt)} ${run.shadow.color}`);
    }
    if (shadows.length > 0) span.style.textShadow = shadows.join(", ");
    text.append(span);
  }

  const hasText = paragraph.runs.some((run) => run.text !== "");
  if (paragraph.bullet === undefined || !hasText) {
    line.append(text);
  } else {
    // Grid placement follows the paragraph's explicit or inherited direction:
    // its first logical column is physical left in LTR and right in RTL.
    line.style.display = "grid";
    line.style.gridTemplateColumns = "auto minmax(0, 1fr)";
    line.style.gap = "0.35em";
    if (paragraphAlign(paragraph) === "left") {
      // A left-aligned RTL bullet occupies only its content width so its marker
      // follows the text instead of staying at the text box's far-right edge.
      line.style.width = "fit-content";
      line.style.alignSelf = "flex-start";
    }
    appendBulletMarker(line, text, paragraph, paragraph.bullet, urls);
  }
  host.append(line);
}

function appendTextBody(host: HTMLElement, body: TextBodyModel, textWidth?: number, slideNumber?: number, urls?: MediaUrls): void {
  // WordArt warp presets paint unwarped; the marker keeps the authored
  // intent visible instead of silently dropping it.
  if (body.textWarp !== undefined) host.dataset.aaredTextWarp = body.textWarp;
  for (const paragraph of body.paragraphs) appendParagraph(host, paragraph, textWidth, body.wrapping, slideNumber, urls);
}

function appendShapeText(host: HTMLElement, shape: ShapeModel, slideNumber?: number, urls?: MediaUrls): void {
  if (shape.text === undefined) return;
  const text = document.createElement("div");
  text.dataset.aaredElement = "shape-text";
  text.style.position = "absolute";
  text.style.inset = "0";
  text.style.display = "flex";
  text.style.flexDirection = "column";
  text.style.minWidth = "0";
  text.style.minHeight = "0";
  text.style.boxSizing = "border-box";
  const autofit = shape.text.autofit ?? "none";
  // Real autofit: `none` clips, `shrinkText` reduces fonts to fit, and
  // `resizeShape` lets the text grow beyond its authored bounds instead of
  // clipping. The source value stays on the model; only the browser behavior
  // changes here.
  text.style.overflow = autofit === "resizeShape" ? "visible" : "hidden";
  if (autofit === "resizeShape") {
    host.style.overflow = "visible";
    text.style.position = "relative";
    text.style.inset = "auto";
    text.style.width = "100%";
    text.style.minHeight = "100%";
    text.style.flexShrink = "0";
  }
  const body = shape.text;
  const bounds = shape.bounds;
  if (bounds !== undefined && body.insets !== undefined) {
    if (body.insets.left !== undefined) text.style.paddingLeft = inchesToCss(body.insets.left);
    if (body.insets.right !== undefined) text.style.paddingRight = inchesToCss(body.insets.right);
    if (body.insets.top !== undefined) text.style.paddingTop = inchesToCss(body.insets.top);
    if (body.insets.bottom !== undefined) text.style.paddingBottom = inchesToCss(body.insets.bottom);
  }
  switch (body.verticalAlign) {
    case "middle":
      text.style.justifyContent = "center";
      break;
    case "bottom":
      text.style.justifyContent = "flex-end";
      break;
    case "justify":
      text.style.justifyContent = "space-between";
      break;
    case "distributed":
      text.style.justifyContent = "space-around";
      break;
    default:
      text.style.justifyContent = "flex-start";
  }
  if (body.wrapping === "none") text.style.whiteSpace = "nowrap";
  text.dataset.aaredVerticalAlign = body.verticalAlign ?? "top";
  if (body.wrapping !== undefined) text.dataset.aaredWrapping = body.wrapping;
  if (body.autofit !== undefined) text.dataset.aaredAutofit = body.autofit;
  appendTextBody(text, body, bounds?.width, slideNumber, urls);
  // Whole-shape links wrap the text layer only (`display: contents` keeps
  // layout identical); run-level anchors inside still win their own clicks
  // by bubbling first. Surfaced as `linkactivate`, never navigated.
  if (shape.hyperlink !== undefined) {
    const anchor = document.createElement("a");
    anchor.dataset.aaredElement = "shape-link";
    anchor.href = shape.hyperlink;
    anchor.style.display = "contents";
    host.append(anchor);
    anchor.append(text);
  } else {
    host.append(text);
  }
}

function scaleFonts(root: HTMLElement, factor: number): void {
  const targets: HTMLElement[] = [root, ...root.querySelectorAll<HTMLElement>("p, span")];
  for (const element of targets) {
    const current = element.style.fontSize;
    if (current === "") continue;
    const match = /^([\d.]+)(pt|px|em|rem|%)$/.exec(current.trim());
    if (match === null) continue;
    element.style.fontSize = `${Number(match[1]) * factor}${match[2]}`;
  }
}

function shrinkTextToFit(text: HTMLElement): void {
  // PowerPoint's `normAutofit` reduces type until the text fits its box.
  // Measure synchronously (forced reflow) and scale authored font sizes down
  // only — never up — so a fitted box keeps its authored metrics.
  for (let iteration = 0; iteration < 8; iteration++) {
    const overflowsVertically = text.scrollHeight > text.clientHeight + 1;
    const overflowsHorizontally = text.scrollWidth > text.clientWidth + 1;
    if (!overflowsVertically && !overflowsHorizontally) break;
    const verticalScale = text.clientHeight > 0 ? text.clientHeight / Math.max(text.scrollHeight, 1) : 1;
    const horizontalScale = text.clientWidth > 0 ? text.clientWidth / Math.max(text.scrollWidth, 1) : 1;
    const scale = Math.min(verticalScale, horizontalScale) * 0.95;
    if (!Number.isFinite(scale) || scale >= 1 || scale <= 0) break;
    // Clamp a single pass so one long word cannot collapse the whole box.
    scaleFonts(text, Math.max(scale, 0.6));
    if (scale > 0.95) break;
  }
  text.dataset.aaredAutofitApplied = "shrink";
}

/** A dev-only dashed outline over a box whose fill could not be painted (a table cell, the slide). */
function appendFillMarker(host: HTMLElement, reason: string): void {
  host.dataset.aaredFillFallback = reason;
  if (!MARK_FALLBACKS) return;
  const marker = document.createElement("div");
  marker.dataset.aaredElement = "fill-fallback";
  marker.setAttribute("aria-hidden", "true");
  marker.style.position = "absolute";
  marker.style.inset = "2px";
  marker.style.border = "1px dashed #9CA3AF";
  marker.style.pointerEvents = "none";
  host.append(marker);
}

function appendShapeGeometry(host: HTMLElement, shape: ShapeModel, gradientId: string, urls: MediaUrls | undefined): void {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.dataset.aaredElement = "shape-geometry";
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute";
  svg.style.inset = "0";
  svg.style.width = "100%";
  svg.style.height = "100%";
  svg.style.overflow = "visible";
  const geometry = shape.geometry;
  const viewBox = geometry?.viewBox ?? { width: 100000, height: 100000 };
  const fallback = geometry?.fallback ?? (geometry === undefined ? "missing" : undefined);
  // Only genuinely unpaintable geometry is marked — and only in development
  // builds. A shape with no authored geometry at all ("missing": text-only
  // placeholders like titles and footers) is an invisible container in
  // Office; marking it would draw dashed boxes and fragment slivers around
  // every title and footer on the slide.
  const marked = fallback === "unsupported" || fallback === "custom";
  if (marked) {
    host.dataset.aaredGeometryFallback = fallback;
  }
  // A sub-pixel authored axis (hairline connectors) rounds the svg viewport
  // to zero clientHeight and Chrome paints nothing. Grow that axis to one
  // real pixel and widen the viewBox by the same factor, so the scale stays
  // exactly 96px per inch and the stroke still paints at its authored width.
  const emuPerPx = EMU_PER_INCH / PX_PER_INCH;
  const vbWidth = Math.max(viewBox.width, emuPerPx);
  const vbHeight = Math.max(viewBox.height, emuPerPx);
  svg.setAttribute("viewBox", `0 0 ${vbWidth} ${vbHeight}`);
  const grewViewport = vbWidth !== viewBox.width || vbHeight !== viewBox.height;
  if (grewViewport) {
    svg.style.width = `${vbWidth / emuPerPx}px`;
    svg.style.height = `${vbHeight / emuPerPx}px`;
  }
  const viewportPx = grewViewport ? vbWidth / emuPerPx : shape.bounds === undefined ? 0 : shape.bounds.width * PX_PER_INCH;
  const viewportPxHeight = grewViewport ? vbHeight / emuPerPx : shape.bounds === undefined ? 0 : shape.bounds.height * PX_PER_INCH;
  let paint = "none";
  let fillMarked = false;
  if (shape.fill !== undefined && !marked) {
    const painted = fillPaint(svg, gradientId, shape.fill, {
      width: vbWidth,
      height: vbHeight,
      unitsPerPxX: viewportPx > 0 ? vbWidth / viewportPx : emuPerPx,
      unitsPerPxY: viewportPxHeight > 0 ? vbHeight / viewportPxHeight : emuPerPx,
      // Group children use their child-space position: an approximation of
      // the slide origin a pattern is anchored to.
      slideX: (shape.bounds?.x ?? 0) * PX_PER_INCH,
      slideY: (shape.bounds?.y ?? 0) * PX_PER_INCH,
    }, {
      urls,
      ...(shape.transform === undefined ? {} : { transform: shape.transform }),
      // A `path="shape"` gradient follows this shape's own outline.
      ...(geometry?.path === undefined ? {} : { outline: geometry.path }),
    });
    if (painted === undefined) {
      // The outline is right but the fill cannot be drawn faithfully:
      // development builds dash the outline; shipped builds keep the
      // authored stroke and leave the inside empty.
      fillMarked = true;
      host.dataset.aaredFillFallback = fillFallbackReason(shape.fill, urls);
    } else {
      paint = painted;
    }
  }
  // A filled rectangle reads as a wrong shape. Development builds mark
  // unpaintable geometry with an empty dashed outline (the picture/chart
  // convention); shipped builds render no box at all — text still paints.
  const dashed = MARK_FALLBACKS && (marked || fillMarked);
  const outlined = marked ? MARK_FALLBACKS : shape.stroke !== undefined || dashed;
  // `a:ln w="0"` is PowerPoint's thinnest line, not an invisible one: it
  // paints a one-pixel hairline (slide 4 of `pptx-coverage.pptx`), so no
  // authored stroke goes below one CSS pixel.
  const strokePx = Math.max((shape.stroke?.width ?? 1 / PX_PER_INCH) * PX_PER_INCH, 1);
  const strokeWidth = viewportPx > 0 ? strokePx / (viewportPx / vbWidth) : 1;
  // Multi-path presets (can, cube, callouts) split into layers: fill-only
  // faces beneath, the main outline, then stroke-only details on top.
  const layers: Array<readonly [d: string, fills: boolean, strokes: boolean]> = [
    ...(geometry?.fillOnly === undefined ? [] : [[geometry.fillOnly, true, false] as const]),
    [geometry?.path ?? "M 0 0 H 100000 V 100000 H 0 Z", true, true],
    ...(geometry?.strokeOnly === undefined ? [] : [[geometry.strokeOnly, false, true] as const]),
  ];
  for (const [d, fills, strokes] of layers) {
    if (d === "") continue;
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", d);
    path.setAttribute("fill", fills ? paint : "none");
    if (strokes && outlined) {
      path.setAttribute("stroke", shape.stroke?.color ?? "#9CA3AF");
      path.setAttribute("stroke-width", `${strokeWidth}`);
      if (dashed) path.setAttribute("stroke-dasharray", `${4 * emuPerPx} ${3 * emuPerPx}`);
    }
    svg.append(path);
  }
  host.append(svg);
}

function position(bounds: BoundsModel | undefined, element: HTMLElement): void {
  if (bounds !== undefined) {
    element.style.left = inchesToCss(bounds.x);
    element.style.top = inchesToCss(bounds.y);
    element.style.width = inchesToCss(bounds.width);
    element.style.height = inchesToCss(bounds.height);
  } else {
    element.style.inset = "0";
  }
}

/** A per-document-unique SVG id: `owner` is the slide part or the layout id the shape belongs to. */
function gradientId(owner: string | undefined, shape: Pick<ShapeModel, "id">, index: number | string): string {
  let hash = 0;
  for (const char of typeof owner === "string" ? owner : "inline") hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return `aared-gradient-${hash.toString(36)}-${shape.id}-${index}`;
}

/**
 * Paint a non-solid cell fill behind the cell's text: an SVG layer sized to
 * the cell in CSS pixels at slide scale, through the shared fill painter.
 * `isolation` keeps the layer's negative z-index inside the cell, above the
 * table grid but under the text.
 */
function appendCellFill(cell: HTMLElement, fill: ShapeFillModel, id: string, box: FillBox, urls: MediaUrls | undefined): void {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.dataset.aaredElement = "cell-fill";
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("viewBox", `0 0 ${box.width} ${box.height}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute";
  svg.style.inset = "0";
  svg.style.width = "100%";
  svg.style.height = "100%";
  svg.style.zIndex = "-1";
  cell.style.position = "relative";
  cell.style.isolation = "isolate";
  const paint = fillPaint(svg, id, fill, box, { urls });
  if (paint === undefined) {
    appendFillMarker(cell, fillFallbackReason(fill, urls));
    return;
  }
  const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  rect.setAttribute("width", String(box.width));
  rect.setAttribute("height", String(box.height));
  rect.setAttribute("fill", paint);
  svg.append(rect);
  cell.prepend(svg);
}

/**
 * Each logical cell's box in slide CSS pixels, from the authored column
 * widths and row heights. Rows without an authored height share what is left
 * of the table's height. RTL tables place logical column 0 at the right.
 */
function tableCellBoxes(table: TableModel): Array<Array<{ x: number; y: number; width: number; height: number }>> {
  const tableWidth = table.columns.reduce((sum, width) => sum + width, 0);
  const authored = table.rows.reduce((sum, row) => sum + (row.height ?? 0), 0);
  const unsized = table.rows.filter((row) => row.height === undefined).length;
  const spare = unsized === 0 ? 0 : Math.max(0, (table.bounds?.height ?? authored) - authored) / unsized;
  const heights = table.rows.map((row) => row.height ?? spare);
  let top = table.bounds?.y ?? 0;
  return table.rows.map((row, rowIndex) => {
    let column = 0;
    const boxes = row.cells.map((cell) => {
      const span = cell.columnSpan ?? 1;
      const before = table.columns.slice(0, column).reduce((sum, width) => sum + width, 0);
      const width = table.columns.slice(column, column + span).reduce((sum, width) => sum + width, 0);
      const height = heights.slice(rowIndex, rowIndex + (cell.rowSpan ?? 1)).reduce((sum, value) => sum + value, 0);
      const left = (table.bounds?.x ?? 0) + (table.direction === "rtl" ? tableWidth - before - width : before);
      column += span;
      return { x: left * PX_PER_INCH, y: top * PX_PER_INCH, width: width * PX_PER_INCH, height: height * PX_PER_INCH };
    });
    top += heights[rowIndex] ?? 0;
    return boxes;
  });
}

function appendTable(scene: HTMLElement, table: TableModel, fillId: string, slideNumber?: number, urls?: MediaUrls): void {
  const cellBoxes = tableCellBoxes(table);
  const grid = createTableGrid(table.direction, table.columns);
  grid.style.position = "absolute";
  grid.style.transformOrigin = "center";
  const tableTransforms: string[] = [];
  if (table.transform?.rotation !== undefined) tableTransforms.push(`rotate(${table.transform.rotation}deg)`);
  if (table.transform?.flipHorizontal === true || table.transform?.flipVertical === true) {
    tableTransforms.push(`scale(${table.transform.flipHorizontal === true ? -1 : 1}, ${table.transform.flipVertical === true ? -1 : 1})`);
  }
  if (tableTransforms.length > 0) grid.style.transform = tableTransforms.join(" ");
  // Styled tables carry resolved per-cell borders and PowerPoint draws no
  // frame around them. Only a table with none gets this explicit neutral
  // frame (never `currentColor`, so it cannot inherit embedding chrome).
  if (table.rows.every((row) => row.cells.every((cell) => cell.borders === undefined))) {
    grid.style.border = "1px solid #6B7280";
  }
  // Keeps the edge overlay's z-index inside the table, so shapes stacked
  // above the table on the slide still paint over its lines.
  grid.style.isolation = "isolate";
  if (table.bounds !== undefined) {
    grid.style.left = inchesToCss(table.bounds.x);
    grid.style.top = inchesToCss(table.bounds.y);
    grid.style.width = inchesToCss(table.bounds.width);
  }
  // PowerPoint's table height is its rows', not the frame's: a frame
  // authored shorter than its rows grows to hold them, and a row grows past
  // its authored height to hold its text. Rows with no height share the
  // frame's height.
  grid.style.gridTemplateRows = table.rows.map((row) => row.height === undefined ? "auto" : `minmax(${inchesToCss(row.height)}, auto)`).join(" ");
  if (table.bounds !== undefined && table.rows.some((row) => row.height === undefined)) {
    grid.style.minHeight = inchesToCss(table.bounds.height);
  }
  const placements = tableCellPlacements(table);
  const overlays: HTMLElement[] = [];
  for (const [rowIndex, row] of table.rows.entries()) {
    for (const [columnIndex, cellModel] of row.cells.entries()) {
      const cell = createTableCell(cellModel);
      cell.style.minHeight = "0";
      const place = placements[rowIndex]![columnIndex]!;
      cell.style.gridRow = `${place.row + 1} / span ${cellModel.rowSpan ?? 1}`;
      cell.style.gridColumn = `${place.column + 1} / span ${cellModel.columnSpan ?? 1}`;
      const margins = cellModel.margins;
      cell.style.paddingTop = inchesToCss(margins?.top ?? 0.05);
      cell.style.paddingRight = inchesToCss(margins?.right ?? 0.1);
      cell.style.paddingBottom = inchesToCss(margins?.bottom ?? 0.05);
      cell.style.paddingLeft = inchesToCss(margins?.left ?? 0.1);
      cell.dataset.aaredCellMargins = "resolved";
      // Resolved edges paint in an overlay (below); a cell with none keeps
      // the faint grid fallback.
      if (cellModel.borders !== undefined) {
        overlays.push(cellEdgeOverlay(cellModel.borders, cell.style.gridRow, cell.style.gridColumn));
        cell.dataset.aaredCellBorders = "resolved";
      } else {
        cell.style.border = "1px solid color-mix(in srgb, #6B7280 25%, transparent)";
      }
      const fill = table.cellFills[rowIndex]?.[columnIndex];
      const box = cellBoxes[rowIndex]?.[columnIndex];
      if (fill?.kind === "solid") {
        cell.style.backgroundColor = fill.color;
      } else if (fill !== undefined && box !== undefined && box.width > 0 && box.height > 0) {
        appendCellFill(cell, fill, `${fillId}-r${rowIndex}c${columnIndex}`, {
          width: box.width,
          height: box.height,
          unitsPerPxX: 1,
          unitsPerPxY: 1,
          slideX: box.x,
          slideY: box.y,
        }, urls);
      }
      if (cellModel.textRotation !== undefined) {
        // `writing-mode` turns the lines; `vert270` reads bottom to top, the
        // same vertical flow turned half a turn.
        const rotated = document.createElement("div");
        rotated.dataset.aaredElement = "cell-rotated-text";
        rotated.style.writingMode = "vertical-rl";
        rotated.style.alignSelf = "center";
        if (cellModel.textRotation === 270) rotated.style.transform = "rotate(180deg)";
        appendTextBody(rotated, cellModel.text, undefined, slideNumber, urls);
        cell.append(rotated);
      } else {
        appendTextBody(cell, cellModel.text, undefined, slideNumber, urls);
      }
      grid.append(cell);
    }
  }
  grid.append(...overlays);
  scene.append(grid);
}

/**
 * Each model cell's top-left grid slot. Merge continuations are absent from
 * the model, so a cell's column is the first slot its row leaves free after
 * earlier cells' column spans and earlier rows' row spans.
 */
function tableCellPlacements(table: TableModel): Array<Array<{ row: number; column: number }>> {
  const taken = new Set<string>();
  return table.rows.map((row, rowIndex) => {
    let column = 0;
    return row.cells.map((cell) => {
      while (taken.has(`${rowIndex}:${column}`)) column++;
      const place = { row: rowIndex, column };
      for (let r = 0; r < (cell.rowSpan ?? 1); r++) {
        for (let c = 0; c < (cell.columnSpan ?? 1); c++) taken.add(`${rowIndex + r}:${column + c}`);
      }
      column += cell.columnSpan ?? 1;
      return place;
    });
  });
}

function borderStyle(stroke: StrokeModel): string {
  return stroke.dash === "dot" ? "dotted" : stroke.dash === "dash" ? "dashed" : "solid";
}

/**
 * A cell's edges as lines centred on the cell boundary, the way PowerPoint
 * strokes them: a shared edge drawn by both neighbours lands on the same
 * pixels instead of doubling, and a thick edge straddles the boundary rather
 * than eating into the cell. The overlay occupies the cell's own grid area
 * above every cell's fill, so a neighbour's background never covers the half
 * of a line that overhangs into it.
 */
function cellEdgeOverlay(borders: NonNullable<TableCellModel["borders"]>, gridRow: string, gridColumn: string): HTMLElement {
  const overlay = document.createElement("div");
  overlay.dataset.aaredElement = "table-cell-edges";
  overlay.setAttribute("aria-hidden", "true");
  overlay.style.gridRow = gridRow;
  overlay.style.gridColumn = gridColumn;
  overlay.style.position = "relative";
  overlay.style.zIndex = "1";
  overlay.style.pointerEvents = "none";
  const sides: Array<["top" | "bottom" | "left" | "right", StrokeModel | undefined]> = [
    ["top", borders.top], ["bottom", borders.bottom], ["left", borders.left], ["right", borders.right],
  ];
  for (const [side, stroke] of sides) {
    if (stroke === undefined) continue;
    const width = stroke.width * PX_PER_INCH;
    const line = document.createElement("div");
    line.dataset.aaredEdge = side;
    line.style.position = "absolute";
    const horizontal = side === "top" || side === "bottom";
    const rule = `${width}px ${borderStyle(stroke)} ${stroke.color}`;
    if (horizontal) {
      line.style.left = `${-width / 2}px`;
      line.style.right = `${-width / 2}px`;
      line.style[side] = `${-width / 2}px`;
      line.style.borderTop = rule;
    } else {
      line.style.top = `${-width / 2}px`;
      line.style.bottom = `${-width / 2}px`;
      line.style[side] = `${-width / 2}px`;
      line.style.borderLeft = rule;
    }
    overlay.append(line);
  }
  const diagonals: Array<[StrokeModel | undefined, string]> = [
    [borders.diagonalDown, "M 0 0 L 1 1"],
    [borders.diagonalUp, "M 0 1 L 1 0"],
  ];
  for (const [stroke, path] of diagonals) {
    if (stroke === undefined) continue;
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 1 1");
    svg.setAttribute("preserveAspectRatio", "none");
    svg.style.position = "absolute";
    svg.style.inset = "0";
    svg.style.width = "100%";
    svg.style.height = "100%";
    svg.style.overflow = "visible";
    const line = document.createElementNS("http://www.w3.org/2000/svg", "path");
    line.setAttribute("d", path);
    line.setAttribute("stroke", stroke.color);
    line.setAttribute("stroke-width", String(stroke.width * PX_PER_INCH));
    line.setAttribute("vector-effect", "non-scaling-stroke");
    if (stroke.dash !== undefined) {
      const unit = stroke.width * PX_PER_INCH;
      line.setAttribute("stroke-dasharray", stroke.dash === "dot" ? `${unit} ${unit}` : `${unit * 4} ${unit * 3}`);
    }
    svg.append(line);
    overlay.append(svg);
  }
  return overlay;
}

function transformFor(transform: { rotation?: number; flipHorizontal?: boolean; flipVertical?: boolean } | undefined): string {
  const parts: string[] = [];
  if (transform?.rotation !== undefined) parts.push(`rotate(${transform.rotation}deg)`);
  if (transform?.flipHorizontal === true || transform?.flipVertical === true) {
    parts.push(`scale(${transform.flipHorizontal === true ? -1 : 1}, ${transform.flipVertical === true ? -1 : 1})`);
  }
  return parts.join(" ");
}

function appendPictureFallback(host: HTMLElement, picture: PictureModel, reason: string): void {
  host.dataset.aaredPictureFallback = reason;
  const fallback = document.createElement("div");
  fallback.dataset.aaredElement = "picture-fallback";
  fallback.style.position = "absolute";
  fallback.style.inset = "0";
  fallback.style.display = "flex";
  fallback.style.alignItems = "center";
  fallback.style.justifyContent = "center";
  fallback.style.boxSizing = "border-box";
  fallback.style.border = "1px dashed currentColor";
  fallback.style.fontSize = "12px";
  fallback.style.padding = "0.5em";
  fallback.style.textAlign = "center";
  // Placeholder chrome, not content: the shape name must never leak into
  // clipboard text (select-all + copy otherwise yields "Ink 52…").
  fallback.style.userSelect = "none";
  fallback.setAttribute("aria-hidden", "true");
  fallback.textContent = picture.name === "" ? "Image unavailable" : picture.name;
  host.append(fallback);
}

function appendPicture(scene: HTMLElement, picture: PictureModel, urls: MediaUrls | undefined): void {
  const wrapper = document.createElement("div");
  wrapper.dataset.aaredElement = "picture";
  wrapper.dataset.aaredPictureId = picture.id;
  wrapper.dataset.aaredMediaId = picture.mediaId;
  wrapper.dataset.aaredContentType = picture.contentType;
  wrapper.style.position = "absolute";
  wrapper.style.overflow = "hidden";
  wrapper.style.transformOrigin = "center";
  wrapper.style.background = "transparent";
  // Pictures (and ink rasters) are not text: keep them out of selections so
  // select-all copies slide text only.
  wrapper.style.userSelect = "none";
  const transform = transformFor(picture.transform);
  if (transform !== "") wrapper.style.transform = transform;
  position(picture.bounds, wrapper);
  if (picture.crop !== undefined) {
    wrapper.dataset.aaredCropLeft = String(picture.crop.left);
    wrapper.dataset.aaredCropTop = String(picture.crop.top);
    wrapper.dataset.aaredCropRight = String(picture.crop.right);
    wrapper.dataset.aaredCropBottom = String(picture.crop.bottom);
  }

  if (urls === undefined) {
    appendPictureFallback(wrapper, picture, "missing-resolver");
    scene.append(wrapper);
    return;
  }

  const image = document.createElement("img");
  image.dataset.aaredElement = "picture-image";
  image.alt = picture.name;
  (image as HTMLImageElement).draggable = false;
  image.style.position = "absolute";
  image.style.display = "block";
  const crop = picture.crop;
  if (crop === undefined) {
    image.style.left = "0";
    image.style.top = "0";
    image.style.width = "100%";
    image.style.height = "100%";
    image.style.objectFit = "fill";
  } else {
    // DrawingML `a:srcRect` crops fractions off the source before stretching
    // the remainder to the authored bounds.
    const visibleWidth = 1 - crop.left - crop.right;
    const visibleHeight = 1 - crop.top - crop.bottom;
    if (visibleWidth <= 0 || visibleHeight <= 0) {
      appendPictureFallback(wrapper, picture, "invalid-crop");
      scene.append(wrapper);
      return;
    }
    image.style.width = `${(100 / visibleWidth).toFixed(4)}%`;
    image.style.height = `${(100 / visibleHeight).toFixed(4)}%`;
    image.style.left = `${((-crop.left / visibleWidth) * 100).toFixed(4)}%`;
    image.style.top = `${((-crop.top / visibleHeight) * 100).toFixed(4)}%`;
    image.style.maxWidth = "none";
    image.style.maxHeight = "none";
    image.style.objectFit = "fill";
  }
  wrapper.append(image);
  scene.append(wrapper);

  let settled = false;
  urls
    .url(picture.mediaId, picture.contentType)
    .then((url) => {
      if (settled || url === undefined || !wrapper.isConnected) return;
      settled = true;
      image.src = url;
      image.dataset.aaredMediaLoaded = "true";
    })
    .catch(() => {
      if (settled || urls.closed || !wrapper.isConnected) return;
      settled = true;
      image.remove();
      appendPictureFallback(wrapper, picture, "load-error");
    });
}

function appendChart(scene: HTMLElement, chart: ChartModel): void {
  const wrapper = document.createElement("div");
  wrapper.dataset.aaredElement = "chart";
  wrapper.dataset.aaredChartId = chart.id;
  if (chart.sourceChartId !== undefined) wrapper.dataset.aaredSourceChartId = chart.sourceChartId;
  wrapper.style.position = "absolute";
  wrapper.style.overflow = "hidden";
  wrapper.style.transformOrigin = "center";
  wrapper.style.background = "white";
  wrapper.style.userSelect = "none";
  const transform = transformFor(chart.transform);
  if (transform !== "") wrapper.style.transform = transform;
  position(chart.bounds, wrapper);
  if (chart.graphic === undefined) {
    wrapper.dataset.aaredChartFallback = "unsupported";
    const fallback = document.createElement("div");
    fallback.dataset.aaredElement = "chart-fallback";
    fallback.style.position = "absolute";
    fallback.style.inset = "0";
    fallback.style.display = "flex";
    fallback.style.alignItems = "center";
    fallback.style.justifyContent = "center";
    fallback.style.boxSizing = "border-box";
    fallback.style.border = "1px dashed currentColor";
    fallback.style.fontSize = "12px";
    fallback.style.userSelect = "none";
    fallback.setAttribute("aria-hidden", "true");
    fallback.textContent = chart.name === "" ? "Chart" : chart.name;
    wrapper.append(fallback);
    scene.append(wrapper);
    return;
  }
  wrapper.dataset.aaredChartPrimitives = String(chart.graphic.primitives.length);
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.dataset.aaredElement = "chart-geometry";
  svg.setAttribute("viewBox", `0 0 ${chart.graphic.viewBox.width} ${chart.graphic.viewBox.height}`);
  svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
  svg.setAttribute("role", "img");
  svg.style.position = "absolute";
  svg.style.inset = "0";
  svg.style.width = "100%";
  svg.style.height = "100%";
  for (const primitive of chart.graphic.primitives) appendChartPrimitive(svg, primitive);
  wrapper.append(svg);
  scene.append(wrapper);
}

/**
 * Paint the slide's authored background under the scene. A solid fill is a
 * plain CSS colour; every other fill goes through the same SVG fill painter
 * as shapes and table cells, so they agree (picture media shares the slide
 * pictures' resolver and object-URL lifecycle). A background it cannot paint
 * leaves the slide white and is marked, so it is not mistaken for an
 * authored white. The background is the viewport's first child: the scene is
 * already attached, and a later sibling would paint over it.
 */
function paintBackground(viewport: HTMLElement, model: SlideModel, urls: MediaUrls | undefined): void {
  const fill = model.background;
  if (fill === undefined) return;
  if (fill.kind === "solid") {
    viewport.style.background = fill.color;
    viewport.dataset.aaredBackground = "solid";
    return;
  }
  const width = model.size.width * PX_PER_INCH;
  const height = model.size.height * PX_PER_INCH;
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.dataset.aaredElement = "slide-background";
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute";
  svg.style.inset = "0";
  svg.style.width = "100%";
  svg.style.height = "100%";
  const paint = fillPaint(svg, gradientId(model.part, { id: "background" }, -1), fill, {
    width,
    height,
    unitsPerPxX: 1,
    unitsPerPxY: 1,
    slideX: 0,
    slideY: 0,
  }, { urls });
  if (paint === undefined) {
    viewport.dataset.aaredBackground = "unsupported";
    viewport.dataset.aaredBackgroundSource = fillFallbackReason(fill, urls);
    appendFillMarker(viewport, fillFallbackReason(fill, urls));
    return;
  }
  const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  rect.setAttribute("width", String(width));
  rect.setAttribute("height", String(height));
  rect.setAttribute("fill", paint);
  svg.append(rect);
  viewport.prepend(svg);
  viewport.dataset.aaredBackground = fill.kind;
  if (fill.kind === "picture") viewport.dataset.aaredMediaId = fill.mediaId;
}

interface ElementPaintContext {
  scene: HTMLElement;
  /** The slide's number, shown by `slidenum` fields (including in shared layout graphics). */
  slideNumber: number | undefined;
  /** Media URLs for pictures and picture bullets; absent when the embedder gave no resolver. */
  urls: MediaUrls | undefined;
  /** SVG gradient-id owner: slide part or layout id; groups keep their children on the same owner. */
  owner: string;
}

/**
 * Append elements back-to-front into the scene. Slide elements and shared
 * layout graphics both go through here; `owner` (slide part or layout id)
 * keeps their SVG gradient ids apart. Groups recurse through `appendGroup`.
 */
function paintElements(context: ElementPaintContext, elements: readonly SlideModel["elements"][number][], path = ""): void {
  const { scene, slideNumber, urls, owner } = context;
  for (const [index, element] of elements.entries()) {
    const gradientKey = path === "" ? `${index}` : `${path}.${index}`;
    if (element.kind === "table") {
      appendTable(scene, element, gradientId(owner, element, gradientKey), slideNumber, urls);
      continue;
    }
    if (element.kind === "picture") {
      appendPicture(scene, element, urls);
      continue;
    }
    if (element.kind === "chart") {
      appendChart(scene, element);
      continue;
    }
    if (element.kind === "group") {
      appendGroup(context, element, gradientKey);
      continue;
    }
    if (element.kind !== "shape") continue;
    const shape = document.createElement("div");
    shape.dataset.aaredElement = "shape";
    shape.dataset.aaredShapeId = element.id;
    shape.style.position = "absolute";
    shape.style.overflow = "visible";
    shape.style.transformOrigin = "center";
    const shapeTransform = transformFor(element.transform);
    if (shapeTransform !== "") shape.style.transform = shapeTransform;
    position(element.bounds, shape);
    appendShapeGeometry(shape, element, gradientId(owner, element, gradientKey), urls);
    appendShapeText(shape, element, slideNumber, urls);
    scene.append(shape);
  }
}

/**
 * Paint one group: an outer wrapper carries the group's bounds and
 * rotate/flip transform (origin center, like shapes); an inner div maps
 * the authored child space onto the wrapper so children keep their
 * authored coordinates.
 */
function appendGroup(context: ElementPaintContext, group: GroupModel, path: string): void {
  const { scene, slideNumber, urls } = context;
  const wrapper = document.createElement("div");
  wrapper.dataset.aaredElement = "group";
  wrapper.dataset.aaredGroupId = group.id;
  wrapper.style.position = "absolute";
  wrapper.style.overflow = "visible";
  wrapper.style.transformOrigin = "center";
  const groupTransform = transformFor(group.transform);
  if (groupTransform !== "") wrapper.style.transform = groupTransform;
  position(group.bounds, wrapper);

  const inner = document.createElement("div");
  inner.dataset.aaredElement = "group-children";
  inner.style.position = "absolute";
  inner.style.inset = "0";
  const mapping = groupChildMapping(group.bounds, group.childSpace);
  if (mapping !== "") {
    inner.style.transformOrigin = "0 0";
    inner.style.transform = mapping;
  }
  wrapper.append(inner);
  scene.append(wrapper);
  paintElements({ ...context, scene: inner }, group.children, path);
}

/**
 * Paint one already-resolved slide synchronously. The caller owns
 * virtualization. Same input shape as `render()`, with the resolved `model`
 * in place of a slide handle.
 */
export function paintSlide(input: PaintSlideInput): PptxRenderHandle {
  return paintSlideAs(slideInfoOf(input.model), input);
}

function paintSlideAs(info: SlideInfo, input: PaintSlideInput): PptxRenderHandle {
  const { model, target: container, mediaResolver } = input;
  // A slide opens whole, as a slideshow viewer shows it: fit its container's
  // height as well as its width (falling back to width when the container
  // has no height of its own).
  const viewer = new PptxViewer(input.zoom ?? "fit-page", info);
  const viewport = document.createElement("div");
  viewport.dataset.aaredElement = "slide";
  viewport.style.position = "relative";
  viewport.style.overflow = "hidden";
  viewport.style.background = "white";
  // Theme is viewer chrome state, not content: the slide itself always
  // paints source-faithful. The variable lets embedding UI key off it.
  // Run colors resolve in the Worker (direct fill, then shape/layout/master
  // style fontRef, then placeholder text-style defaults) so unstyled content
  // does not inherit the embedding chrome; no default-ink pin here.
  viewport.dataset.aaredTheme = resolveTheme(input.appearance?.theme);

  const scene = document.createElement("div");
  scene.dataset.aaredElement = "slide-scene";
  scene.style.position = "absolute";
  scene.style.left = "0";
  scene.style.top = "0";
  scene.style.width = inchesToCss(model.size.width);
  scene.style.height = inchesToCss(model.size.height);
  scene.style.transformOrigin = "top left";
  viewport.append(scene);

  // Pictures load asynchronously through `mediaResolver` while painting stays
  // synchronous. The URLs are owned by one `MediaUrls`: a late resolution after
  // `destroy()` paints nothing, and `destroy()` revokes every URL.
  let destroyed = false;
  const urls = mediaResolver === undefined ? undefined : new MediaUrls(mediaResolver);

  paintBackground(viewport, model, urls);
  // Shared layout/master graphics sit under the slide's own content.
  if (model.layout !== undefined && model.hideBackgroundGraphics !== true) {
    paintElements({ scene, slideNumber: model.number, urls, owner: model.layout.id }, model.layout.graphics);
  }
  paintElements({ scene, slideNumber: model.number, urls, owner: model.part }, model.elements);
  // The scene is absolutely positioned, so the viewport has no height of its
  // own until the fit below sizes it. Attached at zero height, the layout the
  // shrink-to-fit pass forces would shrink a page whose height is the slide,
  // and the browser clamps the page's scroll position right then: going to
  // the next slide from the bottom of a page jumped it up. So it goes in at
  // its full-width box — the largest any fit makes it — and the fit sizes it.
  viewport.style.width = "100%";
  viewport.style.aspectRatio = `${model.size.width} / ${model.size.height}`;
  container.replaceChildren(viewport);

  // Shrink-to-fit must measure attached layout (forced reflow). Running it
  // before `replaceChildren` would observe zero sizes on detached nodes.
  for (const text of viewport.querySelectorAll<HTMLElement>('[data-aared-autofit="shrinkText"]')) {
    shrinkTextToFit(text);
  }

  const natural = { width: model.size.width * PX_PER_INCH, height: model.size.height * PX_PER_INCH };
  let lastScale = 1;
  /** Size the viewport to `scale` × the slide; a fit narrower than its container sits in the middle. */
  const paintScale = (scale: number, centred: boolean): number => {
    viewport.style.width = `${natural.width * scale}px`;
    viewport.style.height = `${natural.height * scale}px`;
    viewport.style.aspectRatio = "";
    viewport.style.marginInline = centred ? "auto" : "";
    scene.style.transform = `scale(${scale})`;
    lastScale = scale;
    return scale;
  };
  const applyZoom = (): number => {
    if (destroyed) return lastScale;
    const zoom = viewer.zoom;
    if (typeof zoom === "number") return paintScale(safeScale(zoom), false);
    // The room is measured with the slide set aside (see `measureFitRoom`), so
    // a container that only grows to fit the slide offers no height and
    // fit-page falls back to width — it can never read back its own result.
    const room = measureFitRoom(container, viewport, zoom === "fit-page");
    if (zoom === "fit-page" && room.height > 0) {
      const scale = fitScale(zoom, room, natural);
      if (scale !== undefined) return paintScale(safeScale(scale), true);
    }
    // Fit to width: the box is sized by CSS (full width, the slide's aspect
    // ratio), so only the scene's transform follows the measurement. Sizing it
    // in pixels here would resize the observed container from inside its own
    // resize callback — a loop whenever a page scrollbar comes and goes.
    viewport.style.width = "100%";
    viewport.style.height = "";
    viewport.style.aspectRatio = `${model.size.width} / ${model.size.height}`;
    viewport.style.marginInline = "";
    // Not laid out yet (no width): keep the last scale; the observer re-fits.
    lastScale = room.width > 0 ? safeScale(room.width / natural.width) : lastScale;
    scene.style.transform = `scale(${lastScale})`;
    return lastScale;
  };
  const observer = new ResizeObserver(() => viewer.refit());
  const removalObserver = new MutationObserver(() => {
    if (viewport.parentElement !== container) destroy();
  });
  // Hyperlinks never navigate: prevent the default and report the activation
  // to the embedder instead (same contract as XLSX/DOCX `linkactivate`).
  // Run-level anchors match first, so they win inside shape-level links.
  const onLinkClick = (event: Event): void => {
    const origin = event.target as Element | null;
    const anchor = origin?.closest?.('a[data-aared-element="run-link"], a[data-aared-element="shape-link"]');
    if (!(anchor instanceof HTMLAnchorElement) || !viewport.contains(anchor)) return;
    event.preventDefault();
    viewer.dispatchEvent(new CustomEvent("linkactivate", {
      detail: { target: anchor.getAttribute("href") ?? "" },
    }));
  };
  viewport.addEventListener("click", onLinkClick);
  viewport.addEventListener("auxclick", onLinkClick);
  const destroy = () => {
    if (destroyed) return;
    viewer.markDestroyed();
    destroyed = true;
    urls?.close();
    observer.disconnect();
    removalObserver.disconnect();
    viewport.removeEventListener("click", onLinkClick);
    viewport.removeEventListener("auxclick", onLinkClick);
  };
  const applyAppearance = (appearance: Appearance | undefined): void => {
    if (destroyed) return;
    viewport.dataset.aaredTheme = resolveTheme(appearance?.theme);
  };
  viewer.attach({ applyZoom, destroy, applyAppearance });
  observer.observe(container);
  removalObserver.observe(container, { childList: true });
  viewer.refit();
  return viewer;
}
