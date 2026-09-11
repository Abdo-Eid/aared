/**
 * PPTX renderer: paints a resolved SlideModel to real DOM/SVG — one
 * positioned <div> per shape, inline <svg> for geometry, HTML for text.
 * Never canvas. See AGENTS.md § Rules #1 and docs/ARCHITECTURE.md.
 *
 * Current focused support: text shapes and DrawingML tables from the
 * Arabic gate fixture. Geometry/pictures/fills/rotation join this same
 * DOM model incrementally; no text path ever falls back to canvas.
 */

import type { BoundsModel, ChartGraphicPrimitive, ChartModel, LinearGradientFillModel, LocalViewBoxModel, PictureModel, ShapeModel, SlideModel, TableModel, TextBodyModel, TextParagraphModel } from "../core/model.js";
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
  ViewerTheme,
  ZoomValue,
} from "../core/index.js";
import { AaredError } from "../core/index.js";
import { WorkerSession } from "../core/session.js";
import { tableCellJustifyContent } from "./table.js";

/** Document metadata for a deck; `format` is narrowed to `"pptx"`. */
export interface PptxInfo extends DocumentInfo {
  format: "pptx";
}

/** Lightweight slide identity from `open()` indexing — no slide is parsed yet. */
export interface SlideInfo {
  index: number;
  title: string;
  size: { width: number; height: number };
}

/** One slide result from `deck.search()`. */
export interface PptxSearchMatch {
  /** Zero-based slide index. */
  slide: number;
  /** First matching paragraph text on the slide. */
  text: string;
}

/** What a readonly presentation preview can and cannot do. */
export interface PptxCapabilities {
  textSelection: boolean;
  media: boolean;
  presentation: {
    notes: boolean;
    charts: boolean;
    animations: false;
  };
  editing: false;
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

  read(options?: ReadOptions): Promise<SlideModel> {
    const signal = options?.signal;
    return this.session.unit<SlideModel>(
      this.info.index,
      signal === undefined ? undefined : { signal },
    );
  }

  /** Picture bytes for this slide's `mediaId`s, via the owning session. */
  media(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.media(mediaId, options);
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
    const handle = this.handles[index];
    if (!Number.isInteger(index) || handle === undefined) {
      throw new AaredError(
        "not-found",
        `aared/pptx: slide index ${String(index)} is outside this ${this.handles.length}-slide deck`,
      );
    }
    return handle;
  }
}

class DeckMediaCollection implements MediaCollectionShape {
  constructor(private readonly session: WorkerSession) {}

  get(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.media(mediaId, options);
  }
}

function slideText(model: SlideModel): string[] {
  const paragraphs: string[] = [];
  const bodyText = (body: TextBodyModel): void => {
    for (const paragraph of body.paragraphs) {
      const text = paragraph.runs.map((run) => run.text).join("").trim();
      if (text !== "") paragraphs.push(text);
    }
  };
  for (const element of model.elements) {
    if (element.kind === "shape") {
      if (element.text !== undefined) bodyText(element.text);
    } else if (element.kind === "table") {
      for (const row of element.rows) {
        for (const cell of row.cells) bodyText(cell.text);
      }
    }
  }
  return paragraphs;
}

/** An opened presentation: collections, search, media, and lifecycle. */
export class PptxDocument {
  readonly info: PptxInfo;
  readonly slides: SlideCollection;
  readonly media: MediaCollectionShape;
  readonly capabilities: PptxCapabilities = {
    textSelection: true,
    media: true,
    presentation: { notes: true, charts: true, animations: false },
    editing: false,
  };

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
    this.media = new DeckMediaCollection(session);
  }

  /**
   * Case-insensitive substring search across slides. Reads each slide in
   * order (aborting stops at the next slide); one match per slide, the
   * first matching paragraph.
   */
  async search(query: string, options?: ReadOptions): Promise<PptxSearchMatch[]> {
    const needle = query.toLowerCase();
    if (needle === "") return [];
    const matches: PptxSearchMatch[] = [];
    for (const handle of this.slides.list()) {
      const signal = options?.signal;
      const model = await handle.read(signal === undefined ? undefined : { signal });
      const hit = slideText(model).find((text) => text.toLowerCase().includes(needle));
      if (hit !== undefined) matches.push({ slide: handle.info.index, text: hit.slice(0, 120) });
    }
    return matches;
  }

  /** Terminate the Worker session and reject pending reads. */
  close(): void {
    this.session.close();
  }
}

/**
 * Open a `.pptx` file. Fails with an `unsupported-format` `AaredError`
 * when the file is not actually a presentation.
 */
export async function open(source: Source, options?: OpenOptions): Promise<PptxDocument> {
  const session = await WorkerSession.open("aared/pptx", source, options);
  if (session.info.format !== "pptx") {
    const detected = session.info.format;
    session.close();
    throw new AaredError(
      "unsupported-format",
      `aared/pptx: expected a PowerPoint presentation; detected "${detected}" — open it with aared/${detected} instead`,
    );
  }
  return new PptxDocument(session);
}

/** Viewer-only appearance for PPTX. Content is always source-faithful. */
export interface PptxAppearance {
  theme?: ViewerTheme;
  /** PPTX supports `source` only: a slide is a designed canvas, never recolored. */
  contentMode?: "source";
}

/** Main-thread resources available to the PPTX renderer. */
/** Internal options for picture rendering (shares fields with PptxRenderInput). */
interface PptxRenderOptions {
  /** Main-thread media boundary for picture rendering; missing resolver renders a marked fallback. */
  mediaResolver?: MediaResolver;
  appearance?: PptxAppearance | Appearance;
}

/** PowerPoint's 96-DPI PNG export and CSS pixels share this layout baseline. */
const PIXELS_PER_INCH = 96;
const MINIMUM_LINE_HEIGHT = 1.1;

export interface PptxRenderHandle extends EventTarget {
  /** Disconnect layout observers after this slide is no longer displayed. */
  destroy(): void;
  /** Change zoom; emits `zoomchange`. Numbers scale the 96-DPI scene. */
  setZoom(zoom: ZoomValue): void;
  /** The active zoom. */
  readonly zoom: ZoomValue;
}

/** `render()` input: a slide handle (read on demand) or an already-resolved model. */
export interface PptxRenderInput {
  slide: SlideHandle | SlideModel;
  target: HTMLElement;
  appearance?: PptxAppearance | Appearance;
  mediaResolver?: MediaResolver;
}

function isSlideHandle(slide: SlideHandle | SlideModel): slide is SlideHandle {
  return slide instanceof SlideHandle;
}

class PptxViewer extends EventTarget implements PptxRenderHandle {
  #zoom: ZoomValue = "fit-width";
  #destroyed = false;
  #applyZoom: () => void = () => {};
  #destroyHook: () => void = () => {};

  get zoom(): ZoomValue {
    return this.#zoom;
  }

  setZoom(zoom: ZoomValue): void {
    this.#zoom = zoom;
    this.#applyZoom();
    this.dispatchEvent(new CustomEvent("zoomchange", { detail: { zoom } }));
  }

  destroy(): void {
    this.#destroyHook();
  }

  /** Wired by `paintSlide()`; `setZoom()` before wiring applies on first paint. */
  attach(applyZoom: () => void, destroy: () => void): void {
    this.#applyZoom = applyZoom;
    this.#destroyHook = destroy;
  }

  markDestroyed(): void {
    this.#destroyed = true;
  }

  get destroyed(): boolean {
    return this.#destroyed;
  }
}

/**
 * Paint one slide. Accepts a `SlideHandle` (read from the Worker on
 * demand) or an already-resolved `SlideModel`; when the handle carries
 * its document session, picture media resolves through it unless an
 * explicit `mediaResolver` is given. The caller owns virtualization.
 */
export async function render(input: PptxRenderInput): Promise<PptxRenderHandle> {
  const { target } = input;
  const appearance = input.appearance;
  if (appearance?.contentMode !== undefined && appearance.contentMode !== "source") {
    throw new AaredError(
      "unsupported-feature",
      `aared/pptx: contentMode "adaptive" is not supported for slides; a slide is a designed canvas and always renders source-faithful`,
    );
  }
  const requested = input.slide;
  let mediaResolver = input.mediaResolver;
  let model: SlideModel;
  if (isSlideHandle(requested)) {
    model = await requested.read();
    mediaResolver ??= (mediaId: string) => requested.media(mediaId);
  } else {
    model = requested;
  }
  return paintSlide(model, target, mediaResolver, appearance?.theme);
}

function inches(value: number): string {
  return `${value * PIXELS_PER_INCH}px`;
}

function textAlign(paragraph: TextParagraphModel): string {
  return paragraph.align ?? (paragraph.direction === "rtl" ? "right" : "left");
}

function appendParagraph(
  host: HTMLElement,
  paragraph: TextParagraphModel,
  textWidth?: number,
  wrapping?: TextBodyModel["wrapping"],
): void {
  const line = document.createElement("p");
  line.style.margin = "0";
  line.style.minWidth = "0";
  line.style.whiteSpace = wrapping === "none" ? "nowrap" : "pre-wrap";
  if (paragraph.fontSizePt !== undefined) line.style.fontSize = `${paragraph.fontSizePt}pt`;
  if (paragraph.lineSpacing !== undefined) {
    // DrawingML permits a baseline distance smaller than an Arabic glyph's ink
    // bounds. CSS would crop a single line in that case, so retain the source
    // value on the model but use the smallest safe browser line box.
    line.style.lineHeight = String(Math.max(paragraph.lineSpacing, MINIMUM_LINE_HEIGHT));
    line.dataset.aaredLineSpacing = String(paragraph.lineSpacing);
  }
  if (paragraph.lineSpacingPt !== undefined) line.style.lineHeight = `${paragraph.lineSpacingPt}pt`;
  if (paragraph.spaceBeforePt !== undefined) line.style.marginTop = `${paragraph.spaceBeforePt}pt`;
  if (paragraph.spaceAfterPt !== undefined) line.style.marginBottom = `${paragraph.spaceAfterPt}pt`;
  if (paragraph.direction !== undefined) {
    line.style.direction = paragraph.direction;
    line.dataset.aaredDirection = paragraph.direction;
  }
  line.style.textAlign = textAlign(paragraph);
  line.dataset.aaredAlign = textAlign(paragraph);
  // Indents apply in shapes and table cells alike; the bullet grid already
  // supplies hanging marker/text separation for bulleted paragraphs.
  if (paragraph.leftIndent !== undefined) line.style.marginLeft = inches(paragraph.leftIndent);
  if (paragraph.rightIndent !== undefined) line.style.marginRight = inches(paragraph.rightIndent);
  if (paragraph.bullet === undefined && paragraph.firstLineIndent !== undefined) {
    line.style.textIndent = inches(paragraph.firstLineIndent);
  }

  const text = document.createElement("span");
  text.dataset.aaredElement = "paragraph-text";
  text.style.minWidth = "0";
  for (const run of paragraph.runs) {
    const span = document.createElement("span");
    span.textContent = run.text;
    if (run.bold === true) span.style.fontWeight = "700";
    if (run.italic === true) span.style.fontStyle = "italic";
    if (run.fontSizePt !== undefined) span.style.fontSize = `${run.fontSizePt}pt`;
    if (run.color !== undefined) span.style.color = run.color;
    if (run.fontFamily !== undefined) span.style.fontFamily = run.fontFamily;
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
    if (textAlign(paragraph) === "left") {
      // A left-aligned RTL bullet occupies only its content width so its marker
      // follows the text instead of staying at the text box's far-right edge.
      line.style.width = "fit-content";
      line.style.alignSelf = "flex-start";
    }
    const marker = document.createElement("span");
    marker.dataset.aaredElement = "bullet-marker";
    marker.textContent = paragraph.bullet.char ?? "•";
    marker.setAttribute("aria-hidden", "true");
    const fontSize = paragraph.runs.find((run) => run.fontSizePt !== undefined)?.fontSizePt;
    if (fontSize !== undefined) marker.style.fontSize = `${fontSize}pt`;
    if (paragraph.bullet.typeface !== undefined) marker.style.fontFamily = paragraph.bullet.typeface;
    if (paragraph.direction !== undefined) marker.dataset.aaredBulletSide = paragraph.direction === "rtl" ? "right" : "left";
    line.append(marker, text);
  }
  host.append(line);
}

function appendTextBody(host: HTMLElement, body: TextBodyModel, textWidth?: number): void {
  for (const paragraph of body.paragraphs) appendParagraph(host, paragraph, textWidth, body.wrapping);
}

function appendShapeText(host: HTMLElement, shape: ShapeModel): void {
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
    if (body.insets.left !== undefined) text.style.paddingLeft = inches(body.insets.left);
    if (body.insets.right !== undefined) text.style.paddingRight = inches(body.insets.right);
    if (body.insets.top !== undefined) text.style.paddingTop = inches(body.insets.top);
    if (body.insets.bottom !== undefined) text.style.paddingBottom = inches(body.insets.bottom);
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
  appendTextBody(text, body, bounds?.width);
  host.append(text);
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

function numberText(value: number): string {
  return `${Math.round(value * 1_000_000) / 1_000_000}`;
}

function gradientAngle(fill: LinearGradientFillModel, shape: ShapeModel): number {
  if (fill.rotatesWithShape) return fill.angle;
  let angle = fill.angle - (shape.transform?.rotation ?? 0);
  if (shape.transform?.flipHorizontal === true) angle = 180 - angle;
  if (shape.transform?.flipVertical === true) angle = -angle;
  return angle;
}

function appendLinearGradient(svg: SVGSVGElement, id: string, fill: LinearGradientFillModel, shape: ShapeModel, viewBox: LocalViewBoxModel): void {
  const angle = (gradientAngle(fill, shape) * Math.PI) / 180;
  const horizontal = Math.cos(angle);
  const vertical = Math.sin(angle);
  const extent = fill.scaled
    ? (Math.abs(horizontal) + Math.abs(vertical)) / 2
    : (Math.abs(horizontal) * viewBox.width + Math.abs(vertical) * viewBox.height) / 2;
  const centerX = fill.scaled ? 0.5 : viewBox.width / 2;
  const centerY = fill.scaled ? 0.5 : viewBox.height / 2;
  const gradient = document.createElementNS("http://www.w3.org/2000/svg", "linearGradient");
  gradient.setAttribute("id", id);
  if (!fill.scaled) gradient.setAttribute("gradientUnits", "userSpaceOnUse");
  gradient.setAttribute("x1", numberText(centerX - horizontal * extent));
  gradient.setAttribute("y1", numberText(centerY - vertical * extent));
  gradient.setAttribute("x2", numberText(centerX + horizontal * extent));
  gradient.setAttribute("y2", numberText(centerY + vertical * extent));
  for (const stopModel of fill.stops) {
    const stop = document.createElementNS("http://www.w3.org/2000/svg", "stop");
    stop.setAttribute("offset", `${numberText(stopModel.position * 100)}%`);
    stop.setAttribute("stop-color", stopModel.color);
    gradient.append(stop);
  }
  const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");
  defs.append(gradient);
  svg.append(defs);
}

function appendShapeGeometry(host: HTMLElement, shape: ShapeModel, gradientId: string): void {
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
  if (fallback !== undefined) {
    // Unsupported and omitted DrawingML geometry remains a visible rectangle,
    // rather than dropping its fill, outline, or text from the slide.
    host.dataset.aaredGeometryFallback = fallback;
  }
  if (geometry === undefined) {
    svg.setAttribute("viewBox", "0 0 100000 100000");
  } else {
    svg.setAttribute("viewBox", `0 0 ${geometry.viewBox.width} ${geometry.viewBox.height}`);
  }
  let paint = "none";
  if (shape.fill?.kind === "solid") paint = shape.fill.color;
  else if (shape.fill?.kind === "linear-gradient") {
    appendLinearGradient(svg, gradientId, shape.fill, shape, viewBox);
    paint = `url(#${gradientId})`;
  } else if (shape.fill?.kind === "unsupported") {
    host.dataset.aaredFillFallback = shape.fill.source;
  }
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", geometry?.path ?? "M 0 0 H 100000 V 100000 H 0 Z");
  path.setAttribute("fill", paint);
  if (shape.stroke !== undefined) {
    path.setAttribute("stroke", shape.stroke.color);
    const width = shape.bounds === undefined || shape.bounds.width === 0
      ? 1
      : (shape.stroke.width / shape.bounds.width) * (geometry?.viewBox.width ?? 100000);
    path.setAttribute("stroke-width", `${width}`);
  }
  svg.append(path);
  host.append(svg);
}

function position(bounds: BoundsModel | undefined, element: HTMLElement): void {
  if (bounds !== undefined) {
    element.style.left = inches(bounds.x);
    element.style.top = inches(bounds.y);
    element.style.width = inches(bounds.width);
    element.style.height = inches(bounds.height);
  } else {
    element.style.inset = "0";
  }
}

function gradientId(slide: SlideModel, shape: ShapeModel, index: number): string {
  let hash = 0;
  const part = typeof slide.part === "string" ? slide.part : "inline";
  for (const char of part) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return `aared-gradient-${hash.toString(36)}-${shape.id}-${index}`;
}

function appendTable(scene: HTMLElement, table: TableModel): void {
  const grid = document.createElement("div");
  grid.setAttribute("role", "table");
  grid.dataset.aaredElement = "table";
  grid.dataset.aaredTableDirection = table.direction;
  grid.style.position = "absolute";
  grid.style.display = "grid";
  grid.style.direction = table.direction;
  grid.style.overflow = "hidden";
  grid.style.transformOrigin = "center";
  const tableTransforms: string[] = [];
  if (table.transform?.rotation !== undefined) tableTransforms.push(`rotate(${table.transform.rotation}deg)`);
  if (table.transform?.flipHorizontal === true || table.transform?.flipVertical === true) {
    tableTransforms.push(`scale(${table.transform.flipHorizontal === true ? -1 : 1}, ${table.transform.flipVertical === true ? -1 : 1})`);
  }
  if (tableTransforms.length > 0) grid.style.transform = tableTransforms.join(" ");
  grid.style.border = "1px solid currentColor";
  if (table.bounds !== undefined) {
    grid.style.left = inches(table.bounds.x);
    grid.style.top = inches(table.bounds.y);
    grid.style.width = inches(table.bounds.width);
    grid.style.height = inches(table.bounds.height);
  }
  const totalColumnWidth = table.columns.reduce((sum, width) => sum + width, 0);
  grid.style.gridTemplateColumns = totalColumnWidth > 0
    ? table.columns.map((width) => `${(width / totalColumnWidth) * 100}%`).join(" ")
    : table.columns.map(() => "1fr").join(" ");
  const rowHeights = table.rows.map((row) => row.height);
  if (rowHeights.every((height): height is number => height !== undefined)) {
    const totalRowHeight = rowHeights.reduce((sum, height) => sum + height, 0);
    if (totalRowHeight > 0) {
      grid.style.gridTemplateRows = rowHeights.map((height) => `${(height / totalRowHeight) * 100}%`).join(" ");
    }
  }
  for (const [rowIndex, row] of table.rows.entries()) {
    for (const [columnIndex, cellModel] of row.cells.entries()) {
      const cell = document.createElement("div");
      cell.setAttribute("role", "cell");
      cell.dataset.aaredElement = "table-cell";
      cell.style.minWidth = "0";
      cell.style.minHeight = "0";
      cell.style.boxSizing = "border-box";
      // Resolved DrawingML cell insets win; otherwise keep the legacy padding
      // so unstyled tables (e.g. the Arabic gate fixture) do not collapse.
      if (cellModel.margins !== undefined) {
        cell.style.paddingTop = cellModel.margins.top !== undefined ? inches(cellModel.margins.top) : "0px";
        cell.style.paddingRight = cellModel.margins.right !== undefined ? inches(cellModel.margins.right) : "0px";
        cell.style.paddingBottom = cellModel.margins.bottom !== undefined ? inches(cellModel.margins.bottom) : "0px";
        cell.style.paddingLeft = cellModel.margins.left !== undefined ? inches(cellModel.margins.left) : "0px";
        cell.dataset.aaredCellMargins = "resolved";
      } else {
        cell.style.padding = "0.2em 0.35em";
      }
      // Resolved per-edge strokes win; otherwise keep the faint grid fallback.
      if (cellModel.borders !== undefined) {
        const edge = (side: "Top" | "Right" | "Bottom" | "Left", model: { color: string; width: number } | undefined): void => {
          if (model === undefined) {
            (cell.style as unknown as Record<string, string>)[`border${side}`] = "none";
          } else {
            (cell.style as unknown as Record<string, string>)[`border${side}`] = `${inches(model.width)} solid ${model.color}`;
          }
        };
        edge("Top", cellModel.borders.top);
        edge("Right", cellModel.borders.right);
        edge("Bottom", cellModel.borders.bottom);
        edge("Left", cellModel.borders.left);
        cell.dataset.aaredCellBorders = "resolved";
      } else {
        cell.style.border = "1px solid color-mix(in srgb, currentColor 25%, transparent)";
      }
      if (cellModel.columnSpan !== undefined) {
        cell.style.gridColumn = `span ${cellModel.columnSpan}`;
        cell.dataset.aaredColumnSpan = String(cellModel.columnSpan);
      }
      if (cellModel.rowSpan !== undefined) {
        cell.style.gridRow = `span ${cellModel.rowSpan}`;
        cell.dataset.aaredRowSpan = String(cellModel.rowSpan);
      }
      if (cellModel.verticalAlign !== undefined) {
        cell.style.display = "flex";
        cell.style.flexDirection = "column";
        cell.style.justifyContent = tableCellJustifyContent(cellModel.verticalAlign);
      }
      const fill = table.cellFills[rowIndex]?.[columnIndex];
      if (fill !== undefined) cell.style.backgroundColor = fill;
      appendTextBody(cell, cellModel.text);
      grid.append(cell);
    }
  }
  scene.append(grid);
}

interface PictureRenderState {
  generation: number;
  destroyed: boolean;
  objectUrls: string[];
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
  fallback.textContent = picture.name === "" ? "Image unavailable" : picture.name;
  host.append(fallback);
}

function appendPicture(
  scene: HTMLElement,
  picture: PictureModel,
  options: PptxRenderOptions | undefined,
  state: PictureRenderState,
): void {
  const wrapper = document.createElement("div");
  wrapper.dataset.aaredElement = "picture";
  wrapper.dataset.aaredPictureId = picture.id;
  wrapper.dataset.aaredMediaId = picture.mediaId;
  wrapper.dataset.aaredContentType = picture.contentType;
  wrapper.style.position = "absolute";
  wrapper.style.overflow = "hidden";
  wrapper.style.transformOrigin = "center";
  wrapper.style.background = "transparent";
  const transform = transformFor(picture.transform);
  if (transform !== "") wrapper.style.transform = transform;
  position(picture.bounds, wrapper);
  if (picture.crop !== undefined) {
    wrapper.dataset.aaredCropLeft = String(picture.crop.left);
    wrapper.dataset.aaredCropTop = String(picture.crop.top);
    wrapper.dataset.aaredCropRight = String(picture.crop.right);
    wrapper.dataset.aaredCropBottom = String(picture.crop.bottom);
  }

  const resolver = options?.mediaResolver;
  if (resolver === undefined) {
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

  const generation = state.generation;
  let settled = false;
  Promise.resolve()
    .then(() => resolver(picture.mediaId))
    .then((bytes) => {
      if (settled || state.destroyed || state.generation !== generation || !wrapper.isConnected) return;
      settled = true;
      const blob = new Blob([bytes as BlobPart], { type: picture.contentType });
      const url = URL.createObjectURL(blob);
      state.objectUrls.push(url);
      if (state.destroyed || state.generation !== generation || !image.isConnected) {
        URL.revokeObjectURL(url);
        return;
      }
      image.src = url;
      image.dataset.aaredMediaLoaded = "true";
    })
    .catch(() => {
      if (settled || state.destroyed || state.generation !== generation || !wrapper.isConnected) return;
      settled = true;
      image.remove();
      appendPictureFallback(wrapper, picture, "load-error");
    });
}

function appendChartPrimitive(svg: SVGSVGElement, primitive: ChartGraphicPrimitive): void {
  const svgNamespace = "http://www.w3.org/2000/svg";
  const applyStyle = (element: SVGElement, styled: Extract<ChartGraphicPrimitive, { fill?: string }>): void => {
    if (styled.fill !== undefined) element.setAttribute("fill", styled.fill);
    if (styled.stroke !== undefined) {
      element.setAttribute("stroke", styled.stroke.color);
      element.setAttribute("stroke-width", String(styled.stroke.width));
    }
    if (styled.opacity !== undefined) element.setAttribute("opacity", String(styled.opacity));
  };
  switch (primitive.kind) {
    case "rect": {
      const rect = document.createElementNS(svgNamespace, "rect");
      rect.setAttribute("x", String(primitive.x));
      rect.setAttribute("y", String(primitive.y));
      rect.setAttribute("width", String(primitive.width));
      rect.setAttribute("height", String(primitive.height));
      if (primitive.fill !== undefined) rect.setAttribute("fill", primitive.fill);
      else rect.setAttribute("fill", "none");
      applyStyle(rect, primitive);
      svg.append(rect);
      break;
    }
    case "line": {
      const line = document.createElementNS(svgNamespace, "line");
      line.setAttribute("x1", String(primitive.x1));
      line.setAttribute("y1", String(primitive.y1));
      line.setAttribute("x2", String(primitive.x2));
      line.setAttribute("y2", String(primitive.y2));
      if (primitive.stroke !== undefined) {
        line.setAttribute("stroke", primitive.stroke.color);
        line.setAttribute("stroke-width", String(primitive.stroke.width));
      } else {
        line.setAttribute("stroke", "currentColor");
      }
      if (primitive.opacity !== undefined) line.setAttribute("opacity", String(primitive.opacity));
      svg.append(line);
      break;
    }
    case "path": {
      const path = document.createElementNS(svgNamespace, "path");
      path.setAttribute("d", primitive.path);
      if (primitive.fill !== undefined) path.setAttribute("fill", primitive.fill);
      else path.setAttribute("fill", "none");
      applyStyle(path, primitive);
      svg.append(path);
      break;
    }
    case "text": {
      const text = document.createElementNS(svgNamespace, "text");
      text.setAttribute("x", String(primitive.x));
      text.setAttribute("y", String(primitive.y));
      text.textContent = primitive.text;
      if (primitive.fontSize !== undefined) text.setAttribute("font-size", String(primitive.fontSize));
      if (primitive.fontFamily !== undefined) text.setAttribute("font-family", primitive.fontFamily);
      if (primitive.fontWeight !== undefined) text.setAttribute("font-weight", primitive.fontWeight);
      if (primitive.textAnchor !== undefined) text.setAttribute("text-anchor", primitive.textAnchor);
      text.setAttribute("fill", primitive.fill ?? "#000000");
      if (primitive.stroke !== undefined) {
        text.setAttribute("stroke", primitive.stroke.color);
        text.setAttribute("stroke-width", String(primitive.stroke.width));
      }
      if (primitive.opacity !== undefined) text.setAttribute("opacity", String(primitive.opacity));
      svg.append(text);
      break;
    }
  }
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

/** Paint one already-resolved slide. The caller owns virtualization. */
export function paintSlide(
  model: SlideModel,
  container: HTMLElement,
  mediaResolver?: MediaResolver,
  theme?: ViewerTheme,
): PptxRenderHandle {
  const viewer = new PptxViewer();
  const viewport = document.createElement("div");
  viewport.dataset.aaredElement = "slide";
  viewport.style.position = "relative";
  viewport.style.overflow = "hidden";
  viewport.style.background = "white";
  // Theme is viewer chrome state, not content: the slide itself always
  // paints source-faithful. The variable lets embedding UI key off it.
  viewport.dataset.aaredTheme = theme ?? "system";

  const scene = document.createElement("div");
  scene.dataset.aaredElement = "slide-scene";
  scene.style.position = "absolute";
  scene.style.left = "0";
  scene.style.top = "0";
  scene.style.width = inches(model.size.width);
  scene.style.height = inches(model.size.height);
  scene.style.transformOrigin = "top left";
  viewport.append(scene);

  // Pictures load asynchronously through `mediaResolver` while painting
  // stays synchronous. The generation guard drops stale resolutions when a
  // slide is replaced, and `destroy()` revokes every created object URL.
  const pictureState: PictureRenderState = { generation: 0, destroyed: false, objectUrls: [] };
  const paintOptions = mediaResolver === undefined ? undefined : { mediaResolver };

  for (const [index, element] of model.elements.entries()) {
    if (element.kind === "table") {
      appendTable(scene, element);
      continue;
    }
    if (element.kind === "picture") {
      appendPicture(scene, element, paintOptions, pictureState);
      continue;
    }
    if (element.kind === "chart") {
      appendChart(scene, element);
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
    appendShapeGeometry(shape, element, gradientId(model, element, index));
    appendShapeText(shape, element);
    scene.append(shape);
  }
  container.replaceChildren(viewport);

  // Shrink-to-fit must measure attached layout (forced reflow). Running it
  // before `replaceChildren` would observe zero sizes on detached nodes.
  for (const text of viewport.querySelectorAll<HTMLElement>('[data-aared-autofit="shrinkText"]')) {
    shrinkTextToFit(text);
  }

  const naturalWidth = model.size.width * PIXELS_PER_INCH;
  const naturalHeight = model.size.height * PIXELS_PER_INCH;
  const applyZoom = () => {
    if (pictureState.destroyed) return;
    const zoom = viewer.zoom;
    // Computed sizes stay in the element's own CSS coordinate space,
    // unlike a transformed bounding rect inherited from an embedder.
    // getBoundingClientRect (not clientWidth) preserves fractional pixels,
    // so a 640.5px container scales by exactly 640.5/1280.
    const availableWidth = container.getBoundingClientRect().width || naturalWidth;
    const availableHeight = container.getBoundingClientRect().height || naturalHeight;
    if (typeof zoom === "number") {
      const scale = zoom > 0 && Number.isFinite(zoom) ? zoom : 1;
      viewport.style.width = `${naturalWidth * scale}px`;
      viewport.style.height = `${naturalHeight * scale}px`;
      viewport.style.aspectRatio = "";
      scene.style.transform = `scale(${scale})`;
    } else if (zoom === "fit-page" && availableHeight > 0) {
      const scale = Math.min(availableWidth / naturalWidth, availableHeight / naturalHeight);
      const safe = Number.isFinite(scale) && scale > 0 ? scale : 1;
      viewport.style.width = `${naturalWidth * safe}px`;
      viewport.style.height = `${naturalHeight * safe}px`;
      viewport.style.aspectRatio = "";
      scene.style.transform = `scale(${safe})`;
    } else {
      viewport.style.width = "100%";
      viewport.style.height = "";
      viewport.style.aspectRatio = `${model.size.width} / ${model.size.height}`;
      const scale = availableWidth / naturalWidth;
      scene.style.transform = `scale(${Number.isFinite(scale) && scale > 0 ? scale : 1})`;
    }
  };
  const observer = new ResizeObserver(() => applyZoom());
  const removalObserver = new MutationObserver(() => {
    if (viewport.parentElement !== container) destroy();
  });
  const destroy = () => {
    if (pictureState.destroyed) return;
    viewer.markDestroyed();
    pictureState.destroyed = true;
    pictureState.generation += 1;
    for (const url of pictureState.objectUrls) {
      try {
        URL.revokeObjectURL(url);
      } catch {
        // Revocation is best-effort; a failed revoke cannot affect rendering.
      }
    }
    pictureState.objectUrls.length = 0;
    observer.disconnect();
    removalObserver.disconnect();
  };
  viewer.attach(applyZoom, destroy);
  observer.observe(container);
  removalObserver.observe(container, { childList: true });
  applyZoom();
  return viewer;
}
