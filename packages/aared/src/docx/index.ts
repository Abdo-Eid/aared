/**
 * DOCX entrypoint (`aared/docx`): open Word documents in a Worker and paint
 * resolved `DocumentSection` flow blocks to HTML with correct
 * paragraph-level RTL. Never canvas. See AGENTS.md Rules #1 and
 * docs/ARCHITECTURE.md.
 *
 * `open()` resolves every section eagerly with one Worker `unit(0)` call:
 * a DOCX section is source/document structure, not a physical page, so
 * sections resolve together and are consumed together by rendering/layout —
 * there is deliberately no `section.read()`. Physical pages belong to the
 * render result only. Flow mode (the default) renders every section as one
 * continuous article with the last section's header/footer, matching the
 * pre-pagination output; layout mode (`pagination: { mode: "layout" }`)
 * paginates into per-page DOM. The returned `DocxRenderHandle` owns zoom,
 * heading navigation, and hyperlink activation events.
 *
 * Source-faithful content (like PPTX): pages always paint as authored —
 * white surface, black default ink. Automatic text over an authored shading
 * takes the ink that reads on it. The resolved `theme` affects viewer chrome
 * only; the document stays light in dark mode, like a slide.
 */

import type {
  DocxHeadingNode,
  DocxImageNode,
  DocxInlineImage,
  DocxListItemNode,
  DocxPageBreakNode,
  DocxParagraphBorder,
  DocxParagraphLayout,
  DocxParagraphNode,
  DocxTableEdge,
  DocxTableNode,
  DocxUnsupportedNode,
  DocumentSection,
  DocxTableRowNode,
  FlowNode,
  RenderWarning,
  TextDirection,
  TextRunModel,
} from "../core/model.js";
import type {
  Appearance,
  Collection,
  DocumentInfo,
  MediaCollection as MediaCollectionShape,
  MediaResolver,
  OpenOptions,
  ReadOptions,
  Source,
  ZoomValue,
} from "../core/index.js";
import { AaredError } from "../core/index.js";
import { RETURN_ARROW } from "../core/chars.js";
import { SessionMedia, WorkerSession } from "../core/session.js";
import { itemAt, throwIfAborted } from "../core/errors.js";
import { BASE_FONT_PX, BLOCK_MARGIN_EM, WORD_SINGLE_LINE, dropsTopSpacing, paginate, sectionBodyWidthPx } from "../core/docx/pagination.js";
import { planFloat } from "./float-layout.js";
import type { AnchorFrame } from "./float-layout.js";
import { registerSplit, settleSplitWraps, splittable } from "./split-wrap.js";
import { cssFontFamily, lineFontSizePt } from "../core/typography.js";
import type { LayoutPage, MeasuredFlow, SoftBreakPolicy } from "../core/docx/pagination.js";
import { automaticInk, paintTokens, resolveTheme } from "../appearance.js";
import { MediaUrls } from "../media-url.js";
import { RenderHandleBase, fitScale, measureFitRoom, safeScale } from "../render-handle.js";
import { createTableCell, createTableGrid } from "../table-dom.js";
import { applyParagraphBox, applyRunBase, exactLinePitch, roundPt } from "../text-style.js";
import { awaitDocumentFonts, fontFidelityWarnings } from "../font-fidelity.js";
import { PX_PER_INCH, inchesToCss, ptToPx, twipsToInches } from "../core/units.js";
import { fillRunText, lazyPagePaint, markTabs, settlePageTabStops, settleTabStops } from "./tabs.js";
import { adoptClassSheet, cssClass, type ClassSheetScope } from "../class-sheet.js";
import { findByTitle, foldQuery } from "../core/search.js";
import { buildDocxSearchIndex, entryMatch } from "./search.js";
import type { DocxSearchEntry } from "./search.js";
import { headingOutline } from "./outline.js";
import type { HeadingOutline } from "./outline.js";

export type { SoftBreakPolicy } from "../core/docx/pagination.js";

/** Hanging indent, in inches, for a list item whose numbering level authors none. */
const DEFAULT_HANGING_IN = 0.25;

// Defaults every paragraph, heading and list marker shares, as class rules in
// the render root's one adopted sheet (class-sheet.ts); inline style keeps
// only what a node authors (its margins, indents, alignment, runs).
/** A paragraph or heading block: authored spacing and indents override the zero margin inline. */
// Headings paint as <h1>-<h6> for outline semantics, so the class also resets
// the browser's heading bold and 2em size: weight and size come from runs, as
// in Word (a Heading 1 without `w:b` is not bold).
const BLOCK = cssClass("margin:0;min-width:0;white-space:pre-wrap;font-weight:normal;font-size:inherit");
/**
 * A justified paragraph Word lays out with its 2013 engine (`justifyShrink`).
 * Word fits one more word on a line by taking back part of each space, up to
 * about a quarter of a space and a third of the word it fits (probe
 * docx-justify). CSS breaks lines at fixed widths, so its spaces break a
 * little narrower here, and justification stretches every full line back to
 * the margin. The amount stays under both of Word's limits for a line of up
 * to about twenty spaces, so a word Word pushes to the next line is never
 * pulled back; it covers the near misses where a line runs a pixel or less
 * over (a justified Arabic cell broke "تيار" to a fourth line for 0.6px).
 * Only the last line keeps the slightly narrower spaces.
 */
const JUSTIFY_SHRINK = cssClass("word-spacing:-0.01em");
const MARKER = cssClass("text-indent:0");
/** A list marker followed by a tab to the text indent: it owns the hanging width (`min-width`, inline). */
const MARKER_TAB = cssClass("text-indent:0;display:inline-block;text-align:start");
/** A laid-out page; its size and padding stay inline. */
const PAGE = cssClass(
  `flex-shrink:0;box-sizing:border-box;position:relative;isolation:isolate;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);font-family:system-ui, sans-serif;font-size:${BASE_FONT_PX}px;line-height:${WORD_SINGLE_LINE}`,
);
const PAGE_LAZY = cssClass("content-visibility:auto");

/** Document metadata for a Word file; `format` is narrowed to `"docx"`. */
export interface DocxInfo extends DocumentInfo {
  format: "docx";
}

/**
 * One outline entry: a body heading, in document order. `index` is its
 * position in `document.headings.list()` and the key `goToHeading()` takes;
 * headings inside headers, footers, tables or notes are not in the outline.
 */
export interface HeadingInfo {
  index: number;
  title: string;
  level: number;
  /** Zero-based section holding the heading. */
  section: number;
}

/** One laid-out page (layout mode). `index` is zero-based — the key `goToPage()` takes. */
export interface DocxPageInfo {
  index: number;
  /** In CSS pixels at 96 DPI. */
  width: number;
  height: number;
  orientation: "portrait" | "landscape";
}

/** One document result from `document.search()`. */
export interface DocxSearchMatch {
  /** What holds the match: a body block, a table-cell block, or a note body. */
  kind: "heading" | "paragraph" | "list" | "table" | "footnote" | "endnote";
  /**
   * Zero-based section index holding the match. For a note, the section of
   * the paragraph that first references it.
   */
  section: number;
  /**
   * Nearest preceding body heading in the same section, if any — its
   * `index` is the `goToHeading()` key, so duplicate headings stay
   * unambiguous. For a note, the heading in force where it is first referenced.
   */
  heading: { index: number; title: string } | null;
  /** Matching block text, sliced to 120 characters. */
  text: string;
  /**
   * Table-cell matches only: the table's document-order index within the
   * section (nested tables count too) and the zero-based row and logical
   * column of the innermost table.
   */
  table?: { index: number; row: number; col: number };
  /** Note matches only: the note's id (`f`/`e`-prefixed) and display label. */
  note?: { id: string; label: string };
}

/**
 * `document.sections`: addressable sections in document order. A DOCX
 * section is source structure, not a physical page, so sections resolve
 * together at `open()` time and this collection reads synchronously from
 * that cache — there is deliberately no `section.read()`.
 */
export class SectionCollection implements Collection<DocumentSection, number> {
  constructor(private readonly sections: readonly DocumentSection[]) {}

  list(): readonly DocumentSection[] {
    return this.sections;
  }

  get(index: number): DocumentSection {
    return itemAt(
      this.sections,
      index,
      () => `aared/docx: section index ${String(index)} is outside this ${this.sections.length}-section document`,
    );
  }
}

/** `document.headings`: the body heading outline in document order. */
export class HeadingCollection implements Collection<HeadingInfo, number | string> {
  constructor(private readonly outline: HeadingOutline) {}

  list(): readonly HeadingInfo[] {
    return this.outline.list;
  }

  /** By index, or by title (exact, then ignoring case — never a substring). */
  get(key: number | string): HeadingInfo {
    const list = this.outline.list;
    if (typeof key === "number") {
      return itemAt(
        list,
        key,
        () => `aared/docx: heading index ${String(key)} is outside this ${list.length}-heading outline`,
      );
    }
    const found = findByTitle(list, key, (heading) => heading.title);
    if (found !== undefined) return found;
    throw new AaredError("not-found", `aared/docx: no heading titled ${JSON.stringify(key)} in this ${list.length}-heading outline`);
  }
}

/**
 * An opened Word document: the Worker session plus every section resolved
 * eagerly at `open()` time. Sections are consumed together by
 * rendering/layout, so they stay cached here — never read lazily per
 * section — while physical pages belong to the render result only.
 */
export class DocxDocument {
  readonly info: DocxInfo;
  readonly sections: SectionCollection;
  readonly headings: HeadingCollection;
  readonly media: MediaCollectionShape;
  /** @internal The outline `headings`, `search()` and `render()` share. */
  readonly outline: HeadingOutline;
  private searchIndex: DocxSearchEntry[] | undefined = undefined;
  private closed = false;

  constructor(
    private readonly session: WorkerSession,
    sections: readonly DocumentSection[],
  ) {
    this.info = { ...session.info, format: "docx" };
    this.sections = new SectionCollection(sections);
    this.outline = headingOutline(sections);
    this.headings = new HeadingCollection(this.outline);
    this.media = new SessionMedia(session);
  }

  /**
   * Case-insensitive substring search over cached heading, paragraph,
   * list-item, table-cell, and footnote/endnote text. The sections resolved
   * at `open()` time, so no Worker round-trip is needed — it is async only so
   * every format's `search()` has the same shape. The searchable blocks are
   * indexed once on first use. One match per matching block in document
   * order, notes last; `heading` names the nearest preceding body heading in
   * the same section (null when none precedes the match — including when the
   * match is itself the first heading). Headers, footers, and text boxes are
   * not searched. Rejects `closed` after `close()`.
   */
  async search(query: string, options?: ReadOptions): Promise<DocxSearchMatch[]> {
    if (this.closed) throw new AaredError("closed", "aared/docx: document is closed");
    throwIfAborted(options?.signal);
    const needle = foldQuery(query);
    if (needle === null) return [];
    this.searchIndex ??= buildDocxSearchIndex(this.sections.list(), this.outline);
    const matches: DocxSearchMatch[] = [];
    for (const entry of this.searchIndex) {
      if (entry.lower.includes(needle)) matches.push(entryMatch(entry));
    }
    return matches;
  }

  /**
   * End the Worker session, reject pending reads, and release the search
   * index. The Worker itself is kept for the next `open()` and stops about
   * 15 s after the last document closes.
   */
  close(): void {
    this.closed = true;
    this.searchIndex = undefined;
    this.session.close();
  }
}

/**
 * Open a `.docx` file. Resolves every section eagerly with one Worker
 * `unit(0)` call — font resolution happens inside `unit()` — and caches the
 * result on the document; the Worker keeps no copy. Fails with an
 * `unsupported-format` `AaredError` when the file is not actually a Word
 * document. Headless-safe: no DOM.
 */
export async function open(source: Source, options?: OpenOptions): Promise<DocxDocument> {
  const session = await WorkerSession.openAs("docx", source, options);
  const signal = options?.signal;
  let sections: DocumentSection[];
  try {
    sections = await session.unit<DocumentSection[]>(0, signal === undefined ? undefined : { signal });
  } catch (error) {
    // The body is read here, not at open: a document that fails to parse
    // (no w:body, an abort) must not leave its Worker session behind.
    session.close();
    throw error;
  }
  return new DocxDocument(session, sections);
}

/**
 * Paint a paragraph or heading node's box, then what only Word adds: the line
 * follows the runs on it, the family the paragraph is mostly set in, shading and
 * borders. The authored line-spacing multiple scales Word's single-line box
 * (about 1.15em), not the font size.
 */
function applyParagraphLayout(line: HTMLElement, node: DocxParagraphLayout & { direction: TextDirection; runs: readonly TextRunModel[] }): void {
  applyParagraphBox(line, node, { singleLine: WORD_SINGLE_LINE });
  if (node.justifyShrink === true) line.classList.add(JUSTIFY_SHRINK);
  // The line box follows the runs on it, not the paragraph default (see `lineFontSizePt`).
  const lineFontPt = lineFontSizePt(node.runs, node.fontSizePt);
  if (lineFontPt !== undefined) line.style.fontSize = `${lineFontPt}pt`;
  const family = dominantFontFamily(node.runs);
  if (family !== undefined) line.style.fontFamily = cssFontFamily(family);
  if (node.shading !== undefined) {
    line.style.backgroundColor = node.shading;
    line.dataset.aaredShade = node.shading;
  }
  applyParagraphBorders(line, node);
}

/**
 * The family most of a paragraph's text is set in. A paragraph's own font is
 * the strut of every line in it, and CSS aligns that strut with the runs'
 * baselines: left as the page's UI font it stretches each line box by the
 * mismatch (a pixel a line for Arial runs), which Word never does. Matching it
 * to the runs keeps the line box at exactly the runs' height.
 */
function dominantFontFamily(runs: readonly TextRunModel[]): string | undefined {
  const weight = new Map<string, number>();
  for (const run of runs) {
    if (run.fontFamily !== undefined) weight.set(run.fontFamily, (weight.get(run.fontFamily) ?? 0) + run.text.length);
  }
  let best: string | undefined;
  let bestWeight = -1;
  for (const [family, length] of weight) {
    if (length > bestWeight) {
      best = family;
      bestWeight = length;
    }
  }
  return best;
}

/** A paragraph or table edge as a CSS border; `none` is an authored `nil` that clears an inherited edge. */
function borderCss(border: DocxTableEdge): string {
  return border === "none" ? "none" : `${border.widthPt}pt ${border.style} ${border.color}`;
}

/**
 * Paint `w:pBdr`. Top and bottom borders wrap the text with their gap as
 * padding. A side border sits outside the text indent — the text stays on its
 * indent and the border plus its gap hang out beside it — so the margin gives
 * up exactly what they take. A side border with nothing beneath it also runs
 * down through the paragraph's space-after, as Word draws it (a pull quote's
 * bar is taller than its text line).
 */
function applyParagraphBorders(line: HTMLElement, layout: DocxParagraphLayout): void {
  const borders = layout.borders;
  if (borders === undefined) return;
  const { top, bottom } = borders;
  if (top !== undefined) {
    line.style.borderTop = borderCss(top);
    if (top.spacePt > 0) line.style.paddingTop = `${top.spacePt}pt`;
  }
  if (bottom !== undefined) {
    line.style.borderBottom = borderCss(bottom);
    if (bottom.spacePt > 0) line.style.paddingBottom = `${bottom.spacePt}pt`;
  }
  const spaceAfterPt = layout.spaceAfterPt ?? 0;
  for (const [side, border, indent] of [
    ["Left", borders.left, layout.leftIndent],
    ["Right", borders.right, layout.rightIndent],
  ] as const) {
    if (border === undefined) continue;
    line.style[`border${side}`] = borderCss(border);
    if (border.spacePt > 0) line.style[`padding${side}`] = `${border.spacePt}pt`;
    line.style[`margin${side}`] = `${(indent ?? 0) * PX_PER_INCH - ptToPx(border.widthPt + border.spacePt)}px`;
    if (bottom === undefined && spaceAfterPt > 0) {
      const extension = document.createElement("span");
      extension.dataset.aaredElement = "border-extension";
      extension.setAttribute("aria-hidden", "true");
      extension.style.position = "absolute";
      extension.style.top = "100%";
      extension.style.height = `${spaceAfterPt}pt`;
      // Absolute offsets measure from the padding box, and the border sits just outside it.
      extension.style[side === "Left" ? "left" : "right"] = `-${border.widthPt}pt`;
      extension.style[`border${side}`] = borderCss(border);
      line.style.position = "relative";
      line.append(extension);
    }
  }
}

/** Paint a run's character formatting onto its span; shared by text runs and list markers. */
function applyRunStyle(span: HTMLElement, run: TextRunModel, exactPitch: boolean): void {
  applyRunBase(span, run, exactPitch);
  if (run.shading !== undefined && run.highlight === undefined) span.style.backgroundColor = run.shading;
  // Only text with no colour of its own follows its fill (see `paintAutomaticInk`).
  const fill = run.highlight ?? run.shading;
  if (fill !== undefined && run.color === undefined) span.dataset.aaredShade = fill;
  if (run.verticalAlign === "subscript" || run.verticalAlign === "superscript") {
    // `superscript`/`subscript` are OOXML `w:vertAlign` names that no browser
    // recognizes; the shift must be stated as CSS (a length, or `super`/`sub`).
    applyScript(span, run.verticalAlign, run.fontSizePt);
  }
}

/**
 * Word sets a super- or subscript at two thirds of its run size and raises a
 * superscript by a third of it (read from a native export's PDF: 6.96pt on a
 * 10.56pt line, baseline 4.7px up). A run without an authored size shrinks
 * relative to its paragraph and falls back to the CSS keyword shift.
 */
const SCRIPT_SCALE = 2 / 3;
const SUPERSCRIPT_RAISE = 1 / 3;

function applyScript(element: HTMLElement, kind: "superscript" | "subscript", basePt: number | undefined): void {
  element.style.fontSize = basePt === undefined ? `${roundPt(SCRIPT_SCALE * 100)}%` : `${roundPt(basePt * SCRIPT_SCALE)}pt`;
  element.style.verticalAlign = kind === "subscript"
    ? "sub"
    : basePt === undefined ? "super" : `${roundPt(basePt * SUPERSCRIPT_RAISE)}pt`;
}

function appendRunText(host: HTMLElement, run: TextRunModel, wiring: ImageWiring, direction: "ltr" | "rtl", exactPitch: boolean): void {
  if (run.image !== undefined) {
    appendInlineImage(host, run.image, run.hyperlink, wiring);
    return;
  }
  // Newlines arrive from `w:br`/`w:cr`; `pre-wrap` on the paragraph paints
  // them without splitting runs or losing bidi context. Live page fields
  // evaluate per layout page; flow rendering shows the field's cached result.
  const text = run.field === "PAGE"
    ? String(wiring.pageNumber ?? run.text)
    : run.field === "NUMPAGES"
      ? String(wiring.pageTotal ?? run.text)
      : run.text;
  const span = document.createElement("span");
  fillRunText(span, text, direction);
  applyRunStyle(span, run, exactPitch);
  if (run.hyperlink !== undefined) {
    const link = document.createElement("a");
    link.href = run.hyperlink;
    // Authored run color/underline decide how a link looks, not the browser default.
    link.style.color = "inherit";
    link.style.textDecoration = "inherit";
    link.append(span);
    if (run.footnoteRef !== undefined) {
      footnoteMarker(host, run, link, wiring);
    } else {
      host.append(link);
    }
    return;
  }
  if (run.footnoteRef !== undefined) {
    footnoteMarker(host, run, span, wiring);
    return;
  }
  host.append(span);
}

/** A list marker painted inline at the start of its item's first line. */
interface ListMarker {
  text: string;
  /** Width the marker owns before the text starts, in pixels (the hanging indent). */
  hangingPx: number;
  /** The numbering level's own marker formatting; absent follows the paragraph text. */
  run?: TextRunModel;
  /** What follows the marker: absent is Word's default tab to the text indent. */
  suffix?: "space" | "nothing";
}

function appendParagraph(host: HTMLElement, paragraph: DocxParagraphNode, wiring: ImageWiring, marker?: ListMarker): HTMLElement {
  const line = document.createElement("p");
  line.setAttribute("data-aared-element", "paragraph");
  line.setAttribute("class", BLOCK);
  applyParagraphLayout(line, paragraph);
  const exactPitch = exactLinePitch(paragraph);
  if (paragraph.floats !== undefined) appendFloats(line, paragraph.floats, paragraph, wiring);
  if (marker !== undefined) {
    // The marker is the first thing on the line in the paragraph's own
    // direction, so an RTL item carries it on the right and the line packs
    // by its alignment exactly as Word paints it.
    const mark = document.createElement("span");
    mark.setAttribute("data-aared-element", "list-marker");
    mark.setAttribute("data-aared-marker", marker.text);
    mark.textContent = marker.text;
    mark.setAttribute("aria-hidden", "true");
    // A tab to the text indent (no suffix): the marker owns the hanging
    // width, and sits at its start edge so the slack falls between marker and
    // text. Left to inherit the paragraph's physical `text-align` it would sit
    // against the text in an RTL item, whose alignment is swapped.
    mark.setAttribute("class", marker.suffix === undefined ? MARKER_TAB : MARKER);
    if (marker.run !== undefined) applyRunStyle(mark, marker.run, exactPitch);
    if (marker.suffix === undefined) mark.style.minWidth = `${marker.hangingPx}px`;
    line.append(mark);
    if (marker.suffix === "space") line.append(document.createTextNode(" "));
  }
  // Runs go straight into the paragraph: no wrapper, one element fewer per paragraph.
  for (const run of paragraph.runs) appendRunText(line, run, wiring, paragraph.direction, exactPitch);
  // A paragraph holding only pictures is still one line tall in Word.
  if (paragraph.runs.length === 0 && paragraph.floats !== undefined) line.append(document.createElement("br"));
  markTabs(line, paragraph.tabStops);
  host.append(line);
  return line;
}

function appendHeading(host: HTMLElement, heading: DocxHeadingNode, wiring: ImageWiring): void {
  const level = Math.min(Math.max(heading.level, 1), 6);
  const element = document.createElement(`h${level}`);
  element.setAttribute("data-aared-element", "heading");
  element.setAttribute("data-aared-level", String(heading.level));
  // Only outline headings are addressable (`goToHeading`, fragment ids); a
  // heading-styled paragraph in a header, footer, table or note paints the
  // same but takes no index, so ids stay unique and match `headings.list()`.
  const index = wiring.outline.get(heading);
  if (index !== undefined) {
    element.id = `aared-heading-${index + 1}`;
    element.setAttribute("data-aared-heading-index", String(index));
  }
  element.setAttribute("class", BLOCK);
  applyParagraphLayout(element, heading);
  const exactPitch = exactLinePitch(heading);
  if (heading.floats !== undefined) appendFloats(element, heading.floats, heading, wiring);
  for (const run of heading.runs) appendRunText(element, run, wiring, heading.direction, exactPitch);
  markTabs(element, heading.tabStops);
  host.append(element);
}

/**
 * How much thicker than authored a table edge is painted. A browser draws a
 * border of a device pixel or more, so a 0.25pt (0.33px) hairline paints 1px;
 * Word lets it be that thin. The cell gives the surplus back out of its padding
 * so rows keep Word's height.
 */
function borderSnapSlackPx(edge: DocxTableEdge | undefined): number {
  if (edge === undefined || edge === "none") return 0;
  return paintedBorderPx(edge) - ptToPx(edge.widthPt);
}

/** The width a table edge paints at: its authored width, at least one pixel. */
function paintedBorderPx(edge: DocxTableEdge | undefined): number {
  if (edge === undefined || edge === "none") return 0;
  const exact = ptToPx(edge.widthPt);
  return exact > 0 ? Math.max(1, Math.floor(exact)) : 0;
}

/**
 * A cell's side padding in pixels: Word's margin (inches) plus the half of
 * `edge` that falls inside the cell, less the border this cell paints there.
 */
function sideInsetPx(marginIn: number, edge: DocxTableEdge | undefined, painted: DocxTableEdge | undefined): number {
  const half = edge === undefined || edge === "none" ? 0 : ptToPx(edge.widthPt) / 2;
  return Math.max(0, marginIn * PX_PER_INCH + half - paintedBorderPx(painted));
}

function appendTable(host: HTMLElement, table: DocxTableNode, wiring: ImageWiring): void {
  const grid = createTableGrid(table.direction, table.columns);
  grid.style.minWidth = "0";
  grid.style.boxSizing = "border-box";
  const totalColumnWidth = table.columns.reduce((sum, width) => sum + width, 0);
  // An authored `w:tblW` wins; else the columns' own sum; else the container.
  const authoredWidth = table.width ?? (totalColumnWidth > 0 ? totalColumnWidth : undefined);
  grid.style.width = authoredWidth === undefined ? "100%" : inchesToCss(authoredWidth);
  grid.style.maxWidth = "100%";
  // The grid carries the table's direction, so inline-start is the side
  // Word measures `w:tblInd` from.
  if (table.indent !== undefined) grid.style.marginInlineStart = inchesToCss(table.indent);
  const columnCount = table.columns.length > 0
    ? table.columns.length
    : Math.max(1, ...table.rows.map((row) => row.cells.reduce((sum, cell) => sum + (cell.columnSpan ?? 1), 0)));
  const busy: number[] = new Array<number>(columnCount).fill(0);
  const lastRow = table.rows.length - 1;
  const borders = table.borders;
  const margins = table.cellMargins;
  for (const [rowIndex, row] of table.rows.entries()) {
    let column = 0;
    for (const cellModel of row.cells) {
      while ((busy[column] ?? 0) > 0) column += 1;
      const columnSpan = cellModel.columnSpan ?? 1;
      const rowSpan = cellModel.rowSpan ?? 1;
      const firstColumn = column === 0;
      const lastColumn = column + columnSpan >= columnCount;
      const firstRow = rowIndex === 0;
      const touchesBottom = rowIndex + rowSpan - 1 >= lastRow;
      const spannedIn = table.columns.slice(column, column + columnSpan).reduce((sum, width) => sum + width, 0);
      for (let offset = 0; offset < columnSpan; offset++) busy[column + offset] = rowSpan;
      column += columnSpan;

      const cell = createTableCell(cellModel);
      // Each shared edge is drawn once: outer edges on the boundary cells,
      // inside edges as the later cell's top/start line. A cell's own edge wins.
      const top = cellModel.borders?.top ?? (firstRow ? borders?.top : borders?.insideH);
      const bottom = cellModel.borders?.bottom ?? (touchesBottom ? borders?.bottom : undefined);
      const start = cellModel.borders?.left ?? (firstColumn ? borders?.left : borders?.insideV);
      const end = cellModel.borders?.right ?? (lastColumn ? borders?.right : undefined);
      if (top !== undefined) cell.style.borderTop = borderCss(top);
      if (bottom !== undefined) cell.style.borderBottom = borderCss(bottom);
      if (start !== undefined) cell.style.borderInlineStart = borderCss(start);
      if (end !== undefined) cell.style.borderInlineEnd = borderCss(end);
      const pad = { ...margins, ...cellModel.margins };
      cell.style.paddingTop = `${Math.max(0, (pad.top ?? 0) * PX_PER_INCH - borderSnapSlackPx(top))}px`;
      cell.style.paddingBottom = `${Math.max(0, (pad.bottom ?? 0) * PX_PER_INCH - borderSnapSlackPx(bottom))}px`;
      // Borders above follow the table direction; cell margins do not: Word
      // keeps a `w:left` margin on the left even in a right-to-left table.
      // Word centres a side border on its grid line and keeps the text a
      // margin beyond the half inside the cell (probe docx-cell-border: a
      // 0.5pt, 3pt and 6pt border move the text 0.3, 1.4 and 2.45pt). The
      // browser paints the whole border inside the cell that draws it, so the
      // padding is the margin plus half the edge, less what is painted here;
      // an inside divider drawn by the next cell still keeps this cell's text
      // half its width away.
      const shared = lastColumn ? undefined : borders?.insideV;
      const [left, right] = table.direction === "rtl"
        ? [{ edge: end ?? shared, painted: end }, { edge: start, painted: start }]
        : [{ edge: start, painted: start }, { edge: end ?? shared, painted: end }];
      cell.style.paddingLeft = `${sideInsetPx(pad.left ?? 0, left.edge, left.painted)}px`;
      cell.style.paddingRight = `${sideInsetPx(pad.right ?? 0, right.edge, right.painted)}px`;
      if (cellModel.shading !== undefined) {
        cell.style.backgroundColor = cellModel.shading;
        cell.dataset.aaredShade = cellModel.shading;
      }
      // A picture anchored in the cell is placed against the cell's text width, not the page's column.
      const outer = wiring.frame;
      if (spannedIn > 0) {
        const insetPx = Number.parseFloat(cell.style.paddingLeft) + Number.parseFloat(cell.style.paddingRight);
        wiring.frame = { columnPx: Math.max(0, spannedIn * PX_PER_INCH - insetPx), columnLeftPx: 0 };
      }
      for (const node of cellModel.nodes) appendFlowNode(cell, node, wiring);
      wiring.frame = outer;
      grid.append(cell);
    }
    for (let index = 0; index < busy.length; index++) if ((busy[index] ?? 0) > 0) busy[index] = (busy[index] ?? 0) - 1;
  }
  host.append(grid);
}

function appendListItem(list: HTMLElement, item: DocxListItemNode, wiring: ImageWiring): void {
  const entry = document.createElement("li");
  entry.dataset.aaredElement = "list-item";
  entry.dataset.aaredLevel = String(item.level);
  entry.dataset.aaredOrdered = String(item.ordered);
  // Word's numbering geometry: the text sits at the left indent and the marker
  // hangs `hanging` before it. Both measure from the paragraph's start edge, so
  // the same rule places the marker on the right of an RTL item.
  const indent = item.paragraph.leftIndent;
  const hanging = item.paragraph.firstLineIndent !== undefined && item.paragraph.firstLineIndent < 0
    ? -item.paragraph.firstLineIndent
    : DEFAULT_HANGING_IN;
  const indentPx = indent === undefined ? (item.level + 1) * DEFAULT_HANGING_IN * PX_PER_INCH : indent * PX_PER_INCH;
  // Copy before stripping: the model keeps its numbering indents, only the
  // painted paragraph re-expresses them in logical (start-edge) terms.
  const unindented: DocxParagraphNode = { ...item.paragraph };
  delete unindented.leftIndent;
  delete unindented.firstLineIndent;
  const line = appendParagraph(entry, unindented, wiring, {
    text: item.marker,
    hangingPx: hanging * PX_PER_INCH,
    ...(item.markerRun === undefined ? {} : { run: item.markerRun }),
    ...(item.markerSuffix === undefined ? {} : { suffix: item.markerSuffix }),
  });
  line.style.marginInlineStart = `${indentPx}px`;
  line.style.textIndent = `${-hanging * PX_PER_INCH}px`;
  list.append(entry);
}

function appendList(host: HTMLElement, nodes: DocxListItemNode[], wiring: ImageWiring): void {
  const ordered = nodes[0]?.ordered === true;
  const list = document.createElement(ordered ? "ol" : "ul");
  list.dataset.aaredElement = "list";
  list.dataset.aaredOrdered = String(ordered);
  list.style.listStyle = "none";
  list.style.margin = `${BLOCK_MARGIN_EM}em 0`;
  list.style.padding = "0";
  if (ordered) {
    const start = Number.parseInt(nodes[0]?.marker ?? "", 10);
    if (Number.isInteger(start)) list.setAttribute("start", String(start));
  }
  for (const item of nodes) appendListItem(list, item, wiring);
  host.append(list);
}

interface ImageWiring {
  /** This render's picture URLs, one per media part; absent when the embedder gave no media resolver. */
  urls: MediaUrls | undefined;
  destroyed: boolean;
  /** Outline index of each body heading node (see `headingOutline`); other headings are absent. */
  outline: ReadonlyMap<DocxHeadingNode, number>;
  /** First in-text marker id per footnote, so backlinks have a target. */
  refAnchors: Map<string, string>;
  /** Layout page fields; absent in flow mode (runs keep their static text). */
  pageNumber?: number;
  pageTotal?: number;
  /** The column and page an anchored picture is placed in; set by whoever paints a container, before painting it. */
  frame?: AnchorFrame | undefined;
  /** Run on destroy: observers and listeners a render started (the split wrap's resize observer). */
  cleanups: Array<() => void>;
}

/** A US Letter text column (8.5in less two one-inch margins), for a picture painted before any container set the frame. */
const DEFAULT_FRAME: AnchorFrame = { columnPx: 6.5 * PX_PER_INCH, columnLeftPx: 0 };

// Anchored pictures' constant styles (what varies per picture stays inline).
/**
 * A picture's holder. The plan is in physical pixels from the left: in a
 * right-to-left paragraph a fixed-width block would sit at the right edge and
 * its left margin be recomputed, so the holder always lays out left to right.
 */
const ANCHOR_HOLDER = cssClass("display:block;box-sizing:content-box;direction:ltr");
/** A zero-size anchor at the paragraph's top-left an overlay is measured from; the auto right margin keeps it left in a right-to-left paragraph. */
const ANCHOR_POINT = cssClass("position:relative;width:0;height:0;margin-left:0;margin-right:auto");
/**
 * A float text wraps around. The room kept around the picture is padding inside
 * the border box and the exclusion is that border box, so a vertical offset (a
 * margin, outside it) leaves the lines above the picture free.
 */
const ANCHOR_FLOAT = cssClass("shape-outside:border-box");
/** Top and bottom: a float the full width of the column, so no line runs beside it. */
const ANCHOR_BAND = cssClass("float:left;width:100%;box-sizing:border-box");
/** The picture itself. Stated so a host reset such as Tailwind's `img { display: block; max-width: 100% }` does not matter. */
const ANCHOR_PICTURE = cssClass("display:block;max-width:none");
const ANCHOR_OVERLAY = cssClass("position:absolute");

/** The data attributes every anchored picture carries, on whichever element is its outermost. */
function markAnchored(element: HTMLElement, image: DocxImageNode, placement: string): void {
  element.dataset.aaredElement = "anchored-image";
  element.dataset.aaredMediaId = image.mediaId;
  element.dataset.aaredWrap = image.anchor!.wrap;
  element.dataset.aaredPlacement = placement;
}

/**
 * Pictures anchored in a paragraph (`wp:anchor`), painted as its first
 * children so they take part in its line layout. `planFloat` decides the kind:
 * a CSS float text wraps around on one side (square, tight, through), a
 * full-width band (top and bottom, and the wraps a float cannot do), or an
 * absolutely positioned overlay (behind or in front of the text), or, for
 * text down both sides of the picture, a split paragraph that `split-wrap.ts`
 * lays out once it is on the page. Everything here is real DOM text.
 */
function appendFloats(
  line: HTMLElement,
  floats: readonly DocxImageNode[],
  paragraph: { leftIndent?: number; rightIndent?: number; runs: readonly TextRunModel[] },
  wiring: ImageWiring,
): void {
  const frame = wiring.frame ?? DEFAULT_FRAME;
  const indentPx = (paragraph.leftIndent ?? 0) * PX_PER_INCH;
  const placed: HTMLElement[] = [];
  let splitHere = false;
  for (const image of floats) {
    const anchor = image.anchor;
    if (anchor === undefined) continue;
    let plan = planFloat(image, frame, indentPx);
    const width = image.width * PX_PER_INCH;
    const height = image.height * PX_PER_INCH;
    if (plan.kind === "split") {
      // One split per paragraph, and only text a split can move between slots
      // (no pictures, links, notes, tabs or line breaks in it): otherwise a band.
      if (!splitHere && splittable(paragraph.runs)) {
        splitHere = true;
        const picture = anchoredPicture(image, width, height, wiring);
        markAnchored(picture, image, "split");
        picture.style.position = "absolute";
        registerSplit(line, {
          image,
          frame,
          leftIndentPx: indentPx,
          rightIndentPx: (paragraph.rightIndent ?? 0) * PX_PER_INCH,
          picture,
        });
        placed.push(picture);
        continue;
      }
      plan = { kind: "band", leftPx: plan.leftPx, topPx: plan.topPx };
    }
    const dist = {
      top: anchor.distance.top * PX_PER_INCH,
      bottom: anchor.distance.bottom * PX_PER_INCH,
      left: anchor.distance.left * PX_PER_INCH,
      right: anchor.distance.right * PX_PER_INCH,
    };
    const picture = anchoredPicture(image, width, height, wiring);
    if (plan.kind === "overlay") {
      picture.classList.add(ANCHOR_OVERLAY);
      picture.style.left = `${plan.leftPx}px`;
      picture.style.top = `${plan.topPx}px`;
      picture.style.zIndex = plan.behind ? "-1" : "1";
      picture.dataset.aaredLayer = plan.behind ? "behind" : "front";
      if (plan.to === "paragraph") {
        // Measured from the paragraph: a zero-size anchor at its top-left, taking no room in the text.
        const point = document.createElement("span");
        point.setAttribute("class", `${ANCHOR_HOLDER} ${ANCHOR_POINT}`);
        markAnchored(point, image, "overlay");
        point.append(picture);
        placed.push(point);
      } else {
        // Measured from the page (or header): no positioned wrapper, so the
        // nearest positioned ancestor is the page itself.
        markAnchored(picture, image, "overlay");
        placed.push(picture);
      }
      continue;
    }
    const holder = document.createElement("span");
    markAnchored(holder, image, plan.kind);
    holder.setAttribute("class", `${ANCHOR_HOLDER} ${ANCHOR_FLOAT}${plan.kind === "band" ? ` ${ANCHOR_BAND}` : ""}`);
    const padTop = Math.min(dist.top, plan.topPx);
    holder.style.paddingTop = `${padTop}px`;
    holder.style.paddingBottom = `${dist.bottom}px`;
    holder.style.marginTop = `${plan.topPx - padTop}px`;
    if (plan.kind === "band") {
      picture.style.marginLeft = `${plan.leftPx - indentPx}px`;
    } else {
      const near = plan.side === "left" ? dist.left : dist.right;
      const facing = plan.side === "left" ? dist.right : dist.left;
      const padNear = Math.min(near, plan.edgePx);
      holder.style.float = plan.side;
      holder.style.width = `${width}px`;
      holder.style.height = `${height}px`;
      holder.style.paddingLeft = `${plan.side === "left" ? padNear : facing}px`;
      holder.style.paddingRight = `${plan.side === "left" ? facing : padNear}px`;
      const outside = plan.edgePx - padNear - (plan.side === "left" ? indentPx : 0);
      if (plan.side === "left") holder.style.marginLeft = `${outside}px`;
      else holder.style.marginRight = `${outside}px`;
      const polygon = anchor.polygon;
      if (polygon !== undefined && (anchor.wrap === "tight" || anchor.wrap === "through")) {
        // Word's polygon is a share of the picture; the border box adds the padding around it.
        const left = Number.parseFloat(holder.style.paddingLeft);
        const points = polygon.map(([x, y]) => `${Math.round((left + x * width) * 100) / 100}px ${Math.round((padTop + y * height) * 100) / 100}px`);
        holder.style.setProperty("shape-outside", `polygon(${points.join(", ")}) border-box`);
        holder.style.setProperty("shape-margin", `${Math.max(dist.left, dist.right)}px`);
      }
    }
    holder.append(picture);
    placed.push(holder);
  }
  line.prepend(...placed);
}

/** The picture of an anchored image: its box while loading, then the media. */
function anchoredPicture(image: DocxImageNode, width: number, height: number, wiring: ImageWiring): HTMLElement {
  if (wiring.urls === undefined) {
    const pending = document.createElement("span");
    pending.dataset.aaredImagePending = "true";
    pending.setAttribute("class", ANCHOR_PICTURE);
    pending.style.width = `${width}px`;
    pending.style.height = `${height}px`;
    return pending;
  }
  const element = document.createElement("img");
  element.alt = image.alt ?? "";
  element.setAttribute("class", ANCHOR_PICTURE);
  element.style.width = `${width}px`;
  element.style.height = `${height}px`;
  wireImageMedia(element, image, wiring.urls, wiring);
  return element;
}

function footnoteMarker(host: HTMLElement, run: TextRunModel, inner: HTMLElement, wiring: ImageWiring): void {
  const footnote = run.footnoteRef!;
  const marker = document.createElement("sup");
  marker.dataset.aaredElement = "footnote-ref";
  marker.dataset.aaredFootnote = footnote;
  // The wrapper keeps the semantics; the run inside already carries the
  // superscript size and shift, which the UA `sup` style would stack on top of.
  marker.style.verticalAlign = "baseline";
  marker.style.fontSize = "inherit";
  let anchor = wiring.refAnchors.get(footnote);
  if (anchor === undefined) {
    anchor = `aared-footnote-ref-${footnote}`;
    wiring.refAnchors.set(footnote, anchor);
    marker.id = anchor;
  }
  const target = document.createElement("a");
  target.href = `#aared-footnote-${footnote}`;
  target.append(inner);
  marker.append(target);
  host.append(marker);
}

/**
 * Point `element` at its media's shared URL (a header logo repeats on every
 * page, so its images share one fetch and one Blob), or fall back to alt text
 * when the fetch fails.
 */
function wireImageMedia(
  element: HTMLImageElement,
  image: { mediaId: string; contentType: string; alt?: string },
  urls: MediaUrls,
  wiring: ImageWiring,
): void {
  element.dataset.aaredMediaId = image.mediaId;
  void urls.url(image.mediaId, image.contentType).then(
    (url) => {
      if (url !== undefined && !wiring.destroyed) element.src = url;
    },
    () => {
      if (!wiring.destroyed) element.alt = image.alt ?? image.mediaId;
    },
  );
}

/** A picture on the text line: it sits on the baseline like a glyph, not in its own block. */
function appendInlineImage(host: HTMLElement, image: DocxInlineImage, hyperlink: string | undefined, wiring: ImageWiring): void {
  const holder = document.createElement("span");
  holder.dataset.aaredElement = "inline-image";
  holder.dataset.aaredMediaId = image.mediaId;
  if (wiring.urls === undefined) {
    holder.textContent = image.alt ?? "";
    host.append(holder);
    return;
  }
  const element = document.createElement("img");
  element.alt = image.alt ?? "";
  element.style.width = inchesToCss(image.width);
  element.style.height = inchesToCss(image.height);
  // State both explicitly: a host reset such as Tailwind's `img { display: block }`
  // would otherwise lift the picture off its line and stack it above the text.
  element.style.display = "inline";
  element.style.verticalAlign = "baseline";
  holder.append(element);
  wireImageMedia(element, image, wiring.urls, wiring);
  if (hyperlink === undefined) {
    host.append(holder);
    return;
  }
  const link = document.createElement("a");
  link.href = hyperlink;
  link.append(holder);
  host.append(link);
}

function appendImage(host: HTMLElement, image: DocxImageNode, wiring: ImageWiring): void {
  const figure = document.createElement("figure");
  figure.dataset.aaredElement = "image";
  figure.dataset.aaredMediaId = image.mediaId;
  figure.style.margin = `${BLOCK_MARGIN_EM}em 0`;
  figure.style.minWidth = "0";
  figure.style.maxWidth = "100%";
  if (wiring.urls !== undefined) {
    const element = document.createElement("img");
    element.alt = image.alt ?? "";
    element.style.display = "block";
    element.style.maxWidth = "100%";
    element.style.height = "auto";
    element.style.width = inchesToCss(image.width);
    // Auto margins center a fixed-width block, so they sit on the `<img>`
    // itself: the `<figure>` fills its container, collapsing them to zero.
    if (image.align === "center") {
      element.style.marginLeft = "auto";
      element.style.marginRight = "auto";
    } else if (image.align === "right") {
      element.style.marginLeft = "auto";
    }
    figure.append(element);
    wireImageMedia(element, image, wiring.urls, wiring);
  } else {
    figure.dataset.aaredImagePending = "true";
    const caption = document.createElement("figcaption");
    caption.textContent = image.alt ?? image.mediaId;
    figure.append(caption);
  }
  host.append(figure);
}

function appendUnsupported(host: HTMLElement, node: DocxUnsupportedNode): void {
  const fallback = document.createElement("div");
  fallback.dataset.aaredElement = "unsupported";
  fallback.dataset.aaredUnsupported = node.source;
  fallback.setAttribute("role", "note");
  fallback.style.border = "1px dashed currentColor";
  fallback.style.padding = "0.5em 0.75em";
  fallback.style.margin = `${BLOCK_MARGIN_EM}em 0`;
  fallback.textContent = node.label;
  host.append(fallback);
}

/** The first authored run size in a footnote body, which its label is set against. */
function firstRunSizePt(nodes: FlowNode[]): number | undefined {
  for (const node of nodes) {
    if (node.type !== "paragraph" && node.type !== "heading") continue;
    const sized = node.runs.find((run) => run.fontSizePt !== undefined);
    if (sized !== undefined) return sized.fontSizePt;
  }
  return undefined;
}

/** Width of a page's footnote separator, as a share of the text column (Word's is about two inches). */
const FOOTNOTE_RULE_SHARE = "30%";

/** The footnote section for `host`, created with its separator rule and list on first use. */
function footnoteList(host: HTMLElement): Element {
  const existing = host.querySelector('[data-aared-element="footnote-list"]');
  if (existing !== null) return existing;
  const section = document.createElement("section");
  section.dataset.aaredElement = "footnotes";
  const rule = document.createElement("hr");
  rule.style.border = "0";
  rule.style.borderTop = "1px solid currentColor";
  rule.style.width = FOOTNOTE_RULE_SHARE;
  rule.style.margin = "0 0 4px";
  const list = document.createElement("ol");
  list.dataset.aaredElement = "footnote-list";
  list.style.listStyle = "none";
  list.style.margin = "0";
  list.style.padding = "0";
  section.append(rule, list);
  host.append(section);
  return list;
}

/**
 * Paint one footnote. Its label is a superscript at the start of the first
 * paragraph, as in Word (not a list number). The back-link is a reading aid
 * for flow output; paged output hides it visually and keeps it for assistive
 * technology.
 */
function appendFootnote(
  host: HTMLElement,
  id: string,
  label: string,
  direction: "ltr" | "rtl",
  nodes: FlowNode[],
  wiring: ImageWiring,
  backlink: "visible" | "hidden" = "visible",
): void {
  const item = document.createElement("li");
  item.id = `aared-footnote-${id}`;
  item.dataset.aaredElement = "footnote";
  item.dataset.aaredFootnote = id;
  item.dataset.aaredLabel = label;
  item.style.direction = direction;
  for (const node of nodes) appendFlowNode(item, node, wiring);
  const mark = document.createElement("sup");
  mark.dataset.aaredElement = "footnote-label";
  mark.textContent = label;
  applyScript(mark, "superscript", firstRunSizePt(nodes));
  (item.querySelector('[data-aared-element="paragraph"]') ?? item).prepend(mark);
  const link = document.createElement("a");
  const anchor = wiring.refAnchors.get(id);
  link.href = anchor === undefined ? `#aared-footnote-${id}` : `#${anchor}`;
  link.textContent = ` ${RETURN_ARROW}`;
  link.setAttribute("aria-label", "Back to reference");
  if (backlink === "hidden") {
    link.style.cssText = "position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap";
  }
  item.append(link);
  footnoteList(host).append(item);
}

function appendPageBreak(host: HTMLElement, node: DocxPageBreakNode): void {
  // Flow mode paints an authored page break as a single blank line — the
  // same visual the old `w:br`-as-`\n` run text produced.
  const line = document.createElement("p");
  line.dataset.aaredElement = "page-break";
  line.dataset.aaredBreakStrength = node.strength;
  line.style.margin = "0";
  line.style.minWidth = "0";
  line.append(document.createElement("br"));
  host.append(line);
}

function appendFlowNode(host: HTMLElement, node: FlowNode, wiring: ImageWiring): void {
  switch (node.type) {
    case "paragraph":
      appendParagraph(host, node, wiring);
      break;
    case "heading":
      appendHeading(host, node, wiring);
      break;
    case "table":
      appendTable(host, node, wiring);
      break;
    case "image":
      appendImage(host, node, wiring);
      break;
    case "list":
      appendList(host, node.items, wiring);
      break;
    case "footnote":
      appendFootnote(host, node.id, node.label, node.direction, node.nodes, wiring);
      break;
    case "unsupported":
      appendUnsupported(host, node);
      break;
    case "pageBreak":
      appendPageBreak(host, node);
      break;
  }
}

/**
 * How a document is laid out: `flow` (the default) is one continuous column
 * that reflows to its container, like Word's Web Layout — there are no pages;
 * `layout` cuts it into physical pages, like Print Layout. Structural: a
 * change needs a new `render()`, so it is not part of `Appearance`.
 */
export interface DocxPagination {
  mode?: "flow" | "layout";
  /** How `w:lastRenderedPageBreak` hints place page breaks (layout mode only). Default `"prefer"`. */
  softBreaks?: SoftBreakPolicy;
}

export interface DocxRenderInput {
  /**
   * The document to paint, or already-resolved sections for callers holding
   * a detached model. Passing the `DocxDocument` defaults `mediaResolver`
   * to the owning session (the same rule as PPTX/XLSX handle inputs);
   * detached sections need an explicit `mediaResolver` for images.
   */
  document: DocxDocument | readonly DocumentSection[];
  target: HTMLElement;
  pagination?: DocxPagination;
  appearance?: Appearance;
  /** Starting zoom, applied without a `zoomchange`. Default `"fit-width"`. */
  zoom?: ZoomValue;
  mediaResolver?: MediaResolver;
}

export interface DocxRenderHandle extends EventTarget {
  /** Disconnect media loading and release image object URLs. Idempotent. */
  destroy(): void;
  /** `true` once destroyed; every method is then a no-op that fires nothing. */
  readonly destroyed: boolean;
  /**
   * Change zoom: in layout mode `fit-width` scales the widest page to the
   * container and `fit-page` also keeps the tallest page in view; in flow mode
   * both fit modes are 100% (the column already flows to its container). A
   * number scales in place. Fires `zoomchange` only when the value changes.
   */
  setZoom(zoom: ZoomValue): void;
  /** The requested zoom. */
  readonly zoom: ZoomValue;
  /** The scale actually painted — the number, or what the fit mode computed. */
  readonly scale: number;
  /**
   * Replace the viewer appearance in place without re-pagination: refreshes
   * the theme marker, keeping scroll position and zoom. Content stays
   * source-faithful; theme never affects layout metrics.
   */
  setAppearance(appearance: Appearance): void;
  /** Laid-out pages: `0` and `[]` in flow mode, which has no pages. */
  readonly pageCount: number;
  readonly pages: readonly DocxPageInfo[];
  /**
   * Non-fatal things this render noticed about its own output: layout mode's
   * pagination edge cases (an oversize block, a footnote that overran its
   * page), plus a font this render couldn't actually paint with in either
   * mode. See `RenderWarning`.
   */
  readonly warnings: readonly RenderWarning[];
  /**
   * Scroll a body heading into view: a number is its `document.headings`
   * index; a string is its title (exact, then ignoring case — never a
   * substring). Rejects `not-found` when nothing matches.
   */
  goToHeading(target: string | number): Promise<void>;
  /**
   * Scroll laid-out page `index` (zero-based) into view. Rejects
   * `unsupported-feature` in flow mode, which has no pages, and `not-found`
   * outside `0 <= index < pageCount`.
   */
  goToPage(index: number): Promise<void>;
}

/**
 * Resolve the automatic text over every authored fill under `root`. Text
 * with no colour of its own takes the ink that reads on the fill it sits on
 * (Word's Automatic colour does the same). The fill is recorded on the
 * element when it is painted (`data-aared-shade`).
 */
function paintAutomaticInk(root: HTMLElement): void {
  for (const element of root.querySelectorAll<HTMLElement>("[data-aared-shade]")) {
    element.style.color = automaticInk(element.dataset.aaredShade!);
  }
}

/** A page-like root's default surface and ink, and the automatic text over the fills beneath it. */
function paintSurface(root: HTMLElement): void {
  root.style.background = "white";
  root.style.color = "black";
  paintAutomaticInk(root);
}

/** What `render()` hands the flow and layout painters once the input is resolved. */
interface RenderOptions {
  mediaResolver: MediaResolver | undefined;
  pagination: DocxPagination | undefined;
  appearance: Appearance | undefined;
  zoom: ZoomValue;
  outline: ReadonlyMap<DocxHeadingNode, number>;
}

function isDark(appearance: Appearance | undefined): boolean {
  return resolveTheme(appearance?.theme) === "dark";
}

/**
 * Viewer chrome state on the painted root: the theme marker (like PPTX) plus
 * light content tokens. The document itself always paints source-faithful.
 */
function paintAppearanceRoot(root: HTMLElement, dark: boolean): void {
  root.dataset.aaredTheme = dark ? "dark" : "light";
  paintTokens(root, "light");
}

class DocxViewer extends RenderHandleBase<Appearance> implements DocxRenderHandle {
  readonly pageCount: number;
  readonly pages: readonly DocxPageInfo[];
  /** Where the painted headings and pages live, once `render()` has painted. */
  #root: HTMLElement | null = null;

  constructor(
    zoom: ZoomValue,
    pages: readonly DocxPageInfo[],
    private readonly outline: readonly HeadingInfo[],
    private readonly layout: boolean,
  ) {
    super(zoom);
    this.pageCount = pages.length;
    this.pages = pages;
  }

  setAppearance(appearance: Appearance): void {
    if (this.destroyed) return;
    this.applyAppearance(appearance);
  }

  async goToHeading(target: string | number): Promise<void> {
    if (this.destroyed) return;
    const heading = typeof target === "number"
      ? itemAt(this.outline, target, () => `aared/docx: heading index ${String(target)} is outside this ${this.outline.length}-heading outline`)
      : findByTitle(this.outline, target, (entry) => entry.title);
    if (heading === undefined) {
      throw new AaredError("not-found", `aared/docx: no heading titled ${JSON.stringify(target)} in this ${this.outline.length}-heading outline`);
    }
    this.#root?.querySelector(`[data-aared-heading-index="${heading.index}"]`)?.scrollIntoView();
  }

  async goToPage(index: number): Promise<void> {
    if (this.destroyed) return;
    if (!this.layout) {
      throw new AaredError(
        "unsupported-feature",
        'aared/docx: this document is rendered in flow mode, which has no pages; render with pagination: { mode: "layout" }',
      );
    }
    itemAt(this.pages, index, () => `aared/docx: page index ${String(index)} is outside this ${this.pageCount}-page layout`);
    this.#root?.querySelectorAll('[data-aared-element="page"]')[index]?.scrollIntoView();
  }

  /** @internal Wired by the painter beside `attach()`: the painted root navigation searches. */
  attachRoot(root: HTMLElement): void {
    this.#root = root;
  }
}

/**
 * Hyperlink runs paint as real `<a>` elements (`appendRunText`). Footnote
 * forward/back links (`#aared-footnote-*`) keep their default in-page jump;
 * every other anchor is view-only preview chrome: never navigate, surface
 * the target to the embedder as `linkactivate` instead.
 */
function wireHyperlinks(root: HTMLElement, viewer: DocxViewer): () => void {
  const onClick = (event: Event): void => {
    const origin = event.target as Element | null;
    const anchor = origin?.closest("a[href]");
    if (!(anchor instanceof HTMLAnchorElement) || !root.contains(anchor)) return;
    const href = anchor.getAttribute("href") ?? "";
    if (href.startsWith("#aared-footnote-")) return;
    if (anchor.closest('[data-aared-element="footnote"]') !== null) return;
    event.preventDefault();
    viewer.dispatchEvent(new CustomEvent("linkactivate", { detail: { target: href } }));
  };
  root.addEventListener("click", onClick);
  root.addEventListener("auxclick", onClick);
  return () => {
    root.removeEventListener("click", onClick);
    root.removeEventListener("auxclick", onClick);
  };
}

function createWiring(options: RenderOptions): ImageWiring {
  return {
    urls: options.mediaResolver === undefined ? undefined : new MediaUrls(options.mediaResolver),
    destroyed: false,
    outline: options.outline,
    refAnchors: new Map<string, string>(),
    cleanups: [],
  };
}

function destroyWiring(wiring: ImageWiring): () => void {
  let destroyed = false;
  return () => {
    if (destroyed) return;
    destroyed = true;
    wiring.destroyed = true;
    wiring.urls?.close();
    wiring.urls = undefined;
    for (const cleanup of wiring.cleanups.splice(0)) cleanup();
  };
}

/**
 * Paint a document. Accepts the `DocxDocument` (media resolves through its
 * session unless an explicit `mediaResolver` is given) or already-resolved
 * sections. Async for envelope parity with `aared/pptx` and `aared/xlsx` —
 * painting itself is synchronous over the in-memory sections, so `await`
 * resolves immediately.
 *
 * Flow mode (the default) concatenates every section's root into one
 * article with the last section's header/footer — the pre-pagination
 * output. Layout mode paginates first and paints one
 * `section[data-aared-element="page"]` per laid-out page. The caller owns
 * block virtualization: `DocumentSection.root` entries are independent
 * blocks and may be sliced by a future block window; this renderer paints
 * whatever it is handed.
 */
export async function render(input: DocxRenderInput): Promise<DocxRenderHandle> {
  const source = input.document;
  const owner = source instanceof DocxDocument ? source : undefined;
  const sections = owner !== undefined ? [...owner.sections.list()] : [...(source as readonly DocumentSection[])];
  const outline = owner?.outline ?? headingOutline(sections);
  const options: RenderOptions = {
    mediaResolver: input.mediaResolver
      ?? (owner !== undefined ? (mediaId: string) => owner.media.get(mediaId) : undefined),
    pagination: input.pagination,
    appearance: input.appearance,
    zoom: input.zoom ?? "fit-width",
    outline: outline.indexOf,
  };
  if (options.pagination?.mode === "layout") {
    return await renderLayout(sections, input.target, options, outline.list);
  }
  return await renderFlow(sections, input.target, options, outline.list);
}

async function renderFlow(
  sections: DocumentSection[],
  container: HTMLElement,
  options: RenderOptions,
  outline: readonly HeadingInfo[],
): Promise<DocxRenderHandle> {
  const wiring = createWiring(options);
  // Flow mode paints no matter what — an unloaded font just reflows once its
  // face arrives, the same as any web page — but the fidelity warnings below
  // read whichever font actually ends up painting, so they wait first too.
  await awaitDocumentFonts(sections);
  // The shared defaults must be in the container's root before anything is attached or measured.
  const classSheet = adoptClassSheet(container);
  const article = document.createElement("article");
  article.dataset.aaredElement = "document";
  article.style.width = "100%";
  article.style.minWidth = "0";
  article.style.padding = "2em clamp(1em, 5%, 4em)";
  article.style.boxSizing = "border-box";
  article.style.fontFamily = "system-ui, sans-serif";
  article.style.lineHeight = "1.5";
  // A picture behind the text stays above the article's own background.
  article.style.isolation = "isolate";
  paintAppearanceRoot(article, isDark(options.appearance));

  // Flow mode keeps the pre-pagination contract: one article, the last
  // section's header/footer (the final `sectPr` references, as before), every
  // section's root concatenated in document order.
  const last = sections[sections.length - 1];
  // Anchored pictures are placed against the section's text column; there is no page.
  if (last !== undefined) wiring.frame = { columnPx: sectionBodyWidthPx(last), columnLeftPx: 0 };
  if (last?.header !== undefined && last.header.length > 0) {
    const header = document.createElement("header");
    header.dataset.aaredElement = "doc-header";
    header.style.borderBottom = "1px solid currentColor";
    header.style.marginBottom = "1em";
    for (const node of last.header) appendFlowNode(header, node, wiring);
    article.append(header);
  }
  const body = document.createElement("div");
  body.dataset.aaredElement = "doc-body";
  body.style.minWidth = "0";
  for (const section of sections) {
    wiring.frame = { columnPx: sectionBodyWidthPx(section), columnLeftPx: 0 };
    for (const node of section.root) appendFlowNode(body, node, wiring);
  }
  article.append(body);
  if (last?.footer !== undefined && last.footer.length > 0) {
    const footer = document.createElement("footer");
    footer.dataset.aaredElement = "doc-footer";
    footer.style.borderTop = "1px solid currentColor";
    footer.style.marginTop = "1em";
    for (const node of last.footer) appendFlowNode(footer, node, wiring);
    article.append(footer);
  }
  // Source-faithful defaults: authored white/black. Explicit
  // run/paragraph/table styles always win — they paint as more-specific
  // inline styles on descendants.
  paintSurface(article);
  container.replaceChildren(article);
  settleTabStops(article, () => !wiring.destroyed);
  // Flow width follows the container, so split paragraphs re-run when it changes.
  wiring.cleanups.push(settleSplitWraps(article, () => !wiring.destroyed, { observeWidth: true }));

  // Flow has no pages: the column flows to its container like a web page.
  const viewer = new DocxViewer(options.zoom, [], outline, false);
  viewer.setWarnings(fontFidelityWarnings(sections));
  wireViewer(viewer, article, container, wiring, classSheet, (appearance) => {
    paintAppearanceRoot(article, isDark(appearance));
  });
  return viewer;
}

/**
 * Connect a painted root to its handle, the same way for flow and layout.
 * A numeric zoom always pins the CSS `zoom` property on the root — it scales
 * the box in place, correctly growing what a scroll container sees as
 * overflow (unlike `transform`, which paints a scaled box without changing
 * its contribution to layout). `fit-width`/`fit-page` fit the widest (and,
 * for fit-page, tallest) laid-out page into the room `container` offers
 * (`measureFitRoom` — measured with the pages set aside, so a container that
 * only grows to fit them offers no height and fit-page falls back to width).
 * Flow mode has no fixed-width page to fit (the article already flows to the
 * container's width like ordinary content), so both fit modes are 100% there.
 * A `ResizeObserver` keeps a fit mode current as the host resizes — the same
 * reason `aared/pptx` observes its own container. Hyperlink clicks never
 * navigate (view-only). Destroy tears down the observer, the links, the media
 * URLs and the handle together. Only the appearance update differs between
 * flow and layout, so it is passed in.
 */
function wireViewer(
  viewer: DocxViewer,
  root: HTMLElement,
  container: HTMLElement,
  wiring: ImageWiring,
  classSheet: ClassSheetScope,
  applyAppearance: (appearance: Appearance | undefined) => void,
): void {
  const natural = {
    width: Math.max(0, ...viewer.pages.map((page) => page.width)),
    height: Math.max(0, ...viewer.pages.map((page) => page.height)),
  };
  let lastScale = 1;
  const applyZoom = (): number => {
    if (wiring.destroyed) return lastScale;
    // The container may have been attached (or moved into a shadow root) since render().
    classSheet.follow();
    const zoom = viewer.zoom;
    if (typeof zoom === "number") {
      root.style.width = "100%";
      lastScale = safeScale(zoom);
      root.style.setProperty("zoom", String(lastScale));
      return lastScale;
    }
    if (viewer.pages.length === 0) {
      root.style.width = "100%";
      root.style.removeProperty("zoom");
      lastScale = 1;
      return lastScale;
    }
    // A percentage width resolves against the containing block regardless of
    // `zoom`, so it never actually shrinks or grows under it — pin the
    // natural, unscaled width in pixels instead, the same way the pages
    // inside already carry their own fixed pixel size, and let `zoom` scale
    // that concrete box like it already correctly scales them.
    root.style.width = `${natural.width}px`;
    const scale = fitScale(zoom, measureFitRoom(container, root, zoom === "fit-page"), natural);
    // Not laid out yet: keep the last scale; the resize observer fits it later.
    lastScale = safeScale(scale ?? lastScale);
    root.style.setProperty("zoom", String(lastScale));
    return lastScale;
  };
  const resizeObserver = new ResizeObserver(() => viewer.refit());
  resizeObserver.observe(container);
  const unwireLinks = wireHyperlinks(root, viewer);
  const baseDestroy = destroyWiring(wiring);
  viewer.attach({
    applyZoom,
    destroy: () => {
      resizeObserver.disconnect();
      unwireLinks();
      baseDestroy();
      classSheet.release();
      viewer.markDestroyed();
    },
    applyAppearance: (appearance) => {
      if (wiring.destroyed) return;
      classSheet.follow();
      applyAppearance(appearance);
    },
  });
  viewer.attachRoot(root);
  viewer.refit();
}

/**
 * Lay every top-level block out once, off screen, at its section's body width
 * and under the page's own text defaults, and read back what the browser
 * built: each block's margins and border-box height, each table row's grid
 * track. Pagination then breaks on real heights, where the estimator could
 * only guess from average glyph widths (DECISION_LOG §21). One forced layout
 * for the whole document. Block images stay with the estimator: their height
 * is the model's until the picture loads.
 */
function measureFlow(sections: DocumentSection[], container: HTMLElement): MeasuredFlow {
  const blocks = new Map<FlowNode, { topPx: number; contentPx: number; bottomPx: number }>();
  const rows = new Map<DocxTableRowNode, number>();
  // Pictures keep their authored box while loading; nothing ever loads here.
  const urls = new MediaUrls(() => new Promise<ArrayBuffer>(() => {}));
  const wiring: ImageWiring = { urls, destroyed: false, outline: new Map(), refAnchors: new Map(), pageNumber: 1, pageTotal: 1, cleanups: [] };
  const probe = document.createElement("section");
  probe.setAttribute("class", PAGE);
  probe.setAttribute("aria-hidden", "true");
  probe.style.position = "absolute";
  probe.style.left = "-100000px";
  probe.style.top = "0";
  probe.style.visibility = "hidden";
  probe.style.overflow = "visible";
  probe.style.height = "auto";
  probe.style.padding = "0";
  const placed: Array<[FlowNode, Element]> = [];
  for (const section of sections) {
    wiring.frame = { columnPx: sectionBodyWidthPx(section), columnLeftPx: 0 };
    const body = document.createElement("div");
    body.style.width = `${sectionBodyWidthPx(section)}px`;
    // Blocks sit apart so no margin collapses into its neighbour's reading.
    body.style.display = "flow-root";
    for (const node of section.root) {
      if (node.type === "pageBreak" || node.type === "footnote" || node.type === "image") continue;
      const holder = document.createElement("div");
      holder.style.display = "flow-root";
      appendFlowNode(holder, node, wiring);
      body.append(holder);
      const element = holder.firstElementChild;
      if (holder.childElementCount === 1 && element !== null) placed.push([node, element]);
    }
    probe.append(body);
  }
  container.append(probe);
  // Split paragraphs are laid out before their heights are read, so pages break on what will paint.
  // Fonts are already awaited here, and the probe is removed right after: no second pass.
  settleSplitWraps(probe, () => true, { refont: false });
  try {
    for (const [node, element] of placed) {
      const style = getComputedStyle(element);
      blocks.set(node, {
        topPx: Number.parseFloat(style.marginTop) || 0,
        contentPx: element.getBoundingClientRect().height,
        bottomPx: Number.parseFloat(style.marginBottom) || 0,
      });
      if (node.type === "table") {
        // Each logical row is one grid track; a spanning cell spans tracks.
        const tracks = style.gridTemplateRows.split(" ").map((track) => Number.parseFloat(track));
        if (tracks.length === node.rows.length && tracks.every((track) => Number.isFinite(track))) {
          node.rows.forEach((row, index) => rows.set(row, tracks[index]!));
        }
      }
    }
  } finally {
    probe.remove();
    urls.close();
  }
  return { block: (node) => blocks.get(node), row: (row) => rows.get(row) };
}

async function renderLayout(
  sections: DocumentSection[],
  container: HTMLElement,
  options: RenderOptions,
  outline: readonly HeadingInfo[],
): Promise<DocxRenderHandle> {
  const wiring = createWiring(options);
  const classSheet = adoptClassSheet(container);
  // Wait for the fonts this document uses to finish loading before the
  // off-screen probe measures anything — a page count decided with the wrong
  // font swaps under it once the real (or calibrated substitute) face
  // finishes loading (DECISION_LOG §45).
  await awaitDocumentFonts(sections);
  const layout = paginate(sections, {
    softBreaks: options.pagination?.softBreaks ?? "prefer",
    measured: measureFlow(sections, container),
  });
  wiring.pageTotal = layout.pages.length;

  const stack = document.createElement("div");
  stack.dataset.aaredElement = "pages";
  stack.style.display = "flex";
  stack.style.flexDirection = "column";
  stack.style.alignItems = "center";
  stack.style.gap = "1.5rem";
  stack.style.width = "100%";
  stack.style.minWidth = "0";
  paintAppearanceRoot(stack, isDark(options.appearance));

  // Every page stays in the DOM (find-in-page, selection, fragment links), but
  // where the browser supports it a page off screen skips style and layout
  // until it is scrolled near: most of a long document's memory is layout.
  const lazy = lazyPagePaint();
  const sheets = layout.pages.map((page) => renderLayoutPage(page, layout.pages.length, wiring, lazy));
  stack.append(...sheets);
  container.replaceChildren(stack);
  if (lazy) settlePageTabStops(sheets, () => !wiring.destroyed);
  else settleTabStops(stack, () => !wiring.destroyed);
  wiring.cleanups.push(settleSplitWraps(stack, () => !wiring.destroyed));

  // Public page indexes are zero-based like every other index (the internal
  // layout numbers pages from 1 for PAGE fields and the `data-page` marker).
  const pageInfos: DocxPageInfo[] = layout.pages.map((page, index) => ({
    index,
    width: page.widthPx,
    height: page.heightPx,
    orientation: page.orientation,
  }));
  const viewer = new DocxViewer(options.zoom, pageInfos, outline, true);
  viewer.setWarnings([...layout.warnings, ...fontFidelityWarnings(sections)]);
  wireViewer(viewer, stack, container, wiring, classSheet, (appearance) => {
    paintAppearanceRoot(stack, isDark(appearance));
  });
  return viewer;
}

function renderLayoutPage(
  page: LayoutPage,
  pageTotal: number,
  wiring: ImageWiring,
  lazy: boolean,
): HTMLElement {
  wiring.pageNumber = page.index;
  wiring.pageTotal = pageTotal;
  const sheet = document.createElement("section");
  sheet.setAttribute("data-aared-element", "page");
  sheet.setAttribute("data-page", String(page.index));
  sheet.setAttribute("data-aared-orientation", page.orientation);
  // Page defaults follow the article contract (`paintSurface`, once the page
  // is populated): authored white/black. Explicit authored fills on
  // descendants are untouched. Layout pages paint with the estimator's
  // Word-like metrics (`PAGE`), not the flow article's roomier rhythm, so
  // estimates stay self-consistent with output. A lazy page (`PAGE_LAZY`) is
  // `content-visibility: auto`: its fixed box already sizes it while skipped,
  // and the containment that implies is what the page already is (its own
  // containing block, clipped, no counters).
  sheet.setAttribute("class", lazy ? `${PAGE} ${PAGE_LAZY}` : PAGE);
  sheet.style.width = `${page.widthPx}px`;
  sheet.style.height = `${page.heightPx}px`;
  sheet.style.paddingTop = `${page.bodyInsetsPx.top}px`;
  sheet.style.paddingRight = `${page.marginsPx.right}px`;
  sheet.style.paddingBottom = `${page.bodyInsetsPx.bottom}px`;
  sheet.style.paddingLeft = `${page.marginsPx.left}px`;
  // Where anchored pictures are placed: the text column, and the page they can be measured from.
  const pageFrame: AnchorFrame = {
    columnPx: page.widthPx - page.marginsPx.left - page.marginsPx.right,
    columnLeftPx: page.marginsPx.left,
    page: { widthPx: page.widthPx, heightPx: page.heightPx, marginTopPx: page.marginsPx.top, originXPx: 0, originYPx: 0 },
  };

  // Word's box model: a header hangs from its distance at the top of the page
  // and grows downward; the body starts below it (`bodyInsetsPx`). Headers are
  // never clipped, and carry no rule of their own — a header's line is its
  // paragraph's authored border.
  if (page.header !== undefined && page.header.length > 0) {
    const header = document.createElement("header");
    header.dataset.aaredElement = "page-header";
    header.style.position = "absolute";
    header.style.top = `${page.marginsPx.header}px`;
    header.style.left = `${page.marginsPx.left}px`;
    header.style.right = `${page.marginsPx.right}px`;
    // The header is a positioned box: page coordinates are translated by where it sits.
    wiring.frame = { ...pageFrame, page: { ...pageFrame.page!, originXPx: page.marginsPx.left, originYPx: page.marginsPx.header } };
    for (const node of page.header) appendFlowNode(header, node, wiring);
    sheet.append(header);
  }
  const body = document.createElement("div");
  body.dataset.aaredElement = "page-body";
  body.style.minWidth = "0";
  wiring.frame = pageFrame;
  for (const node of page.body) appendFlowNode(body, node, wiring);
  // Pages that continue the flow open flush against the header band.
  if (dropsTopSpacing(page.startReason) && body.firstElementChild instanceof HTMLElement) {
    body.firstElementChild.style.marginTop = "0";
  }
  sheet.append(body);
  if (page.footnotes.length > 0) {
    // Word anchors a page's footnotes to the bottom of the body area, above the
    // footer; the paginator already reserved their height in the body box.
    const notes = document.createElement("div");
    notes.dataset.aaredElement = "page-footnotes";
    notes.style.position = "absolute";
    notes.style.left = `${page.marginsPx.left}px`;
    notes.style.right = `${page.marginsPx.right}px`;
    notes.style.bottom = `${page.bodyInsetsPx.bottom}px`;
    for (const note of page.footnotes) {
      appendFootnote(notes, note.id, note.label, note.direction, note.nodes, wiring, "hidden");
    }
    sheet.append(notes);
  }
  if (page.footer !== undefined && page.footer.length > 0) {
    const footer = document.createElement("footer");
    footer.dataset.aaredElement = "page-footer";
    footer.style.position = "absolute";
    footer.style.bottom = `${page.marginsPx.footer}px`;
    footer.style.left = `${page.marginsPx.left}px`;
    footer.style.right = `${page.marginsPx.right}px`;
    // A footer hangs from the page bottom, so pictures in it are placed against its own column only.
    wiring.frame = { columnPx: pageFrame.columnPx, columnLeftPx: pageFrame.columnLeftPx };
    for (const node of page.footer) appendFlowNode(footer, node, wiring);
    sheet.append(footer);
  }
  paintSurface(sheet);
  return sheet;
}
