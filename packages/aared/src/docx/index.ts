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
 * Adaptive content (`contentMode` defaults to `"adaptive"`): dark surfaces
 * apply only when the resolved theme is dark (matching XLSX). Light theme
 * renders authored; `source` is available as an explicit fidelity escape
 * hatch.
 */

import type {
  DocxHeadingNode,
  DocxImageNode,
  DocxListItemNode,
  DocxPageBreakNode,
  DocxParagraphLayout,
  DocxParagraphNode,
  DocxTableNode,
  DocxUnsupportedNode,
  DocumentSection,
  FlowNode,
  TextRunModel,
} from "../core/model.js";
import type {
  Appearance,
  Collection,
  DocumentInfo,
  MediaCollection as MediaCollectionShape,
  MediaOptions,
  MediaResolver,
  OpenOptions,
  Source,
  ViewerTheme,
  ZoomValue,
} from "../core/index.js";
import { AaredError } from "../core/index.js";
import { WorkerSession } from "../core/session.js";
import { paginate } from "../core/docx/pagination.js";
import type { LayoutPage, SoftBreakPolicy } from "../core/docx/pagination.js";

export type { SoftBreakPolicy } from "../core/docx/pagination.js";

/** PowerPoint's 96-DPI baseline shared for inch geometry, matching `aared/pptx`. */
const PIXELS_PER_INCH = 96;
/**
 * Minimum browser line box. DrawingML/Word relative spacing may name a
 * baseline distance smaller than Arabic glyph ink; CSS would crop a single
 * line, so the model keeps the authored multiplier while the renderer uses
 * the smallest safe box (see HANDOFF's title-layout correction).
 */
const MINIMUM_LINE_HEIGHT = 1.1;

function inches(value: number): string {
  return `${value * PIXELS_PER_INCH}px`;
}

/** Document metadata for a Word file; `format` is narrowed to `"docx"`. */
export interface DocxInfo extends DocumentInfo {
  format: "docx";
}

/** One outline entry derived from the cached sections, in document order. */
export interface HeadingInfo {
  text: string;
  level: number;
  sectionIndex: number;
}

/** One document result from `document.search()`. */
export interface DocxSearchMatch {
  /** Zero-based section index holding the match. */
  section: number;
  /** Nearest preceding heading text in the same section, if any. */
  heading: string | null;
  /** Matching block text, sliced to 120 characters. */
  text: string;
}

/** What a readonly document preview can and cannot do. */
export interface DocxCapabilities {
  textSelection: boolean;
  media: boolean;
  document: {
    flow: boolean;
    layout: boolean;
    footnotes: boolean;
  };
  editing: false;
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
    const section = this.sections[index];
    if (!Number.isInteger(index) || section === undefined) {
      throw new AaredError(
        "not-found",
        `aared/docx: section index ${String(index)} is outside this ${this.sections.length}-section document`,
      );
    }
    return section;
  }
}

/** `document.headings`: the heading outline in document order. */
export class HeadingCollection implements Collection<HeadingInfo, number> {
  private cached: HeadingInfo[] | undefined = undefined;

  constructor(private readonly sections: readonly DocumentSection[]) {}

  list(): readonly HeadingInfo[] {
    if (this.cached === undefined) {
      const headings: HeadingInfo[] = [];
      for (const [sectionIndex, section] of this.sections.entries()) {
        for (const node of section.root) {
          if (node.type === "heading") {
            headings.push({ text: node.text, level: node.level, sectionIndex });
          }
        }
      }
      this.cached = headings;
    }
    return this.cached;
  }

  get(index: number): HeadingInfo {
    const outline = this.list();
    const heading = outline[index];
    if (!Number.isInteger(index) || heading === undefined) {
      throw new AaredError(
        "not-found",
        `aared/docx: heading index ${String(index)} is outside this ${outline.length}-heading outline`,
      );
    }
    return heading;
  }
}

class DocxMediaCollection implements MediaCollectionShape {
  constructor(private readonly session: WorkerSession) {}

  get(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.media(mediaId, options);
  }
}

function paragraphText(runs: readonly TextRunModel[]): string {
  return runs.map((run) => run.text).join("").trim();
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
  readonly capabilities: DocxCapabilities = {
    textSelection: true,
    media: true,
    document: { flow: true, layout: true, footnotes: true },
    editing: false,
  };

  constructor(
    private readonly session: WorkerSession,
    sections: readonly DocumentSection[],
  ) {
    const info = session.info;
    this.info = { ...info, format: "docx" };
    this.sections = new SectionCollection(sections);
    this.headings = new HeadingCollection(sections);
    this.media = new DocxMediaCollection(session);
  }

  /**
   * Case-insensitive substring search over cached paragraph, heading, and
   * list-item text. Synchronous: every section resolved at `open()` time.
   * One match per matching block in document order; `heading` names the
   * nearest preceding heading in the same section (null when none precedes
   * the match — including when the match is itself the first heading).
   */
  search(query: string): DocxSearchMatch[] {
    const needle = query.toLowerCase();
    if (needle === "") return [];
    const matches: DocxSearchMatch[] = [];
    for (const [sectionIndex, section] of this.sections.list().entries()) {
      let heading: string | null = null;
      const consider = (text: string): void => {
        if (text !== "" && text.toLowerCase().includes(needle)) {
          matches.push({ section: sectionIndex, heading, text: text.slice(0, 120) });
        }
      };
      for (const node of section.root) {
        if (node.type === "heading") {
          consider(node.text);
          heading = node.text;
        } else if (node.type === "paragraph") {
          consider(paragraphText(node.runs));
        } else if (node.type === "list") {
          for (const item of node.items) consider(paragraphText(item.paragraph.runs));
        }
      }
    }
    return matches;
  }

  /** Terminate the Worker session and reject pending reads. */
  close(): void {
    this.session.close();
  }
}

/**
 * Open a `.docx` file. Resolves every section eagerly with one Worker
 * `unit(0)` call — font resolution happens inside `unit()` — and caches the
 * result on the document. Fails with an `unsupported-format` `AaredError`
 * when the file is not actually a Word document. Headless-safe: no DOM.
 */
export async function open(source: Source, options?: OpenOptions): Promise<DocxDocument> {
  const session = await WorkerSession.open("aared/docx", source, options);
  if (session.info.format !== "docx") {
    const detected = session.info.format;
    session.close();
    throw new AaredError(
      "unsupported-format",
      `aared/docx: expected a Word document; detected "${detected}" — open it with aared/${detected} instead`,
    );
  }
  const signal = options?.signal;
  const sections = await session.unit<DocumentSection[]>(
    0,
    signal === undefined ? undefined : { signal },
  );
  return new DocxDocument(session, sections);
}

function applyParagraphLayout(line: HTMLElement, layout: DocxParagraphLayout, direction: "ltr" | "rtl"): void {
  line.style.direction = direction;
  line.dataset.aaredDirection = direction;
  const align = layout.align ?? (direction === "rtl" ? "right" : "left");
  line.style.textAlign = align;
  line.dataset.aaredAlign = align;
  if (layout.fontSizePt !== undefined) line.style.fontSize = `${layout.fontSizePt}pt`;
  if (layout.lineSpacing !== undefined) {
    line.style.lineHeight = String(Math.max(layout.lineSpacing, MINIMUM_LINE_HEIGHT));
    line.dataset.aaredLineSpacing = String(layout.lineSpacing);
  }
  if (layout.lineSpacingPt !== undefined) line.style.lineHeight = `${layout.lineSpacingPt}pt`;
  if (layout.spaceBeforePt !== undefined) line.style.marginTop = `${layout.spaceBeforePt}pt`;
  if (layout.spaceAfterPt !== undefined) line.style.marginBottom = `${layout.spaceAfterPt}pt`;
  if (layout.leftIndent !== undefined) line.style.marginLeft = inches(layout.leftIndent);
  if (layout.rightIndent !== undefined) line.style.marginRight = inches(layout.rightIndent);
  if (layout.firstLineIndent !== undefined) line.style.textIndent = inches(layout.firstLineIndent);
  if (layout.shading !== undefined) line.style.backgroundColor = layout.shading;
  if (layout.borders !== undefined) {
    const { top, right, bottom, left } = layout.borders;
    if (top !== undefined) {
      line.style.borderTop = `${top.widthPt}pt ${top.style} ${top.color}`;
      if (top.spacePt > 0) line.style.paddingTop = `${top.spacePt}pt`;
    }
    if (right !== undefined) {
      line.style.borderRight = `${right.widthPt}pt ${right.style} ${right.color}`;
      if (right.spacePt > 0) line.style.paddingRight = `${right.spacePt}pt`;
    }
    if (bottom !== undefined) {
      line.style.borderBottom = `${bottom.widthPt}pt ${bottom.style} ${bottom.color}`;
      if (bottom.spacePt > 0) line.style.paddingBottom = `${bottom.spacePt}pt`;
    }
    if (left !== undefined) {
      line.style.borderLeft = `${left.widthPt}pt ${left.style} ${left.color}`;
      if (left.spacePt > 0) line.style.paddingLeft = `${left.spacePt}pt`;
    }
  }
}

function appendRunText(host: HTMLElement, run: TextRunModel, wiring: ImageWiring): void {
  // Newlines arrive from `w:br`/`w:cr`; `pre-wrap` on the paragraph paints
  // them without splitting runs or losing bidi context. Live page fields
  // evaluate per layout page; flow rendering keeps the static fiction.
  const text = run.field === "PAGE"
    ? String(wiring.pageNumber ?? run.text)
    : run.field === "NUMPAGES"
      ? String(wiring.pageTotal ?? run.text)
      : run.text;
  const span = document.createElement("span");
  span.textContent = text;
  if (run.bold === true) span.style.fontWeight = "700";
  if (run.italic === true) span.style.fontStyle = "italic";
  if (run.underline === true) span.style.textDecoration = "underline";
  if (run.strikethrough === true) {
    span.style.textDecoration = span.style.textDecoration === "" ? "line-through" : `${span.style.textDecoration} line-through`;
  }
  if (run.fontSizePt !== undefined) span.style.fontSize = `${run.fontSizePt}pt`;
  if (run.color !== undefined) span.style.color = run.color;
  if (run.fontFamily !== undefined) span.style.fontFamily = run.fontFamily;
  if (run.highlight !== undefined) span.style.backgroundColor = run.highlight;
  if (run.shading !== undefined && run.highlight === undefined) span.style.backgroundColor = run.shading;
  if (run.verticalAlign === "subscript" || run.verticalAlign === "superscript") {
    span.style.verticalAlign = run.verticalAlign;
    span.style.fontSize = run.fontSizePt === undefined ? "smaller" : `${run.fontSizePt * 0.75}pt`;
  }
  if (run.hyperlink !== undefined) {
    const link = document.createElement("a");
    link.href = run.hyperlink;
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

function appendParagraph(host: HTMLElement, paragraph: DocxParagraphNode, wiring: ImageWiring): void {
  const line = document.createElement("p");
  line.dataset.aaredElement = "paragraph";
  line.style.margin = "0";
  line.style.minWidth = "0";
  line.style.whiteSpace = "pre-wrap";
  applyParagraphLayout(line, paragraph, paragraph.direction);
  const text = document.createElement("span");
  text.dataset.aaredElement = "paragraph-text";
  for (const run of paragraph.runs) appendRunText(text, run, wiring);
  line.append(text);
  host.append(line);
}

function appendHeading(host: HTMLElement, heading: DocxHeadingNode, index: number, wiring: ImageWiring): void {
  const level = Math.min(Math.max(heading.level, 1), 6);
  const element = document.createElement(`h${level}`);
  element.dataset.aaredElement = "heading";
  element.dataset.aaredLevel = String(heading.level);
  element.id = `aared-heading-${index + 1}`;
  element.dataset.aaredHeadingIndex = String(index);
  element.style.margin = "0";
  element.style.minWidth = "0";
  element.style.whiteSpace = "pre-wrap";
  applyParagraphLayout(element, heading, heading.direction);
  const text = document.createElement("span");
  text.dataset.aaredElement = "paragraph-text";
  for (const run of heading.runs) appendRunText(text, run, wiring);
  element.append(text);
  host.append(element);
}

function cellJustifyContent(verticalAlign: "top" | "middle" | "bottom"): string {
  switch (verticalAlign) {
    case "middle":
      return "center";
    case "bottom":
      return "flex-end";
    default:
      return "flex-start";
  }
}

function appendTable(host: HTMLElement, table: DocxTableNode, wiring: ImageWiring): void {
  const grid = document.createElement("div");
  grid.setAttribute("role", "table");
  grid.dataset.aaredElement = "table";
  grid.dataset.aaredTableDirection = table.direction;
  grid.style.display = "grid";
  grid.style.direction = table.direction;
  grid.style.width = "100%";
  grid.style.minWidth = "0";
  grid.style.border = "1px solid currentColor";
  grid.style.margin = "0.5em 0";
  const totalColumnWidth = table.columns.reduce((sum, width) => sum + width, 0);
  if (table.columns.length > 0) {
    grid.style.gridTemplateColumns = totalColumnWidth > 0
      ? table.columns.map((width) => `${(width / totalColumnWidth) * 100}%`).join(" ")
      : table.columns.map(() => "1fr").join(" ");
  }
  for (const row of table.rows) {
    for (const cellModel of row.cells) {
      const cell = document.createElement("div");
      cell.setAttribute("role", "cell");
      cell.dataset.aaredElement = "table-cell";
      cell.style.minWidth = "0";
      cell.style.padding = "0.2em 0.35em";
      cell.style.border = "1px solid color-mix(in srgb, currentColor 25%, transparent)";
      if (cellModel.shading !== undefined) cell.style.backgroundColor = cellModel.shading;
      if (cellModel.verticalAlign !== undefined) {
        cell.style.display = "flex";
        cell.style.flexDirection = "column";
        cell.style.justifyContent = cellJustifyContent(cellModel.verticalAlign);
      }
      if (cellModel.columnSpan !== undefined) cell.style.gridColumn = `span ${cellModel.columnSpan}`;
      if (cellModel.rowSpan !== undefined) cell.style.gridRow = `span ${cellModel.rowSpan}`;
      for (const node of cellModel.nodes) appendFlowNode(cell, node, wiring);
      grid.append(cell);
    }
  }
  host.append(grid);
}

function appendListItem(list: HTMLElement, item: DocxListItemNode, wiring: ImageWiring): void {
  const entry = document.createElement("li");
  entry.dataset.aaredElement = "list-item";
  entry.dataset.aaredLevel = String(item.level);
  entry.dataset.aaredOrdered = String(item.ordered);
  entry.style.display = "grid";
  entry.style.gridTemplateColumns = "auto minmax(0, 1fr)";
  entry.style.gap = "0.5em";
  // The numbering indent positions the text once: the item starts there and
  // the inner paragraph paints without its own start indents (the same box
  // the paginator estimates). The marker hangs before the text, Word-style.
  const indent = item.paragraph.leftIndent;
  entry.style.marginInlineStart = indent === undefined ? `${item.level * 1.5}em` : inches(indent);
  const marker = document.createElement("span");
  marker.dataset.aaredElement = "list-marker";
  marker.dataset.aaredMarker = item.marker;
  marker.textContent = item.marker;
  marker.setAttribute("aria-hidden", "true");
  marker.style.justifySelf = "end";
  const body = document.createElement("div");
  body.style.minWidth = "0";
  // Copy before stripping: the model keeps its numbering indents, only the
  // painted item drops them (the `<li>` offset already positions the text).
  const unindented: DocxParagraphNode = { ...item.paragraph };
  if (indent !== undefined || item.paragraph.firstLineIndent !== undefined) {
    delete unindented.leftIndent;
    delete unindented.firstLineIndent;
  }
  appendParagraph(body, unindented, wiring);
  entry.append(marker, body);
  list.append(entry);
}

function appendList(host: HTMLElement, nodes: DocxListItemNode[], wiring: ImageWiring): void {
  const ordered = nodes[0]?.ordered === true;
  const list = document.createElement(ordered ? "ol" : "ul");
  list.dataset.aaredElement = "list";
  list.dataset.aaredOrdered = String(ordered);
  list.style.listStyle = "none";
  list.style.margin = "0.5em 0";
  list.style.padding = "0";
  if (ordered) {
    const start = Number.parseInt(nodes[0]?.marker ?? "", 10);
    if (Number.isInteger(start)) list.setAttribute("start", String(start));
  }
  for (const item of nodes) appendListItem(list, item, wiring);
  host.append(list);
}

interface ImageWiring {
  media: MediaResolver | undefined;
  revoke: Set<string> | undefined;
  destroyed: boolean;
  /** First in-text marker id per footnote, so backlinks have a target. */
  refAnchors: Map<string, string>;
  /** Layout page fields; absent in flow mode (runs keep their static text). */
  pageNumber?: number;
  pageTotal?: number;
}

function footnoteMarker(host: HTMLElement, run: TextRunModel, inner: HTMLElement, wiring: ImageWiring): void {
  const footnote = run.footnoteRef!;
  const marker = document.createElement("sup");
  marker.dataset.aaredElement = "footnote-ref";
  marker.dataset.aaredFootnote = footnote;
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

function appendImage(host: HTMLElement, image: DocxImageNode, wiring: ImageWiring): void {
  const figure = document.createElement("figure");
  figure.dataset.aaredElement = "image";
  figure.dataset.aaredMediaId = image.mediaId;
  figure.style.margin = "0.5em 0";
  figure.style.minWidth = "0";
  figure.style.maxWidth = "100%";
  if (image.float === "left" || image.float === "right") {
    figure.style.float = image.float;
    figure.style.marginInlineEnd = image.float === "left" ? "1em" : "0";
    figure.style.marginInlineStart = image.float === "right" ? "1em" : "0";
  }
  if (wiring.media !== undefined) {
    const element = document.createElement("img");
    element.alt = image.alt ?? "";
    element.style.display = "block";
    element.style.maxWidth = "100%";
    element.style.height = "auto";
    element.style.width = inches(image.width);
    // Auto margins center a fixed-width block, so they sit on the `<img>`
    // itself: the `<figure>` fills its container, collapsing them to zero.
    if (image.float === undefined && image.align === "center") {
      element.style.marginLeft = "auto";
      element.style.marginRight = "auto";
    } else if (image.float === undefined && image.align === "right") {
      element.style.marginLeft = "auto";
    }
    element.dataset.aaredMediaId = image.mediaId;
    figure.append(element);
    void wiring.media(image.mediaId).then(
      (bytes) => {
        if (wiring.destroyed || wiring.revoke === undefined) return;
        const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: image.contentType }));
        wiring.revoke.add(url);
        if (wiring.destroyed) {
          URL.revokeObjectURL(url);
          wiring.revoke.delete(url);
          return;
        }
        element.src = url;
      },
      () => {
        if (!wiring.destroyed) element.alt = image.alt ?? image.mediaId;
      },
    );
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
  fallback.style.margin = "0.5em 0";
  fallback.textContent = node.label;
  host.append(fallback);
}

function appendFootnote(host: HTMLElement, id: string, label: string, direction: "ltr" | "rtl", nodes: FlowNode[], wiring: ImageWiring): void {
  let section = host.querySelector<HTMLElement>('[data-aared-element="footnotes"]');
  if (section === null) {
    section = document.createElement("section");
    section.dataset.aaredElement = "footnotes";
    const rule = document.createElement("hr");
    section.append(rule);
    const list = document.createElement("ol");
    list.dataset.aaredElement = "footnote-list";
    list.style.paddingInlineStart = "1.5em";
    section.append(list);
    host.append(section);
  }
  const list = section.querySelector('[data-aared-element="footnote-list"]')!;
  const item = document.createElement("li");
  item.id = `aared-footnote-${id}`;
  item.dataset.aaredElement = "footnote";
  item.dataset.aaredFootnote = id;
  item.dataset.aaredLabel = label;
  item.style.direction = direction;
  for (const node of nodes) appendFlowNode(item, node, wiring);
  const backlink = document.createElement("a");
  const anchor = wiring.refAnchors.get(id);
  backlink.href = anchor === undefined ? `#aared-footnote-${id}` : `#${anchor}`;
  backlink.textContent = " ↩";
  backlink.setAttribute("aria-label", "Back to reference");
  item.append(backlink);
  list.append(item);
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

function appendFlowNode(host: HTMLElement, node: FlowNode, wiring: ImageWiring, headingIndex?: { value: number }): void {
  switch (node.type) {
    case "paragraph":
      appendParagraph(host, node, wiring);
      break;
    case "heading": {
      const counter = headingIndex ?? { value: 0 };
      appendHeading(host, node, counter.value, wiring);
      counter.value++;
      break;
    }
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

/** Viewer-only appearance for DOCX. */
export interface DocxAppearance {
  theme?: ViewerTheme;
  /**
   * `source` (default) paints the file as authored; `adaptive` paints a
   * dark reading mode that keeps every explicit authored style and only
   * recolors default surfaces. Never inversion or filters.
   */
  contentMode?: "source" | "adaptive";
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
  mediaResolver?: MediaResolver;
  pagination?: { mode?: "flow" | "layout"; softBreaks?: SoftBreakPolicy };
  appearance?: DocxAppearance | Appearance;
}

export interface DocxRenderOptions {
  mediaResolver?: MediaResolver;
  pagination?: { mode?: "flow" | "layout"; softBreaks?: SoftBreakPolicy };
  appearance?: DocxAppearance | Appearance;
}

export interface DocxRenderHandle extends EventTarget {
  /** Disconnect media loading and release image object URLs. */
  destroy(): void;
  /** Change zoom; emits `zoomchange`. Numbers scale via the CSS `zoom` property. */
  setZoom(zoom: ZoomValue): void;
  /** The active zoom. */
  readonly zoom: ZoomValue;
  /** Physical pages: 1 in flow mode, laid-out pages in layout mode. */
  pageCount: number;
  pages: Array<{ index: number; width: number; height: number; orientation: "portrait" | "landscape" }>;
  /**
   * Scroll a heading into view: a number selects the nth heading in
   * document order (zero-based, matching `document.headings`); a string
   * selects the first heading with exactly that text, else the first
   * heading containing it (case-insensitive). Throws a `not-found`
   * `AaredError` when nothing matches.
   */
  goToHeading(target: string | number): void;
}

/** Dark reading-mode surfaces for `adaptive`; `source` keeps authored white/black. */
const ADAPTIVE_SURFACE = "#1a1a1a";
const ADAPTIVE_INK = "#e8e8e8";
const ADAPTIVE_BORDER = "#3a3a3a";
const ADAPTIVE_MUTED = "#9a9a9a";
const ADAPTIVE_ACCENT = "#8ab4f8";
const SOURCE_BORDER = "#d1d5db";
const SOURCE_MUTED = "#6b7280";
const SOURCE_ACCENT = "#2563eb";
const ADAPTIVE_PAGE_EDGE = "#101010";

function resolveAppearance(options?: DocxRenderOptions): { contentMode: "source" | "adaptive"; dark: boolean } {
  const appearance = options?.appearance;
  const theme = appearance?.theme ?? "system";
  let dark = theme === "dark";
  if (theme === "system" && typeof matchMedia !== "undefined") {
    dark = matchMedia("(prefers-color-scheme: dark)").matches;
  }
  return {
    contentMode: appearance?.contentMode ?? "adaptive",
    dark,
  };
}

/**
 * Viewer chrome state on the painted root: the theme marker plus the
 * surface/ink variables the adaptive pass reads. Setting variables alone
 * never changes source-mode visuals — every default there stays a literal
 * style, exactly as before. Adaptive surfaces only apply when the resolved
 * theme is dark — matching XLSX's dark-gating behavior.
 */
function paintAppearanceRoot(root: HTMLElement, dark: boolean, contentMode: "source" | "adaptive"): void {
  root.dataset.aaredTheme = dark ? "dark" : "light";
  if (contentMode === "adaptive" && dark) {
    root.style.setProperty("--aared-surface", ADAPTIVE_SURFACE);
    root.style.setProperty("--aared-ink", ADAPTIVE_INK);
    root.style.setProperty("--aared-border", ADAPTIVE_BORDER);
    root.style.setProperty("--aared-muted-text", ADAPTIVE_MUTED);
    root.style.setProperty("--aared-accent", ADAPTIVE_ACCENT);
    root.style.setProperty("--aared-focus-ring", ADAPTIVE_ACCENT);
  } else {
    root.style.setProperty("--aared-surface", "#ffffff");
    root.style.setProperty("--aared-ink", "#000000");
    root.style.setProperty("--aared-border", SOURCE_BORDER);
    root.style.setProperty("--aared-muted-text", SOURCE_MUTED);
    root.style.setProperty("--aared-accent", SOURCE_ACCENT);
    root.style.setProperty("--aared-focus-ring", SOURCE_ACCENT);
  }
}

class DocxViewer extends EventTarget implements DocxRenderHandle {
  #zoom: ZoomValue = "fit-width";
  #destroyed = false;
  #applyZoom: () => void = () => {};
  #destroyHook: () => void = () => {};
  #findHeadings: () => HTMLElement[] = () => [];
  readonly pageCount: number;
  readonly pages: Array<{ index: number; width: number; height: number; orientation: "portrait" | "landscape" }>;

  constructor(
    pageCount: number,
    pages: Array<{ index: number; width: number; height: number; orientation: "portrait" | "landscape" }>,
  ) {
    super();
    this.pageCount = pageCount;
    this.pages = pages;
  }

  get zoom(): ZoomValue {
    return this.#zoom;
  }

  setZoom(zoom: ZoomValue): void {
    this.#zoom = zoom;
    this.#applyZoom();
    this.dispatchEvent(new CustomEvent("zoomchange", { detail: { zoom } }));
  }

  goToHeading(target: string | number): void {
    const headings = this.#findHeadings();
    if (typeof target === "number") {
      const element = Number.isInteger(target) ? headings[target] : undefined;
      if (element === undefined) {
        throw new AaredError(
          "not-found",
          `aared/docx: heading index ${String(target)} is outside this ${headings.length}-heading document`,
        );
      }
      element.scrollIntoView();
      return;
    }
    const exact = headings.find((element) => (element.textContent ?? "") === target);
    if (exact !== undefined) {
      exact.scrollIntoView();
      return;
    }
    const needle = target.toLowerCase();
    const partial = headings.find((element) => (element.textContent ?? "").toLowerCase().includes(needle));
    if (partial !== undefined) {
      partial.scrollIntoView();
      return;
    }
    throw new AaredError("not-found", `aared/docx: no heading matching "${target}"`);
  }

  destroy(): void {
    this.#destroyHook();
  }

  /** Wired by `render()`; `setZoom()` before wiring applies on first paint. */
  attach(applyZoom: () => void, destroy: () => void, findHeadings: () => HTMLElement[]): void {
    this.#applyZoom = applyZoom;
    this.#destroyHook = destroy;
    this.#findHeadings = findHeadings;
  }

  markDestroyed(): void {
    this.#destroyed = true;
  }

  get destroyed(): boolean {
    return this.#destroyed;
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

function destroyWiring(wiring: ImageWiring): () => void {
  let destroyed = false;
  return () => {
    if (destroyed) return;
    destroyed = true;
    wiring.destroyed = true;
    if (wiring.revoke !== undefined) {
      for (const url of wiring.revoke) URL.revokeObjectURL(url);
      wiring.revoke.clear();
    }
    wiring.revoke = undefined;
    wiring.media = undefined;
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
  const sections = source instanceof DocxDocument ? [...source.sections.list()] : [...source];
  const owner = source instanceof DocxDocument ? source : undefined;
  const mediaResolver = input.mediaResolver
    ?? (owner !== undefined ? (mediaId: string) => owner.media.get(mediaId) : undefined);
  const options: DocxRenderOptions | undefined = mediaResolver === undefined
    && input.pagination === undefined
    && input.appearance === undefined
    ? undefined
    : {
      ...(mediaResolver === undefined ? {} : { mediaResolver }),
      ...(input.pagination === undefined ? {} : { pagination: input.pagination }),
      ...(input.appearance === undefined ? {} : { appearance: input.appearance }),
    };
  if (options?.pagination?.mode === "layout") {
    return renderLayout(sections, input.target, options);
  }
  return renderFlow(sections, input.target, options);
}

function renderFlow(
  sections: DocumentSection[],
  container: HTMLElement,
  options?: DocxRenderOptions,
): DocxRenderHandle {
  const wiring: ImageWiring = {
    media: options?.mediaResolver,
    revoke: new Set<string>(),
    destroyed: false,
    refAnchors: new Map<string, string>(),
  };
  const article = document.createElement("article");
  article.dataset.aaredElement = "document";
  article.style.width = "100%";
  article.style.minWidth = "0";
  // Default surfaces only: authored white/black in source mode, dark reading
  // surfaces in adaptive. Explicit run/paragraph/table styles always win —
  // they paint as more-specific inline styles on descendants.
  const { contentMode, dark } = resolveAppearance(options);
  const adaptive = contentMode === "adaptive" && dark;
  article.style.background = adaptive ? ADAPTIVE_SURFACE : "white";
  article.style.color = adaptive ? ADAPTIVE_INK : "black";
  article.style.padding = "2em clamp(1em, 5%, 4em)";
  article.style.boxSizing = "border-box";
  article.style.fontFamily = "system-ui, sans-serif";
  article.style.lineHeight = "1.5";
  paintAppearanceRoot(article, dark, contentMode);

  const headings = { value: 0 };
  // Flow mode keeps the pre-pagination contract: one article, the last
  // section's header/footer (the final `sectPr` references, as before), every
  // section's root concatenated in document order.
  const last = sections[sections.length - 1];
  if (last?.header !== undefined && last.header.length > 0) {
    const header = document.createElement("header");
    header.dataset.aaredElement = "doc-header";
    header.style.borderBottom = "1px solid currentColor";
    header.style.marginBottom = "1em";
    for (const node of last.header) appendFlowNode(header, node, wiring, headings);
    article.append(header);
  }
  const body = document.createElement("div");
  body.dataset.aaredElement = "doc-body";
  body.style.minWidth = "0";
  for (const section of sections) {
    for (const node of section.root) appendFlowNode(body, node, wiring, headings);
  }
  article.append(body);
  if (last?.footer !== undefined && last.footer.length > 0) {
    const footer = document.createElement("footer");
    footer.dataset.aaredElement = "doc-footer";
    footer.style.borderTop = "1px solid currentColor";
    footer.style.marginTop = "1em";
    for (const node of last.footer) appendFlowNode(footer, node, wiring, headings);
    article.append(footer);
  }
  container.replaceChildren(article);

  // Fit modes are the natural flow: only a numeric zoom pins the CSS `zoom`
  // property on the article. Hyperlink clicks never navigate (view-only).
  const viewer = new DocxViewer(1, []);
  const applyZoom = (): void => {
    if (wiring.destroyed) return;
    const zoom = viewer.zoom;
    if (typeof zoom === "number") article.style.setProperty("zoom", String(zoom));
    else article.style.removeProperty("zoom");
  };
  const unwireLinks = wireHyperlinks(article, viewer);
  const baseDestroy = destroyWiring(wiring);
  viewer.attach(
    applyZoom,
    () => {
      unwireLinks();
      baseDestroy();
      viewer.markDestroyed();
    },
    () => [...container.querySelectorAll<HTMLElement>("[data-aared-heading-index]")],
  );
  applyZoom();
  return viewer;
}

function renderLayout(
  sections: DocumentSection[],
  container: HTMLElement,
  options?: DocxRenderOptions,
): DocxRenderHandle {
  const wiring: ImageWiring = {
    media: options?.mediaResolver,
    revoke: new Set<string>(),
    destroyed: false,
    refAnchors: new Map<string, string>(),
  };
  const layout = paginate(sections, { softBreaks: options?.pagination?.softBreaks ?? "prefer" });
  wiring.pageTotal = layout.pages.length;

  const stack = document.createElement("div");
  stack.dataset.aaredElement = "pages";
  stack.style.display = "flex";
  stack.style.flexDirection = "column";
  stack.style.alignItems = "center";
  stack.style.gap = "1.5rem";
  stack.style.width = "100%";
  stack.style.minWidth = "0";
  const { contentMode, dark } = resolveAppearance(options);
  const adaptive = contentMode === "adaptive" && dark;
  if (adaptive) stack.style.background = ADAPTIVE_PAGE_EDGE;
  paintAppearanceRoot(stack, dark, contentMode);

  const headings = { value: 0 };
  for (const page of layout.pages) {
    stack.append(renderLayoutPage(page, layout.pages.length, wiring, headings, adaptive));
  }
  container.replaceChildren(stack);

  const pageInfos = layout.pages.map((page) => ({
    index: page.index,
    width: page.widthPx,
    height: page.heightPx,
    orientation: page.orientation,
  }));
  const viewer = new DocxViewer(layout.pages.length, pageInfos);
  const applyZoom = (): void => {
    if (wiring.destroyed) return;
    const zoom = viewer.zoom;
    if (typeof zoom === "number") stack.style.setProperty("zoom", String(zoom));
    else stack.style.removeProperty("zoom");
  };
  const unwireLinks = wireHyperlinks(stack, viewer);
  const baseDestroy = destroyWiring(wiring);
  viewer.attach(
    applyZoom,
    () => {
      unwireLinks();
      baseDestroy();
      viewer.markDestroyed();
    },
    () => [...container.querySelectorAll<HTMLElement>("[data-aared-heading-index]")],
  );
  applyZoom();
  return viewer;
}

function renderLayoutPage(
  page: LayoutPage,
  pageTotal: number,
  wiring: ImageWiring,
  headings: { value: number },
  adaptive: boolean,
): HTMLElement {
  wiring.pageNumber = page.index;
  wiring.pageTotal = pageTotal;
  const sheet = document.createElement("section");
  sheet.dataset.aaredElement = "page";
  sheet.dataset.page = String(page.index);
  sheet.dataset.aaredOrientation = page.orientation;
  sheet.style.width = `${page.widthPx}px`;
  sheet.style.height = `${page.heightPx}px`;
  sheet.style.flexShrink = "0";
  // Page defaults follow the article contract: authored white/black in
  // source mode, dark reading surfaces in adaptive. Explicit authored fills
  // on descendants are untouched.
  sheet.style.background = adaptive ? ADAPTIVE_SURFACE : "white";
  sheet.style.color = adaptive ? ADAPTIVE_INK : "black";
  sheet.style.boxSizing = "border-box";
  sheet.style.position = "relative";
  sheet.style.paddingTop = `${page.marginsPx.top}px`;
  sheet.style.paddingRight = `${page.marginsPx.right}px`;
  sheet.style.paddingBottom = `${page.marginsPx.bottom}px`;
  sheet.style.paddingLeft = `${page.marginsPx.left}px`;
  sheet.style.overflow = "hidden";
  sheet.style.boxShadow = "0 2px 8px rgba(0,0,0,0.1)";
  // Layout pages paint with the estimator's Word-like metrics (not the flow
  // article's roomier rhythm) so estimates stay self-consistent with output.
  sheet.style.fontFamily = "system-ui, sans-serif";
  sheet.style.fontSize = "16px";
  sheet.style.lineHeight = "1.15";

  // Word's box model: headers/footers live in the margin zones, and the body
  // owns the full page-minus-margins box the paginator filled. Absolutely
  // positioning them there (rather than stacking them into the body) is what
  // keeps estimates and paint consistent. A header taller than its zone
  // clips instead of overlapping body text — the safe side of Word's
  // overlap behavior.
  if (page.header !== undefined && page.header.length > 0) {
    const header = document.createElement("header");
    header.dataset.aaredElement = "page-header";
    header.style.position = "absolute";
    header.style.top = "0";
    header.style.left = `${page.marginsPx.left}px`;
    header.style.right = `${page.marginsPx.right}px`;
    header.style.height = `${page.marginsPx.top}px`;
    header.style.overflow = "hidden";
    header.style.borderBottom = "1px solid currentColor";
    for (const node of page.header) appendFlowNode(header, node, wiring, headings);
    sheet.append(header);
  }
  const body = document.createElement("div");
  body.dataset.aaredElement = "page-body";
  body.style.minWidth = "0";
  for (const node of page.body) appendFlowNode(body, node, wiring, headings);
  for (const note of page.footnotes) {
    appendFootnote(body, note.id, note.label, note.direction, note.nodes, wiring);
  }
  sheet.append(body);
  if (page.footer !== undefined && page.footer.length > 0) {
    const footer = document.createElement("footer");
    footer.dataset.aaredElement = "page-footer";
    footer.style.position = "absolute";
    footer.style.bottom = "0";
    footer.style.left = `${page.marginsPx.left}px`;
    footer.style.right = `${page.marginsPx.right}px`;
    footer.style.height = `${page.marginsPx.bottom}px`;
    footer.style.overflow = "hidden";
    footer.style.borderTop = "1px solid currentColor";
    for (const node of page.footer) appendFlowNode(footer, node, wiring, headings);
    sheet.append(footer);
  }
  return sheet;
}
