/**
 * Heuristic DOCX paginator: `DocumentSection[]` → `LayoutDocument`.
 * Headless and synchronous — no DOM, no fonts, no measurement host — so it
 * runs inside the Worker alongside the parser. Block heights are estimated
 * from the resolved model (font sizes, text length, image extents, authored
 * row heights) using the same CSS rules the DOM renderer paints with, which
 * keeps estimates self-consistent with the laid-out output.
 *
 * This aims for structural fidelity (correct page count, orientations, hard
 * breaks, section transitions, footnote assignment), never pixel-identical
 * Word pagination.
 */

import type {
  DocxFootnoteNode,
  DocxParagraphBorders,
  DocumentSection,
  FlowNode,
  TextRunModel,
} from "../model.js";
import { runsText } from "../search.js";
import { MINIMUM_LINE_HEIGHT, arabicShare, lineFontSizePt } from "../typography.js";
import { PX_PER_INCH, ptToPx, twipsToPx } from "../units.js";

/** Renderer article base font size: block `em` margins resolve against it. */
export const BASE_FONT_PX = 16;
/** Vertical margin, in `em`, the renderer gives lists, figures and fallback blocks. */
export const BLOCK_MARGIN_EM = 0.5;
const BLOCK_MARGIN_PX = BASE_FONT_PX * BLOCK_MARGIN_EM;
/**
 * Word's "single" line box as a multiple of the font size (about 1.15em for
 * Arial). A paragraph without a multiplier uses it as is, and an authored
 * `w:line` multiple scales it — not the font size — so 1.25 on 10.5pt Arial is
 * a 20px line. The renderer paints with the same value (imported from here),
 * which is what keeps estimates self-consistent with the laid-out output.
 */
export const WORD_SINGLE_LINE = 1.15;
/**
 * Average glyph advance per script as a fraction of the font size, used to
 * guess line wraps without shaping. Calibrated against painted Arial and Arabic
 * body text: Latin mixes measured ~0.44–0.48em, Arabic ~0.41em (0.5 wrapped
 * narrow table cells that really fit).
 */
const AVERAGE_CHAR_EM = { latin: 0.46, arabic: 0.4 } as const;
/**
 * A final line holding less than this share of a line is taken to fit on the
 * line above: real text varies around the average, and rounding every spill up
 * would systematically over-count.
 */
const WRAP_TOLERANCE_LINES = 0.1;

function averageCharEm(text: string): number {
  return AVERAGE_CHAR_EM.latin + (AVERAGE_CHAR_EM.arabic - AVERAGE_CHAR_EM.latin) * arabicShare(text);
}
/** Footnote separator rule plus its spacing, in pixels. */
const FOOTNOTE_SEPARATOR_PX = 12;

export type PageStartReason =
  | "document"
  | "explicit-page-break"
  | "overflow"
  | "section-next-page"
  | "section-even-page"
  | "section-odd-page";

export type SoftBreakPolicy = "ignore" | "prefer" | "force";

export interface LayoutPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
}

export interface LayoutPage {
  /** 1-based physical page number. */
  index: number;
  sectionIndex: number;
  widthPx: number;
  heightPx: number;
  orientation: "portrait" | "landscape";
  marginsPx: LayoutPageMargins;
  /**
   * Where the body actually sits. Word grows a header down from its distance
   * and starts the body below it when it outgrows the top margin (same at the
   * bottom for the footer), so these are the authored margins or the header/
   * footer band, whichever reaches further. `marginsPx` stays as authored.
   */
  bodyInsetsPx: { top: number; bottom: number };
  startReason: PageStartReason;
  body: FlowNode[];
  footnotes: DocxFootnoteNode[];
  header?: FlowNode[];
  footer?: FlowNode[];
}

export interface LayoutWarning {
  code: "oversize-block" | "footnote-overflow";
  message: string;
}

export interface LayoutDocument {
  pages: LayoutPage[];
  warnings: LayoutWarning[];
}

export interface PaginateOptions {
  /** How `preferred` (`w:lastRenderedPageBreak`) hints split pages. Default `"prefer"`. */
  softBreaks?: SoftBreakPolicy;
}

interface EstimateContext {
  bodyWidthPx: number;
}

/** Word's default text size when neither a paragraph nor a run states one. */
const FALLBACK_FONT_PT = 11;

/**
 * The size the text actually wraps at: each run's own size weighted by how much
 * text it holds, else the paragraph default. A 9pt run in a 10.5pt paragraph
 * fits more per line than the paragraph default alone would predict.
 */
function wrapFontSizePt(runs: TextRunModel[], fallbackPt: number): number {
  let weighted = 0;
  let length = 0;
  for (const run of runs) {
    weighted += run.text.length * (run.fontSizePt ?? fallbackPt);
    length += run.text.length;
  }
  return length > 0 ? weighted / length : fallbackPt;
}

interface TextBlockShape {
  text: string;
  /** Size that sets the line box (the paragraph's own). */
  lineFontPt: number;
  /** Size the glyphs wrap at (see `wrapFontSizePt`). */
  wrapFontPt: number;
  lineSpacing?: number;
  lineSpacingPt?: number;
  widthPx: number;
}

function textBlockHeightPx(shape: TextBlockShape): number {
  const charsPerLine = Math.max(1, Math.floor(shape.widthPx / (ptToPx(shape.wrapFontPt) * averageCharEm(shape.text))));
  const lines = Math.max(1, Math.ceil(shape.text.length / charsPerLine - WRAP_TOLERANCE_LINES));
  if (shape.lineSpacingPt !== undefined) return lines * ptToPx(shape.lineSpacingPt);
  // An authored multiple scales Word's single-line box (see the renderer).
  const multiple = (shape.lineSpacing ?? 1) * WORD_SINGLE_LINE;
  return lines * ptToPx(shape.lineFontPt) * Math.max(multiple, MINIMUM_LINE_HEIGHT);
}

/**
 * Text column available to a paragraph: the body width minus its indents.
 * Side borders hang outside the indent (as in Word), so they take none of it.
 */
function paragraphTextWidthPx(
  node: { leftIndent?: number; rightIndent?: number; borders?: DocxParagraphBorders },
  ctx: EstimateContext,
): number {
  const width = ctx.bodyWidthPx - (node.leftIndent ?? 0) * PX_PER_INCH - (node.rightIndent ?? 0) * PX_PER_INCH;
  return Math.max(24, width);
}

function paragraphContentPx(
  node: {
    runs: TextRunModel[];
    fontSizePt?: number;
    lineSpacing?: number;
    lineSpacingPt?: number;
    leftIndent?: number;
    rightIndent?: number;
    borders?: DocxParagraphBorders;
  },
  ctx: EstimateContext,
): number {
  const fontSizePt = lineFontSizePt(node.runs, node.fontSizePt) ?? FALLBACK_FONT_PT;
  const textPx = textBlockHeightPx({
    text: runsText(node.runs),
    lineFontPt: fontSizePt,
    wrapFontPt: wrapFontSizePt(node.runs, fontSizePt),
    ...(node.lineSpacing === undefined ? {} : { lineSpacing: node.lineSpacing }),
    ...(node.lineSpacingPt === undefined ? {} : { lineSpacingPt: node.lineSpacingPt }),
    widthPx: paragraphTextWidthPx(node, ctx),
  });
  // An inline picture sits on the baseline, so its line is the picture plus
  // the descent below it when that outgrows the text line.
  let imagePx = 0;
  for (const run of node.runs) {
    if (run.image !== undefined) imagePx = Math.max(imagePx, run.image.height * PX_PER_INCH);
  }
  return imagePx > 0 ? Math.max(textPx, imagePx + ptToPx(fontSizePt) * 0.25) : textPx;
}

/** Vertical metrics of one block: collapsible margins plus content height. */
interface BlockMetrics {
  topPx: number;
  contentPx: number;
  bottomPx: number;
  standalone: number;
}

function metrics(topPx: number, contentPx: number, bottomPx: number): BlockMetrics {
  return { topPx, contentPx, bottomPx, standalone: topPx + contentPx + bottomPx };
}

/**
 * Estimated painted metrics of one flow block in pixels, given the body
 * column width. Margins mirror the renderer's CSS exactly (paragraph
 * `spaceBefore`/`spaceAfter`, `BLOCK_MARGIN_EM` list and figure margins; tables
 * paint flush), so stacked estimates fold with CSS margin-collapsing semantics.
 * Wrap counts are guesses from average glyph widths, not measurements
 * (`AVERAGE_CHAR_EM`), calibrated against one fixture. Exported for unit
 * tests and estimator calibration; the paginator uses it internally.
 */
export function estimateBlockHeightPx(block: FlowNode, bodyWidthPx: number): number {
  return measureBlock(block, { bodyWidthPx }).standalone;
}

/**
 * Metrics per block and column width. The paginator re-folds a page's whole
 * body after every placed block, so without this each block would be measured
 * once per block placed after it. Blocks are immutable while paginating, and
 * the weak keys let the cache go with the model.
 */
const metricsCache = new WeakMap<FlowNode, { widthPx: number; metrics: BlockMetrics }>();

function measureBlock(block: FlowNode, ctx: EstimateContext): BlockMetrics {
  const cached = metricsCache.get(block);
  if (cached !== undefined && cached.widthPx === ctx.bodyWidthPx) return cached.metrics;
  const measured = measureUncached(block, ctx);
  metricsCache.set(block, { widthPx: ctx.bodyWidthPx, metrics: measured });
  return measured;
}

function measureUncached(block: FlowNode, ctx: EstimateContext): BlockMetrics {
  switch (block.type) {
    case "paragraph":
    case "heading": {
      // Top/bottom borders paint as border plus padding above/below the text;
      // unlike margins they never collapse.
      let extra = 0;
      if (block.borders?.top !== undefined) extra += ptToPx(block.borders.top.widthPt + block.borders.top.spacePt);
      if (block.borders?.bottom !== undefined) extra += ptToPx(block.borders.bottom.widthPt + block.borders.bottom.spacePt);
      return metrics(
        ptToPx(block.spaceBeforePt ?? 0),
        paragraphContentPx(block, ctx) + extra,
        ptToPx(block.spaceAfterPt ?? 0),
      );
    }
    case "list": {
      // A marker is inline at the start of its item's first line, so items
      // are ordinary paragraphs that fold with the same collapsing the
      // renderer paints. Each wraps in the column left after its numbering
      // indent.
      let total = 0;
      let pendingBottom = 0;
      let first = true;
      for (const item of block.items) {
        // The marker hangs inside the numbering indent, which the paragraph's
        // own `leftIndent` already takes off the column.
        const part = measureBlock(item.paragraph, ctx);
        total += first ? part.topPx : Math.max(pendingBottom, part.topPx);
        total += part.contentPx;
        pendingBottom = part.bottomPx;
        first = false;
      }
      if (!first) total += pendingBottom;
      return metrics(BLOCK_MARGIN_PX, total, BLOCK_MARGIN_PX);
    }
    case "table": {
      let total = 0;
      // The hairline between rows at its real thickness (0.25pt is 0.33px): the
      // renderer draws it snapped to a device pixel and takes the difference
      // back out of the cell padding, so a row is padding plus this.
      const insideH = block.borders?.insideH;
      const rowEdgePx = typeof insideH === "object" ? ptToPx(insideH.widthPt) : 0;
      for (const row of block.rows) {
        let rowHeight = 0;
        let column = 0;
        for (const cell of row.cells) {
          // Cells wrap within their own column width, not the body width.
          const span = cell.columnSpan ?? 1;
          const cellWidthPx = block.columns.length === 0
            ? ctx.bodyWidthPx
            : block.columns.slice(column, column + span).reduce((sum, width) => sum + width, 0) * PX_PER_INCH;
          column += span;
          const pad = { ...block.cellMargins, ...cell.margins };
          const padX = ((pad.left ?? 0) + (pad.right ?? 0)) * PX_PER_INCH;
          const padY = ((pad.top ?? 0) + (pad.bottom ?? 0)) * PX_PER_INCH;
          const cellCtx: EstimateContext = { bodyWidthPx: Math.max(24, Math.min(ctx.bodyWidthPx, cellWidthPx) - padX) };
          // A row is its text plus the cell padding plus the hairline between rows.
          const cellHeight = foldBlocksHeightPx(cell.nodes, cellCtx) + padY + rowEdgePx;
          if (cellHeight > rowHeight) rowHeight = cellHeight;
        }
        if (row.height !== undefined) {
          const authored = row.height * PX_PER_INCH;
          if (authored > rowHeight) rowHeight = authored;
        }
        total += rowHeight;
      }
      // Tables paint flush: Word gives them no margin of their own.
      return metrics(0, total, 0);
    }
    case "image":
      return metrics(BLOCK_MARGIN_PX, block.height * PX_PER_INCH, BLOCK_MARGIN_PX);
    case "footnote":
      return metrics(0, FOOTNOTE_SEPARATOR_PX + foldBlocksHeightPx(block.nodes, ctx), 0);
    case "unsupported":
      return metrics(BLOCK_MARGIN_PX, 48, BLOCK_MARGIN_PX);
    case "pageBreak":
      return metrics(0, 0, 0);
  }
}

/** Stacked height of consecutive blocks with CSS margin-collapsing semantics. */
function foldBlocksHeightPx(blocks: FlowNode[], ctx: EstimateContext): number {
  let total = 0;
  let pendingBottom = 0;
  let empty = true;
  for (const block of blocks) {
    const part = measureBlock(block, ctx);
    total += empty ? part.topPx : Math.max(pendingBottom, part.topPx);
    total += part.contentPx;
    pendingBottom = part.bottomPx;
    empty = false;
  }
  if (!empty) total += pendingBottom;
  return total;
}

function collectFootnoteRefs(blocks: FlowNode[], into: Set<string>): void {
  const runs = (list: TextRunModel[]): void => {
    for (const run of list) {
      if (run.footnoteRef !== undefined) into.add(run.footnoteRef);
    }
  };
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph":
      case "heading":
        runs(block.runs);
        break;
      case "list":
        for (const item of block.items) runs(item.paragraph.runs);
        break;
      case "table":
        for (const row of block.rows) {
          for (const cell of row.cells) collectFootnoteRefs(cell.nodes, into);
        }
        break;
      default:
        break;
    }
  }
}

interface PageBox {
  widthPx: number;
  heightPx: number;
}

interface BodyInsets {
  top: number;
  bottom: number;
}

/** Header/footer bands push the body inward once they outgrow the margins (see `LayoutPage.bodyInsetsPx`). */
function bodyInsets(section: DocumentSection): BodyInsets {
  const { margins, widthTwips } = section.properties.page;
  const ctx: EstimateContext = { bodyWidthPx: twipsToPx(widthTwips - margins.left - margins.right) };
  const top = twipsToPx(margins.top);
  const bottom = twipsToPx(margins.bottom);
  const headerBand = section.header !== undefined && section.header.length > 0
    ? twipsToPx(margins.header) + foldBlocksHeightPx(section.header, ctx)
    : 0;
  const footerBand = section.footer !== undefined && section.footer.length > 0
    ? twipsToPx(margins.footer) + foldBlocksHeightPx(section.footer, ctx)
    : 0;
  return { top: Math.max(top, headerBand), bottom: Math.max(bottom, footerBand) };
}

function bodyBox(section: DocumentSection): PageBox {
  const { page } = section.properties;
  const insets = bodyInsets(section);
  return {
    widthPx: twipsToPx(page.widthTwips - page.margins.left - page.margins.right),
    heightPx: twipsToPx(page.heightTwips) - insets.top - insets.bottom,
  };
}

/**
 * Whether a page opens flush against its header band. Word keeps a paragraph's
 * space-before at the top of a document's or a section's first page, and drops
 * it on a page that merely continues the flow — after a page break or an
 * overflow (read from a native export: an appendix heading starting a section
 * sits its 24px lower than one after a break).
 */
export function dropsTopSpacing(reason: PageStartReason): boolean {
  return reason === "explicit-page-break" || reason === "overflow";
}

/** Space above a page's first block (see `dropsTopSpacing`). */
function pageTopGap(reason: PageStartReason, part: BlockMetrics): number {
  return dropsTopSpacing(reason) ? 0 : part.topPx;
}

interface OpenPage {
  page: LayoutPage;
  box: PageBox;
  /** Folded content height (margins collapsed, trailing bottom excluded). */
  usedPx: number;
  prevBottomPx: number;
}

function newPage(index: number, sectionIndex: number, section: DocumentSection, startReason: PageStartReason): OpenPage {
  const { page } = section.properties;
  const box = bodyBox(section);
  return {
    page: {
      index,
      sectionIndex,
      widthPx: twipsToPx(page.widthTwips),
      heightPx: twipsToPx(page.heightTwips),
      orientation: page.orientation,
      marginsPx: {
        top: twipsToPx(page.margins.top),
        right: twipsToPx(page.margins.right),
        bottom: twipsToPx(page.margins.bottom),
        left: twipsToPx(page.margins.left),
        header: twipsToPx(page.margins.header),
        footer: twipsToPx(page.margins.footer),
      },
      bodyInsetsPx: bodyInsets(section),
      startReason,
      body: [],
      footnotes: [],
      ...(section.header === undefined ? {} : { header: section.header }),
      ...(section.footer === undefined ? {} : { footer: section.footer }),
    },
    box,
    usedPx: 0,
    prevBottomPx: 0,
  };
}

function sectionStartReason(sectionIndex: number, breakType: DocumentSection["properties"]["breakType"]): PageStartReason {
  if (sectionIndex === 0) return "document";
  switch (breakType) {
    case "evenPage":
      return "section-even-page";
    case "oddPage":
      return "section-odd-page";
    default:
      return "section-next-page";
  }
}

/**
 * Fill physical pages from document sections. Blocks are atomic: a block
 * that fits on an empty page but not the current one moves whole; a block
 * taller than an empty page is placed anyway with an `oversize-block`
 * warning. Footnote entries are assigned to the pages that reference them.
 */
export function paginate(sections: DocumentSection[], options?: PaginateOptions): LayoutDocument {
  const softBreaks = options?.softBreaks ?? "prefer";
  const warnings: LayoutWarning[] = [];
  const pages: LayoutPage[] = [];
  if (sections.length === 0) return { pages, warnings };

  // Flow-level footnote entries keyed by id; layout assigns them per page.
  const footnoteById = new Map<string, DocxFootnoteNode>();
  for (const section of sections) {
    for (const block of section.root) {
      if (block.type === "footnote" && !footnoteById.has(block.id)) footnoteById.set(block.id, block);
    }
  }

  let sectionIndex = 0;
  let current: OpenPage | undefined;
  const estimateCtx: EstimateContext = { bodyWidthPx: 0 };

  /** Push the open page when it has content; empty speculative pages vanish. */
  const flushPage = (): void => {
    if (current === undefined) return;
    assignFootnotes();
    if (current.page.body.length > 0) pages.push(current.page);
    current = undefined;
  };

  const openPage = (section: DocumentSection, reason: PageStartReason): void => {
    flushPage();
    current = newPage(pages.length + 1, sectionIndex, section, reason);
    estimateCtx.bodyWidthPx = current.box.widthPx;
  };

  /** Re-fold the open page's used height after any mutation. */
  const recomputeUsed = (): void => {
    if (current === undefined) return;
    let total = 0;
    let pendingBottom = 0;
    let empty = true;
    for (const block of current.page.body) {
      const part = measureBlock(block, estimateCtx);
      total += empty ? pageTopGap(current.page.startReason, part) : Math.max(pendingBottom, part.topPx);
      total += part.contentPx;
      pendingBottom = part.bottomPx;
      empty = false;
    }
    current.usedPx = total;
    current.prevBottomPx = empty ? 0 : pendingBottom;
  };

  const footnoteReservePx = (): number => {
    if (current === undefined) return 0;
    const refs = new Set<string>();
    collectFootnoteRefs(current.page.body, refs);
    let reserve = 0;
    for (const id of refs) {
      const note = footnoteById.get(id);
      if (note === undefined) continue;
      if (reserve === 0) reserve += FOOTNOTE_SEPARATOR_PX;
      reserve += foldBlocksHeightPx(note.nodes, estimateCtx);
    }
    return reserve;
  };

  const assignFootnotes = (): void => {
    if (current === undefined) return;
    const refs = new Set<string>();
    collectFootnoteRefs(current.page.body, refs);
    for (const id of refs) {
      const note = footnoteById.get(id);
      if (note === undefined) continue;
      if (!current.page.footnotes.some((entry) => entry.id === id)) current.page.footnotes.push(note);
    }
  };

  /**
   * Place one block with the footnote fixed-point. A block that fits the body
   * but overflows once its footnote set is reserved closes the page: spilled
   * blocks always continue on a fresh page, never retried on the same one.
   */
  const commitBlock = (section: DocumentSection, block: FlowNode): void => {
    if (current === undefined) openPage(section, sectionStartReason(sectionIndex, section.properties.breakType));
    const part = measureBlock(block, estimateCtx);
    if (part.standalone > current!.box.heightPx) {
      warnings.push({
        code: "oversize-block",
        message: `aared/core: block taller than an empty page (${Math.round(part.standalone)}px > ${Math.round(current!.box.heightPx)}px); placed with overflow`,
      });
      if (current!.page.body.length > 0) openPage(section, "overflow");
      current!.page.body.push(block);
      recomputeUsed();
      flushPage();
      return;
    }
    const pending: FlowNode[] = [block];
    while (pending.length > 0) {
      if (current === undefined) openPage(section, "overflow");
      const next = pending[0]!;
      const nextPart = measureBlock(next, estimateCtx);
      const gap = current!.page.body.length === 0
        ? pageTopGap(current!.page.startReason, nextPart)
        : Math.max(current!.prevBottomPx, nextPart.topPx);
      if (current!.usedPx + gap + nextPart.contentPx > current!.box.heightPx && current!.page.body.length > 0) {
        flushPage();
        continue;
      }
      pending.shift();
      current!.page.body.push(next);
      recomputeUsed();
      if (current!.usedPx + footnoteReservePx() > current!.box.heightPx) {
        // The footnote set overflows: spill trailing blocks until the page
        // is stable, then close it — spilled blocks continue fresh.
        while (current!.page.body.length > 1 && current!.usedPx + footnoteReservePx() > current!.box.heightPx) {
          const spilled = current!.page.body.pop()!;
          pending.unshift(spilled);
          recomputeUsed();
        }
        if (current!.usedPx + footnoteReservePx() > current!.box.heightPx) {
          warnings.push({
            code: "footnote-overflow",
            message: `aared/core: footnote reservation overflows page ${current!.page.index}; kept with overflow`,
          });
        }
        flushPage();
      }
    }
  };

  const splitOnBreak = (section: DocumentSection, strength: "mandatory" | "preferred"): void => {
    if (current === undefined) return;
    if (strength === "mandatory" || softBreaks === "force") {
      // A leading break on an empty page would mint a blank page; skip it.
      if (current.page.body.length > 0) openPage(section, "explicit-page-break");
      return;
    }
    if (softBreaks === "prefer" && current.page.body.length > 0 && current.usedPx > current.box.heightPx / 2) {
      openPage(section, "explicit-page-break");
    }
    // "ignore" drops the hint.
  };

  for (const section of sections) {
    if (sectionIndex > 0 && section.properties.breakType !== "continuous") {
      // Fresh page for the new section; even/odd insert one blank filler
      // when the parity is wrong. Parity is computed after flushing so the
      // filler math sees final page numbers.
      flushPage();
      const reason = sectionStartReason(sectionIndex, section.properties.breakType);
      const nextNumber = pages.length + 1;
      const parityWrong =
        (section.properties.breakType === "evenPage" && nextNumber % 2 !== 0) ||
        (section.properties.breakType === "oddPage" && nextNumber % 2 === 0);
      if (parityWrong) {
        pages.push(newPage(nextNumber, sectionIndex, section, reason).page);
      }
      // Content opens lazily on the first block with the section reason.
    } else if (sectionIndex > 0) {
      // Continuous: keep filling the open page under the new geometry.
      if (current !== undefined) estimateCtx.bodyWidthPx = bodyBox(section).widthPx;
    }
    for (const block of section.root) {
      if (block.type === "pageBreak") {
        splitOnBreak(section, block.strength);
        continue;
      }
      // Flow-level footnote entries are assigned per page, never placed inline.
      if (block.type === "footnote") continue;
      commitBlock(section, block);
    }
    sectionIndex += 1;
  }
  flushPage();
  // Page numbers are final once every page exists.
  for (const [position, page] of pages.entries()) page.index = position + 1;
  return { pages, warnings };
}
