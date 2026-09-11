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

const TWIPS_PER_INCH = 1440;
const PX_PER_INCH = 96;
const PT_PER_INCH = 72;
/** Renderer article base font size: block `em` margins resolve against it. */
const BASE_FONT_PX = 16;
/**
 * Word "single" line box used when a paragraph sets no multiplier. Layout
 * pages paint with this same base (tighter than the flow article's 1.5), so
 * estimates stay self-consistent with the laid-out output.
 */
const DEFAULT_LINE_MULTIPLIER = 1.15;
/** Minimum browser line box; mirrors the renderer's clamp. */
const MINIMUM_LINE_MULTIPLIER = 1.1;
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

function twipsToPx(twips: number): number {
  return (twips / TWIPS_PER_INCH) * PX_PER_INCH;
}

function ptToPx(points: number): number {
  return (points / PT_PER_INCH) * PX_PER_INCH;
}

interface EstimateContext {
  bodyWidthPx: number;
}

function defaultFontSizePt(runs: TextRunModel[]): number {
  let size = 0;
  for (const run of runs) {
    if (run.fontSizePt !== undefined && run.fontSizePt > size) size = run.fontSizePt;
  }
  return size > 0 ? size : 11;
}

function textBlockHeightPx(
  text: string,
  fontSizePt: number,
  lineSpacing: number | undefined,
  lineSpacingPt: number | undefined,
  spaceBeforePt: number | undefined,
  spaceAfterPt: number | undefined,
  bodyWidthPx: number,
): number {
  if (lineSpacingPt !== undefined) {
    const lines = Math.max(1, Math.ceil(text.length / Math.max(1, Math.floor(bodyWidthPx / ptToPx(fontSizePt)))));
    return ptToPx(spaceBeforePt ?? 0) + lines * ptToPx(lineSpacingPt) + ptToPx(spaceAfterPt ?? 0);
  }
  const fontPx = ptToPx(fontSizePt);
  const charsPerLine = Math.max(1, Math.floor(bodyWidthPx / (fontPx * 0.5)));
  const lines = Math.max(1, Math.ceil(text.length / charsPerLine));
  const multiplier = Math.max(lineSpacing ?? DEFAULT_LINE_MULTIPLIER, MINIMUM_LINE_MULTIPLIER);
  return ptToPx(spaceBeforePt ?? 0) + lines * fontPx * multiplier + ptToPx(spaceAfterPt ?? 0);
}

/**
 * Text column available to a paragraph: the body width minus owned indents
 * and left/right border-plus-space, mirroring the painted box.
 */
function paragraphTextWidthPx(
  node: { leftIndent?: number; rightIndent?: number; borders?: DocxParagraphBorders },
  ctx: EstimateContext,
): number {
  let width = ctx.bodyWidthPx - (node.leftIndent ?? 0) * PX_PER_INCH - (node.rightIndent ?? 0) * PX_PER_INCH;
  if (node.borders?.left !== undefined) width -= ptToPx(node.borders.left.widthPt + node.borders.left.spacePt);
  if (node.borders?.right !== undefined) width -= ptToPx(node.borders.right.widthPt + node.borders.right.spacePt);
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
  const text = node.runs.map((run) => run.text).join("");
  return textBlockHeightPx(
    text,
    node.fontSizePt ?? defaultFontSizePt(node.runs),
    node.lineSpacing,
    node.lineSpacingPt,
    undefined,
    undefined,
    paragraphTextWidthPx(node, ctx),
  );
}

/** Marker column plus gap the renderer reserves before list item text. */
const LIST_MARKER_PX = 24;

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
 * `spaceBefore`/`spaceAfter`, `0.5em` list/table/figure margins), so stacked
 * estimates fold with CSS margin-collapsing semantics. Exported for unit
 * tests and estimator calibration; the paginator uses it internally.
 */
export function estimateBlockHeightPx(block: FlowNode, bodyWidthPx: number): number {
  return measureBlock(block, { bodyWidthPx }).standalone;
}

function blockHeightPx(block: FlowNode, ctx: EstimateContext): BlockMetrics {
  return measureBlock(block, ctx);
}

function measureBlock(block: FlowNode, ctx: EstimateContext): BlockMetrics {
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
      // Markers sit beside their item text; item paragraphs fold with the
      // same collapsing the renderer paints. Each item wraps in the column
      // left after its numbering indent and the marker column.
      let total = 0;
      let pendingBottom = 0;
      let first = true;
      for (const item of block.items) {
        const itemCtx: EstimateContext = { bodyWidthPx: Math.max(24, ctx.bodyWidthPx - LIST_MARKER_PX) };
        const part = measureBlock(item.paragraph, itemCtx);
        total += first ? part.topPx : Math.max(pendingBottom, part.topPx);
        total += part.contentPx;
        pendingBottom = part.bottomPx;
        first = false;
      }
      if (!first) total += pendingBottom;
      return metrics(BASE_FONT_PX / 2, total, BASE_FONT_PX / 2);
    }
    case "table": {
      let total = 0;
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
          const cellCtx: EstimateContext = { bodyWidthPx: Math.max(24, Math.min(ctx.bodyWidthPx, cellWidthPx)) };
          const cellHeight = foldBlocksHeightPx(cell.nodes, cellCtx) + 8;
          if (cellHeight > rowHeight) rowHeight = cellHeight;
        }
        if (row.height !== undefined) {
          const authored = row.height * PX_PER_INCH;
          if (authored > rowHeight) rowHeight = authored;
        }
        total += rowHeight;
      }
      return metrics(BASE_FONT_PX / 2, total, BASE_FONT_PX / 2);
    }
    case "image":
      return metrics(BASE_FONT_PX / 2, block.height * PX_PER_INCH, BASE_FONT_PX / 2);
    case "footnote":
      return metrics(0, FOOTNOTE_SEPARATOR_PX + foldBlocksHeightPx(block.nodes, ctx), 0);
    case "unsupported":
      return metrics(BASE_FONT_PX / 2, 48, BASE_FONT_PX / 2);
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

function bodyBox(section: DocumentSection): PageBox {
  const { page } = section.properties;
  return {
    widthPx: twipsToPx(page.widthTwips - page.margins.left - page.margins.right),
    heightPx: twipsToPx(page.heightTwips - page.margins.top - page.margins.bottom),
  };
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
      total += empty ? part.topPx : Math.max(pendingBottom, part.topPx);
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
      const gap = current!.page.body.length === 0 ? nextPart.topPx : Math.max(current!.prevBottomPx, nextPart.topPx);
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
