/**
 * Text down both sides of a picture, line by line: Word's `bothSides` wrap for a
 * picture with room on each side. A CSS float wraps one side only, so a
 * paragraph with such a picture is laid out here instead, after it is attached
 * (fonts and widths are known then):
 *
 *   words -> measure (canvas, cached per word and font) -> fill each line's
 *   slots in reading order (`fitUnits`) -> rows of real DOM text, one per line
 *   beside the picture -> the rest of the words as ordinary paragraph text
 *
 * The text stays browser text: canvas only measures. A painted-width check
 * catches any slot the canvas measure got wrong and narrows it. The
 * paragraph's original run spans stay in the DOM, hidden, as the source the
 * layout is rebuilt from when a width or a font changes. Every paragraph under
 * a root is laid out in three passes, reads never interleaved with writes:
 * read all (styles, widths, word widths), build all, check all once.
 */

import { cssClass } from "../class-sheet.js";
import type { DocxImageNode } from "../core/model.js";
import { PX_PER_INCH } from "../core/units.js";
import { planFloat, polygonSpan } from "./float-layout.js";
import type { AnchorFrame, FloatPlan } from "./float-layout.js";
import { fitUnits } from "./split-fit.js";
import type { Slot } from "./split-fit.js";

/** One line beside the picture: a grid of the slots and the hole between them. */
const ROW = cssClass("display:grid;white-space:nowrap");
/** A slot of a row; `min-width:0` so a long word cannot stretch its column. */
const CELL = cssClass("min-width:0");

/** Breakable white space. A no-break space (U+00A0) is not in it: it joins its words, as in Word. */
const BREAK_SPACE = /([ \t\r\n\f]+)/;
const ONLY_BREAK_SPACE = /^[ \t\r\n\f]+$/;

/** Narrower slack steps tried when a slot's painted text is wider than the canvas measured. */
const SLACK_STEPS_PX = [1, 2, 4];

interface SplitState {
  image: DocxImageNode;
  frame: AnchorFrame;
  leftIndentPx: number;
  rightIndentPx: number;
  picture: HTMLElement;
  /** The paragraph's original run spans, hidden; created on first layout. */
  source?: HTMLElement;
}

const states = new WeakMap<HTMLElement, SplitState>();

/** Whether every run of a paragraph is plain text a split can move between slots. */
export function splittable(runs: ReadonlyArray<{ image?: unknown; hyperlink?: unknown; footnoteRef?: unknown; field?: unknown; text: string }>): boolean {
  return runs.length > 0 && runs.every((run) =>
    run.image === undefined && run.hyperlink === undefined && run.footnoteRef === undefined && run.field === undefined
    && !run.text.includes("\t") && !run.text.includes("\n"));
}

/** Mark `line` for a split layout; `settleSplitWraps` does it once the paragraph is on the page. */
export function registerSplit(line: HTMLElement, state: SplitState): void {
  states.set(line, state);
  line.dataset.aaredSplit = "true";
  line.style.position = "relative";
}

/** A word: one or more text parts (a word can change run mid-way) and its measured width. */
interface Unit {
  parts: Array<{ text: string; run: HTMLElement }>;
  width: number;
}

interface RunMeasure {
  font: string;
  letterSpacing: string;
  upper: boolean;
}

const canvas = typeof document === "undefined" ? undefined : document.createElement("canvas");
const context = canvas?.getContext("2d") ?? undefined;
/** Word widths per font: a paragraph repeats its words, and so do its rebuilds. */
const widthCache = new Map<string, number>();
const WIDTH_CACHE_LIMIT = 20000;

function measureText(text: string, measure: RunMeasure): number {
  const key = `${measure.font}|${measure.letterSpacing}|${measure.upper ? "U" : ""}|${text}`;
  const hit = widthCache.get(key);
  if (hit !== undefined) return hit;
  if (context === undefined) return 0;
  context.font = measure.font;
  if ("letterSpacing" in context) (context as unknown as { letterSpacing: string }).letterSpacing = measure.letterSpacing;
  const width = context.measureText(measure.upper ? text.toUpperCase() : text).width;
  widthCache.set(key, width);
  // Bounded: a very long document must not grow this without end.
  if (widthCache.size > WIDTH_CACHE_LIMIT) widthCache.delete(widthCache.keys().next().value as string);
  return width;
}

function runMeasure(element: HTMLElement): RunMeasure {
  const style = getComputedStyle(element);
  return {
    font: `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`,
    letterSpacing: style.letterSpacing === "normal" ? "0px" : style.letterSpacing,
    upper: style.textTransform === "uppercase",
  };
}

/** The paragraph's words, from its hidden source spans (one style read per run). */
function units(source: HTMLElement): Unit[] {
  const list: Unit[] = [];
  let open: Unit | undefined;
  for (const child of Array.from(source.children) as HTMLElement[]) {
    const measure = runMeasure(child);
    for (const piece of (child.textContent ?? "").split(BREAK_SPACE)) {
      if (piece === "") continue;
      if (ONLY_BREAK_SPACE.test(piece)) {
        open = undefined;
        continue;
      }
      if (open === undefined) {
        open = { parts: [], width: 0 };
        list.push(open);
      }
      open.parts.push({ text: piece, run: child });
      open.width += measureText(piece, measure);
    }
  }
  return list;
}

/** `units` as run-shaped DOM: consecutive words of one run share a clone of that run's span. */
function paint(target: HTMLElement, taken: readonly Unit[]): void {
  let current: { run: HTMLElement; element: HTMLElement; text: string } | undefined;
  const flush = (): void => {
    if (current !== undefined) current.element.textContent = current.text;
    current = undefined;
  };
  taken.forEach((unit, index) => {
    unit.parts.forEach((part, partIndex) => {
      const text = index > 0 && partIndex === 0 ? ` ${part.text}` : part.text;
      if (current === undefined || current.run !== part.run) {
        flush();
        const element = part.run.cloneNode(false) as HTMLElement;
        target.append(element);
        current = { run: part.run, element, text: "" };
      }
      current.text += text;
    });
  });
  flush();
}

function clearRows(paragraph: HTMLElement): void {
  paragraph.querySelectorAll("[data-aared-split-row], [data-aared-split-rest]").forEach((node) => node.remove());
}

/** Move the paragraph's run spans into a hidden source, once. A write. */
function ensureSource(paragraph: HTMLElement, state: SplitState): HTMLElement {
  if (state.source !== undefined) return state.source;
  const source = document.createElement("span");
  source.hidden = true;
  source.dataset.aaredSplitSource = "true";
  for (const child of Array.from(paragraph.childNodes)) if (child !== state.picture) source.append(child);
  paragraph.append(source);
  state.source = source;
  return source;
}

/** Undo the split: the source spans show again and the picture sits above them as a block. */
function restore(paragraph: HTMLElement, state: SplitState): void {
  clearRows(paragraph);
  if (state.source !== undefined) state.source.hidden = false;
  paragraph.style.minHeight = "";
  state.picture.style.position = "static";
  state.picture.style.display = "block";
  state.picture.style.margin = "0 auto 8px";
  paragraph.prepend(state.picture);
}

type SplitPlan = Extract<FloatPlan, { kind: "split" }>;

/** Everything a split layout reads from the page, gathered before anything is written. */
interface Prepared {
  paragraph: HTMLElement;
  state: SplitState;
  plan: FloatPlan;
  contentPx: number;
  lineHeightPx: number;
  /** The paragraph's spacing after, which the picture may hang into. */
  spaceAfterPx: number;
  rtl: boolean;
  justify: boolean;
  list: Unit[];
  widths: number[];
  spacePx: number;
}

/** The read pass for one paragraph. */
function prepare(paragraph: HTMLElement, state: SplitState): Prepared {
  const style = getComputedStyle(paragraph);
  const contentPx = paragraph.clientWidth > 0 ? paragraph.clientWidth : state.frame.columnPx - state.leftIndentPx - state.rightIndentPx;
  const plan = planFloat(state.image, { ...state.frame, columnPx: contentPx + state.leftIndentPx + state.rightIndentPx }, state.leftIndentPx);
  const list = units(state.source!);
  const fontPx = Number.parseFloat(style.fontSize);
  return {
    paragraph,
    state,
    plan,
    contentPx,
    lineHeightPx: Number.parseFloat(style.lineHeight) || fontPx * 1.2,
    spaceAfterPx: Number.parseFloat(style.marginBottom) || 0,
    rtl: style.direction === "rtl",
    justify: style.textAlign === "justify",
    list,
    widths: list.map((unit) => unit.width),
    spacePx: measureText(" ", runMeasure(paragraph)),
  };
}

/** A painted slot, for the check pass: the cell and the width it had to fit in. */
interface PaintedSlot {
  cell: HTMLElement;
  widthPx: number;
}

/** The write pass for one paragraph: rows beside the picture, then the rest. Returns its slots. */
function build(prepared: Prepared & { plan: SplitPlan }, slack: number): PaintedSlot[] {
  const { paragraph, state, plan, contentPx, lineHeightPx, spaceAfterPx, rtl, justify, list, widths, spacePx } = prepared;
  const anchor = state.image.anchor!;
  const width = state.image.width * PX_PER_INCH;
  const height = state.image.height * PX_PER_INCH;
  const dist = {
    top: anchor.distance.top * PX_PER_INCH,
    bottom: anchor.distance.bottom * PX_PER_INCH,
    left: anchor.distance.left * PX_PER_INCH,
    right: anchor.distance.right * PX_PER_INCH,
  };
  clearRows(paragraph);
  const pictureLeft = plan.leftPx - state.leftIndentPx;
  const picture = state.picture;
  picture.style.position = "absolute";
  picture.style.display = "";
  picture.style.margin = "";
  picture.style.left = `${pictureLeft}px`;
  picture.style.top = `${plan.topPx}px`;
  const bandTop = Math.max(0, plan.topPx - dist.top);
  const bandBottom = plan.topPx + height + dist.bottom;
  // The picture takes no room of its own. Word starts the next paragraph where
  // it would start anyway and wraps it around what of the picture remains; the
  // split cannot reach into the next paragraph, so the paragraph reaches just
  // far enough that the next one starts below the picture. Its spacing after
  // and the wrap distance below may overlap the picture, as in Word.
  paragraph.style.minHeight = `${Math.max(0, plan.topPx + height - spaceAfterPx)}px`;
  const above = Math.floor(bandTop / lineHeightPx + 1e-6);
  const beside = Math.max(1, Math.ceil(bandBottom / lineHeightPx - 1e-6) - above);
  const boxHole = { left: pictureLeft - dist.left, width: width + dist.left + dist.right };
  // Tight and through follow the picture's outline: each line's hole is what the polygon reaches at its height.
  const polygon = anchor.polygon !== undefined && (anchor.wrap === "tight" || anchor.wrap === "through")
    ? anchor.polygon.map(([x, y]): [number, number] => [pictureLeft + x * width, plan.topPx + y * height])
    : undefined;

  const rows: HTMLElement[] = [];
  const painted: PaintedSlot[] = [];
  let next = 0;
  for (let index = 0; index < above + beside && next < list.length; index++) {
    let isSplit = index >= above;
    let hole = boxHole;
    if (isSplit && polygon !== undefined) {
      const span = polygonSpan(polygon, index * lineHeightPx, (index + 1) * lineHeightPx);
      // A line the outline does not reach runs the full width.
      if (span === undefined) isSplit = false;
      else hole = { left: span[0] - dist.left, width: span[1] - span[0] + dist.left + dist.right };
    }
    const leftSlotPx = hole.left;
    const rightSlotPx = contentPx - (hole.left + hole.width);
    const slots: Slot[] = isSplit
      ? (rtl ? [rightSlotPx, leftSlotPx] : [leftSlotPx, rightSlotPx]).map((px) => ({ widthPx: px - slack }))
      : [{ widthPx: contentPx - slack, force: true }];
    const counts = fitUnits(widths, spacePx, slots, next);
    const row = document.createElement("div");
    row.dataset.aaredSplitRow = isSplit ? "beside" : "above";
    row.setAttribute("class", ROW);
    row.style.gridTemplateColumns = isSplit
      ? (rtl ? `${rightSlotPx}px ${hole.width}px ${leftSlotPx}px` : `${leftSlotPx}px ${hole.width}px ${rightSlotPx}px`)
      : "1fr";
    slots.forEach((slot, slotIndex) => {
      const cell = document.createElement("span");
      cell.setAttribute("class", CELL);
      if (isSplit && slotIndex === 1) cell.style.gridColumn = "3";
      const taken = list.slice(next, next + counts[slotIndex]!);
      next += counts[slotIndex]!;
      // Word justifies each segment of a line to its edges, except the paragraph's last line.
      if (justify && next < list.length) cell.style.textAlignLast = "justify";
      paint(cell, taken);
      row.append(cell);
      painted.push({ cell, widthPx: slot.widthPx + slack });
    });
    rows.push(row);
  }
  const rest = document.createElement("span");
  rest.dataset.aaredSplitRest = "true";
  paint(rest, list.slice(next));
  paragraph.append(...rows, rest);
  return painted;
}

/** The check pass for one paragraph: whether any slot's painted text is wider than the slot. */
function overflows(paragraph: HTMLElement, slots: readonly PaintedSlot[], range: Range): boolean {
  // Not laid out (a page the browser is skipping): the canvas measure stands.
  if (paragraph.getBoundingClientRect().height <= 0) return false;
  return slots.some(({ cell, widthPx }) => {
    range.selectNodeContents(cell);
    return range.getBoundingClientRect().width > widthPx + 0.5;
  });
}

/** Lay out `paragraphs` in three passes: read all, build all, check all (and narrow the few that overflow). */
function layoutAll(paragraphs: readonly HTMLElement[]): void {
  const entries: Array<{ paragraph: HTMLElement; state: SplitState }> = [];
  for (const paragraph of paragraphs) {
    const state = states.get(paragraph);
    if (state === undefined || !paragraph.isConnected) continue;
    ensureSource(paragraph, state);
    entries.push({ paragraph, state });
  }
  const prepared = entries.map(({ paragraph, state }) => prepare(paragraph, state));
  const built: Array<{ prepared: Prepared & { plan: SplitPlan }; slots: PaintedSlot[] }> = [];
  for (const entry of prepared) {
    const { plan } = entry;
    if (plan.kind === "split") {
      const ready = { ...entry, plan };
      built.push({ prepared: ready, slots: build(ready, 0) });
    } else {
      // The container has grown too narrow for text on both sides.
      restore(entry.paragraph, entry.state);
    }
  }
  const range = document.createRange();
  let pending = built;
  for (const slack of SLACK_STEPS_PX) {
    pending = pending.filter(({ prepared: entry, slots }) => overflows(entry.paragraph, slots, range));
    if (pending.length === 0) return;
    for (const item of pending) item.slots = build(item.prepared, slack);
  }
}

/**
 * Lay out every split paragraph under `root` once it is on the page; again
 * when fonts still loading finish (`refont`); and, with `observeWidth`, when
 * the root's width changes (flow mode: the column follows the container).
 * Returns a disposer that stops both.
 */
export function settleSplitWraps(root: HTMLElement, alive: () => boolean, options: { observeWidth?: boolean; refont?: boolean } = {}): () => void {
  const paragraphs = Array.from(root.querySelectorAll<HTMLElement>("[data-aared-split]"));
  if (paragraphs.length === 0) return () => {};
  let disposed = false;
  const run = (): void => {
    if (!disposed && alive()) layoutAll(paragraphs);
  };
  run();
  // Fonts already loaded need no second pass; only a face still arriving changes the widths.
  if (options.refont !== false && typeof document !== "undefined" && document.fonts !== undefined && document.fonts.status !== "loaded") {
    void document.fonts.ready.then(run);
  }
  let observer: ResizeObserver | undefined;
  if (options.observeWidth === true && typeof ResizeObserver === "function") {
    let width = root.clientWidth;
    observer = new ResizeObserver(() => {
      if (root.clientWidth === width) return;
      width = root.clientWidth;
      run();
    });
    observer.observe(root);
  }
  return () => {
    disposed = true;
    observer?.disconnect();
  };
}
