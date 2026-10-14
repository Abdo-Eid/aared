/**
 * Word tab stops for the DOCX renderer.
 *
 * A tab is painted as a measurable spacer, and once its paragraph is attached
 * `layoutTabStops` resolves each spacer to the width that lands the following
 * text on its stop: a left stop starts the text there, a right stop ends it
 * there, a center stop straddles it, and a leader fills the gap. Stops measure
 * from the start edge of the paragraph's containing block, so an RTL paragraph
 * measures from the right, as in Word. Layout pages that paint lazily lay out
 * their tabs page by page as each is first shown (`settlePageTabStops`).
 */

import type { DocxTabStop } from "../core/model.js";
import { LEFT_TO_RIGHT_MARK, RIGHT_TO_LEFT_MARK } from "../core/chars.js";
import { PX_PER_INCH } from "../core/units.js";

/** Word's default stops past the last custom one: every half inch. */
const DEFAULT_TAB_PX = PX_PER_INCH / 2;
/** Sub-pixel slack so a tab that already sits on a stop moves to the next one. */
const STOP_EPSILON_PX = 0.5;
/** Leaders narrower than this paint nothing. */
const MIN_LEADER_PX = 2;
/** Glyphs measured at once when sampling a leader's advance width. */
const LEADER_PROBE_LENGTH = 40;

const TAB = "\t";
/**
 * Word treats a tab as a segment separator at paragraph level, so the text
 * after it never joins the bidi run before it. The spacer is an inline-block —
 * only a neutral to the browser — so a mark in the paragraph's direction after
 * it puts that text back at paragraph level.
 */
const DIRECTION_MARK = { ltr: LEFT_TO_RIGHT_MARK, rtl: RIGHT_TO_LEFT_MARK } as const;

const LEADER_GLYPHS: Record<NonNullable<DocxTabStop["leader"]>, string> = {
  dot: ".",
  hyphen: "-",
  underscore: "_",
  middleDot: "·",
};

const TAB_ATTRIBUTE = "data-aared-tab";
const PARAGRAPH_ATTRIBUTE = "data-aared-tab-stops";

/**
 * Stops per paragraph element. Held off the DOM so a stop list shared by many
 * paragraphs is never serialized onto each one, and freed with the element.
 */
const stopsByParagraph = new WeakMap<HTMLElement, readonly DocxTabStop[]>();

/**
 * Paint run text, turning each tab character into a spacer. A spacer starts at
 * the default stop width so tabs stay visible if no layout pass ever runs.
 */
export function fillRunText(span: HTMLElement, text: string, direction: "ltr" | "rtl"): void {
  if (!text.includes(TAB)) {
    span.textContent = text;
    return;
  }
  const parts = text.split(TAB);
  for (const [index, part] of parts.entries()) {
    if (index > 0) {
      const tab = document.createElement("span");
      tab.setAttribute(TAB_ATTRIBUTE, "");
      tab.style.display = "inline-block";
      tab.style.width = `${DEFAULT_TAB_PX}px`;
      span.append(tab, document.createTextNode(DIRECTION_MARK[direction]));
    }
    if (part !== "") span.append(document.createTextNode(part));
  }
}

/** Record a paragraph's stops when it holds tabs, so the layout pass finds it. */
export function markTabs(element: HTMLElement, stops: readonly DocxTabStop[] | undefined): void {
  if (element.querySelector(`[${TAB_ATTRIBUTE}]`) === null) return;
  element.setAttribute(PARAGRAPH_ATTRIBUTE, "");
  stopsByParagraph.set(element, stops ?? []);
}

/**
 * Lay out the tabs under `root` now, and again once web fonts land (they change
 * advance widths after first paint). Renders without tabs schedule nothing.
 */
export function settleTabStops(root: HTMLElement, alive: () => boolean): void {
  if (root.querySelector(`[${TAB_ATTRIBUTE}]`) === null) return;
  layoutTabStops(root);
  void document.fonts?.ready.then(() => {
    if (alive()) layoutTabStops(root);
  });
}

/** Layout pages can paint lazily here: `content-visibility: auto` and its state event are both supported. */
export function lazyPagePaint(): boolean {
  return (
    typeof ContentVisibilityAutoStateChangeEvent === "function" &&
    typeof CSS !== "undefined" &&
    CSS.supports("content-visibility", "auto")
  );
}

/**
 * The tab pass for lazily painted pages (`content-visibility: auto`): lay out
 * a page's tabs the first time it becomes relevant, never by forcing layout
 * of pages the browser is skipping. When a web font finishes loading (which
 * changes advance widths), pages on screen re-run and the rest are marked to
 * re-run when next shown; before printing, every page not laid out with the
 * current fonts runs, since print shows them all.
 */
export function settlePageTabStops(pages: readonly HTMLElement[], alive: () => boolean): () => void {
  const tabbed = pages.filter((page) => page.querySelector(`[${TAB_ATTRIBUTE}]`) !== null);
  if (tabbed.length === 0) return () => {};
  /** Bumped on every font load; a page laid out under an older one is stale. */
  let fontEpoch = 0;
  const laidOut = new WeakMap<HTMLElement, number>();
  const shown = new Set<HTMLElement>();
  const run = (page: HTMLElement): void => {
    layoutTabStops(page);
    laidOut.set(page, fontEpoch);
  };
  const show = (page: HTMLElement): void => {
    shown.add(page);
    if (laidOut.get(page) !== fontEpoch) run(page);
  };
  const onFonts = (): void => {
    if (!alive()) return;
    fontEpoch += 1;
    for (const page of shown) run(page);
  };
  document.fonts?.addEventListener("loadingdone", onFonts);
  const onPrint = (): void => {
    if (!alive()) return;
    for (const page of tabbed) if (laidOut.get(page) !== fontEpoch) run(page);
  };
  window.addEventListener("beforeprint", onPrint);
  for (const page of tabbed) {
    page.addEventListener("contentvisibilityautostatechange", (event) => {
      if (!alive()) return;
      if ((event as ContentVisibilityAutoStateChangeEvent).skipped) shown.delete(page);
      else show(page);
    });
  }
  // Pages relevant from the start are not guaranteed a first event: lay those
  // out on the next frame. A skipped page's box still renders, so it is the
  // content that is checked.
  requestAnimationFrame(() => {
    if (!alive()) return;
    for (const page of tabbed) {
      const content = page.firstElementChild;
      if (!laidOut.has(page) && content !== null && content.checkVisibility({ contentVisibilityAuto: true })) show(page);
    }
  });
  // The listeners live on the document and window: the caller's disposer takes them off.
  return () => {
    document.fonts?.removeEventListener("loadingdone", onFonts);
    window.removeEventListener("beforeprint", onPrint);
  };
}

/** The paragraph start edge, in unzoomed pixels, for each text direction. */
interface StartEdges {
  ltr: number;
  rtl: number;
}

/** Everything one pass measures once and shares across its paragraphs. */
interface PassContext {
  /** Ratio of painted to layout pixels (CSS `zoom` on an ancestor). */
  zoom: number;
  edges: Map<HTMLElement, StartEdges>;
  /** Leader glyph advance per font, so a font is probed once, not once per tab. */
  glyphWidths: Map<string, number>;
  /**
   * The one range every following-text read reuses. A range stays registered
   * with its document until it is collected, and each removal of a child (a
   * collapsing leader) walks all of them: a range per tab made a second pass
   * over a document with thousands of tabs take a minute.
   */
  range: Range;
}

export function layoutTabStops(root: HTMLElement): void {
  layoutTabParagraphs(root, Array.from(root.querySelectorAll<HTMLElement>(`[${PARAGRAPH_ATTRIBUTE}]`)));
}

/** A pixel of slack: a tab that landed this close to what it is now asked for is settled. */
const SETTLED_PX = 1;
/** Check-and-repair rounds after the batched pass; a document that still moves after these keeps its last layout. */
const REPAIR_ROUNDS = 3;

/** One paragraph's tabs, with what the pass measured once and the width each tab was given. */
interface TabJob {
  line: HTMLElement;
  stops: readonly DocxTabStop[];
  tabs: HTMLElement[];
  rtl: boolean;
  /** The paragraph's start edge, in unzoomed pixels. */
  origin: number;
  applied: number[];
}

/**
 * Resolve the tabs of `paragraphs` (under `root`) in rounds, not one paragraph
 * at a time. Reading a rect after writing a width forces a layout, and per
 * paragraph that costs one layout for every tabbed paragraph, each of them
 * walking the whole tree: a long document with a thousand tabs (a table of
 * contents, a table of authorities) paid minutes. Here every start edge is read,
 * then every tab collapses, then each round reads the k-th tab of every
 * paragraph and writes their widths, so a pass costs a few layouts however many
 * paragraphs there are.
 *
 * Within a paragraph the order is the one Word has (a tab measures against the
 * tabs before it). Across paragraphs it can differ in one case: paragraphs that
 * share a formatting context with floats, where a line that grows above moves
 * what a paragraph below runs beside. A final read-only check finds any tab
 * that landed off what it is now asked for and resolves those paragraphs one at
 * a time, top down, as before, so the result is the same either way.
 */
export function layoutTabParagraphs(root: HTMLElement, paragraphs: readonly HTMLElement[]): void {
  if (paragraphs.length === 0) return;
  const rootWidth = root.offsetWidth;
  const rootRect = root.getBoundingClientRect();
  const pass: PassContext = {
    zoom: rootWidth > 0 && rootRect.width > 0 ? rootRect.width / rootWidth : 1,
    edges: new Map(),
    glyphWidths: new Map(),
    range: document.createRange(),
  };
  // Reads first, before any tab moves.
  const jobs: TabJob[] = [];
  for (const line of paragraphs) {
    const parent = line.parentElement;
    if (parent === null) continue;
    const tabs = [...line.querySelectorAll<HTMLElement>(`[${TAB_ATTRIBUTE}]`)];
    if (tabs.length === 0) continue;
    const rtl = line.getAttribute("data-aared-direction") === "rtl";
    const edges = startEdges(parent, pass);
    jobs.push({ line, stops: stopsByParagraph.get(line) ?? [], tabs, rtl, origin: rtl ? edges.rtl : edges.ltr, applied: [] });
  }
  // Collapse first so each tab measures against the widths of those before it.
  for (const job of jobs) collapseTabs(job);
  for (let round = 0; ; round++) {
    const live = jobs.filter((job) => job.tabs.length > round);
    if (live.length === 0) break;
    const plans = live.map((job) => planTab(job, round, pass));
    live.forEach((job, index) => applyTab(job, round, plans[index]!, pass));
  }
  for (let attempt = 0; attempt < REPAIR_ROUNDS; attempt++) {
    const moved = jobs.filter((job) => job.tabs.some((_, index) => Math.abs(planTab(job, index, pass).width - job.applied[index]!) > SETTLED_PX));
    if (moved.length === 0) return;
    for (const job of moved) resolveInOrder(job, pass);
  }
}

function startEdges(parent: HTMLElement, pass: PassContext): StartEdges {
  let edges = pass.edges.get(parent);
  if (edges === undefined) {
    const rect = parent.getBoundingClientRect();
    const style = getComputedStyle(parent);
    const inset = (side: "Left" | "Right"): number =>
      (Number.parseFloat(style[`padding${side}`]) || 0) + (Number.parseFloat(style[`border${side}Width`]) || 0);
    edges = { ltr: rect.left / pass.zoom + inset("Left"), rtl: rect.right / pass.zoom - inset("Right") };
    pass.edges.set(parent, edges);
  }
  return edges;
}

/** The first custom stop ahead of `x`, or `undefined` when the default grid takes over. */
function nextCustomStop(stops: readonly DocxTabStop[], x: number): DocxTabStop | undefined {
  return stops.find((stop) => stop.pos * PX_PER_INCH > x + STOP_EPSILON_PX);
}

function collapseTabs(job: TabJob): void {
  for (const tab of job.tabs) {
    tab.style.width = "0px";
    if (tab.firstChild !== null) tab.textContent = "";
  }
}

/** What a tab is asked for: its width, and a leader glyph with the font to sample it in when one fits. */
interface TabPlan {
  width: number;
  leader: string | undefined;
  font: string | undefined;
}

/**
 * The width tab `index` of a paragraph is asked for where it sits now; reads
 * only, the leader's computed font included, because a style read after the
 * previous tab's writes recalculates every style in the document.
 */
function planTab(job: TabJob, index: number, pass: PassContext): TabPlan {
  const { zoom } = pass;
  const tab = job.tabs[index]!;
  const rect = tab.getBoundingClientRect();
  const x = job.rtl ? job.origin - rect.right / zoom : rect.left / zoom - job.origin;
  const custom = nextCustomStop(job.stops, x);
  const stopPx = custom !== undefined
    ? custom.pos * PX_PER_INCH
    : (Math.floor(x / DEFAULT_TAB_PX + 1e-6) + 1) * DEFAULT_TAB_PX;
  const align = custom?.align ?? "left";
  const follow = align === "left" ? 0 : followingTextWidth(pass.range, job.line, tab, job.tabs[index + 1], zoom);
  const width = Math.max(0, stopPx - x - (align === "center" ? follow / 2 : follow));
  const leader = custom?.leader !== undefined && width > MIN_LEADER_PX ? LEADER_GLYPHS[custom.leader] : undefined;
  return { width, leader, font: leader === undefined ? undefined : getComputedStyle(tab).font };
}

function applyTab(job: TabJob, index: number, plan: TabPlan, pass: PassContext): void {
  const tab = job.tabs[index]!;
  job.applied[index] = plan.width;
  tab.style.width = `${plan.width}px`;
  if (plan.leader !== undefined) fillLeader(tab, plan.leader, plan.font!, plan.width, pass);
}

/** One paragraph on its own, each tab measured after those before it were written: Word's order, at a layout per tab. */
function resolveInOrder(job: TabJob, pass: PassContext): void {
  collapseTabs(job);
  for (let index = 0; index < job.tabs.length; index++) applyTab(job, index, planTab(job, index, pass), pass);
}

/** Width of the text between a tab and the next one (or the line's end). */
function followingTextWidth(range: Range, line: HTMLElement, tab: HTMLElement, next: HTMLElement | undefined, zoom: number): number {
  range.setStartAfter(tab);
  if (next !== undefined) range.setEndBefore(next);
  else range.setEndAfter(line.lastChild ?? tab);
  return range.getBoundingClientRect().width / zoom;
}

/** Fill a spacer with as many leader glyphs as fit, hugging the text that follows. */
function fillLeader(tab: HTMLElement, glyph: string, font: string, width: number, pass: PassContext): void {
  tab.style.whiteSpace = "pre";
  tab.style.userSelect = "none";
  tab.setAttribute("aria-hidden", "true");
  const key = `${font}|${glyph}`;
  let glyphPx = pass.glyphWidths.get(key);
  if (glyphPx === undefined) {
    // Sample the advance with the box shrink-wrapped to its content.
    tab.style.width = "auto";
    tab.textContent = glyph.repeat(LEADER_PROBE_LENGTH);
    glyphPx = tab.getBoundingClientRect().width / pass.zoom / LEADER_PROBE_LENGTH;
    pass.glyphWidths.set(key, glyphPx);
  }
  tab.textContent = glyphPx > 0 ? glyph.repeat(Math.floor(width / glyphPx)) : "";
  // Any remainder pads the start, so the last glyph meets the following text.
  tab.style.textAlign = "end";
  tab.style.width = `${width}px`;
}
