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
}

export function layoutTabStops(root: HTMLElement): void {
  const paragraphs = root.querySelectorAll<HTMLElement>(`[${PARAGRAPH_ATTRIBUTE}]`);
  if (paragraphs.length === 0) return;
  const rootWidth = root.offsetWidth;
  const rootRect = root.getBoundingClientRect();
  const pass: PassContext = {
    zoom: rootWidth > 0 && rootRect.width > 0 ? rootRect.width / rootWidth : 1,
    edges: new Map(),
    glyphWidths: new Map(),
  };
  for (const line of paragraphs) layoutParagraph(line, stopsByParagraph.get(line) ?? [], pass);
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

function layoutParagraph(line: HTMLElement, stops: readonly DocxTabStop[], pass: PassContext): void {
  const parent = line.parentElement;
  if (parent === null) return;
  const tabs = [...line.querySelectorAll<HTMLElement>(`[${TAB_ATTRIBUTE}]`)];
  if (tabs.length === 0) return;
  const rtl = line.getAttribute("data-aared-direction") === "rtl";
  const edges = startEdges(parent, pass);
  const origin = rtl ? edges.rtl : edges.ltr;
  const distanceFromStart = (rect: DOMRect): number => (rtl ? origin - rect.right / pass.zoom : rect.left / pass.zoom - origin);

  // Collapse first so each tab measures against the widths of those before it.
  for (const tab of tabs) {
    tab.style.width = "0px";
    tab.textContent = "";
  }
  for (const [index, tab] of tabs.entries()) {
    const x = distanceFromStart(tab.getBoundingClientRect());
    const custom = nextCustomStop(stops, x);
    const stopPx = custom !== undefined
      ? custom.pos * PX_PER_INCH
      : (Math.floor(x / DEFAULT_TAB_PX + 1e-6) + 1) * DEFAULT_TAB_PX;
    const align = custom?.align ?? "left";
    const follow = align === "left" ? 0 : followingTextWidth(line, tab, tabs[index + 1], pass.zoom);
    const width = Math.max(0, stopPx - x - (align === "center" ? follow / 2 : follow));
    tab.style.width = `${width}px`;
    if (custom?.leader !== undefined && width > MIN_LEADER_PX) fillLeader(tab, LEADER_GLYPHS[custom.leader], width, pass);
  }
}

/** Width of the text between a tab and the next one (or the line's end). */
function followingTextWidth(line: HTMLElement, tab: HTMLElement, next: HTMLElement | undefined, zoom: number): number {
  const range = document.createRange();
  range.setStartAfter(tab);
  if (next !== undefined) range.setEndBefore(next);
  else range.setEndAfter(line.lastChild ?? tab);
  return range.getBoundingClientRect().width / zoom;
}

/** Fill a spacer with as many leader glyphs as fit, hugging the text that follows. */
function fillLeader(tab: HTMLElement, glyph: string, width: number, pass: PassContext): void {
  tab.style.whiteSpace = "pre";
  tab.style.userSelect = "none";
  tab.setAttribute("aria-hidden", "true");
  const key = `${getComputedStyle(tab).font}|${glyph}`;
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
