/**
 * Paints a search query over the rendered DOM without wrapping runs or
 * changing the text nodes that renderer layout, bidi and split-wrap logic uses.
 * One controller belongs to one rendered viewer root.
 */

import { searchKey } from "./core/search.js";
import { AaredError } from "./core/errors.js";
import { resourceId } from "./resource-id.js";

/** One viewer owns the controller's lifetime. */
export interface SearchHighlighter {
  readonly count: number;
  readonly index: number | null;
  /**
   * The painted occurrences in `activate()` order, each with the block it
   * sits in and that block's search key, so a caller can pick the one a
   * source result names. Read after the query is set; valid until the DOM changes.
   */
  readonly occurrences: readonly PaintedOccurrence[];
  setQuery(query: string): void;
  activate(index: number, scroll?: boolean): void;
  activateWithin(selector: string): void;
  destroy(): void;
}

/** One painted occurrence: its text range, the paragraph/cell/SVG text block holding it, and that block's `searchKey`. */
export interface PaintedOccurrence {
  readonly range: Range;
  readonly block: Element;
  readonly key: string;
}

interface HighlightLike {}
interface HighlightRegistryLike {
  set(name: string, highlight: HighlightLike): void;
  delete(name: string): boolean;
}

interface HighlightWindow extends Window {
  CSS?: { highlights?: HighlightRegistryLike };
  Highlight?: new (...ranges: Range[]) => HighlightLike;
  MutationObserver: typeof MutationObserver;
}

interface TextSegment {
  node: Text;
  start: number;
  end: number;
}

interface SourcePosition {
  start: number;
  end: number;
}

const BLOCK_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,[role="cell"],[data-aared-element="cell-text"],svg text';

/** True for renderer source/decoration text that should never be searched. */
function excludedTextNode(node: Text, block: Element): boolean {
  if (block.closest('[hidden], [aria-hidden="true"], [data-aared-split-source]') !== null) return true;
  let element = node.parentElement;
  while (element !== null) {
    if (
      element.hasAttribute("hidden") ||
      element.getAttribute("aria-hidden")?.toLowerCase() === "true" ||
      element.hasAttribute("data-aared-split-source") ||
      element.tagName === "SCRIPT" ||
      element.tagName === "STYLE"
    ) return true;
    // A paragraph may contain an anchored chart; its SVG text is a separate
    // searchable block and must not also be counted as paragraph text.
    if (element !== block && element.namespaceURI === "http://www.w3.org/2000/svg") return true;
    // A cell is a fallback block only for any direct text it owns. Paragraphs
    // and headings inside it are processed independently.
    if (element !== block && element.matches("p,h1,h2,h3,h4,h5,h6,[role=cell],svg text")) return true;
    if (element === block) break;
    element = element.parentElement;
  }
  return false;
}

/**
 * Collect visible text in one logical paragraph/cell/SVG text block. Text
 * nodes are joined for matching so a query can cross Word run spans or SVG
 * tspans, while segments retain the original UTF-16 positions for DOM Ranges.
 */
function readBlock(block: Element): { text: string; segments: TextSegment[] } {
  const segments: TextSegment[] = [];
  // PPTX paragraphs keep bullets outside their searchable text span.
  const content = block.querySelector('[data-aared-element="paragraph-text"]') ?? block;
  let text = "";
  const walker = block.ownerDocument.createTreeWalker(content, 4 /* SHOW_TEXT */);
  let current = walker.nextNode();
  while (current !== null) {
    const node = current as Text;
    if (!excludedTextNode(node, block) && node.data.length > 0) {
      const start = text.length;
      text += node.data;
      segments.push({ node, start, end: text.length });
    }
    current = walker.nextNode();
  }
  return { text, segments };
}

/** Find a source text-node boundary, preferring the correct side of a run boundary. */
function boundaryAt(segments: readonly TextSegment[], offset: number, isEnd: boolean): { node: Text; offset: number } | undefined {
  for (const segment of segments) {
    if (isEnd ? offset > segment.start && offset <= segment.end : offset >= segment.start && offset < segment.end) {
      return { node: segment.node, offset: offset - segment.start };
    }
  }
  return undefined;
}

/**
 * Create source-offset entries for the folded key. A code point can fold to
 * nothing (Arabic tashkeel/tatweel), or to several code units (presentation
 * forms and Unicode case mappings); every resulting key unit points back to
 * that code point's original UTF-16 span. `searchKey` itself remains the
 * authority for the full haystack and query, keeping highlight matching in
 * lockstep with the public search rules.
 */
function sourceMap(text: string, key: string): SourcePosition[] {
  const positions: SourcePosition[] = [];
  for (let at = 0; at < text.length;) {
    const codePoint = String.fromCodePoint(text.codePointAt(at)!);
    const end = at + codePoint.length;
    const folded = searchKey(codePoint);
    for (let index = 0; index < folded.length; index++) positions.push({ start: at, end });
    at = end;
  }

  // When NFKC composes adjacent source characters, prefix folds recover the
  // source interval instead of painting unrelated text across the whole block.
  if (positions.length !== key.length) {
    const composed: SourcePosition[] = [];
    for (let at = 0; at < text.length;) {
      const end = at + String.fromCodePoint(text.codePointAt(at)!).length;
      const before = searchKey(text.slice(0, at)).length;
      const after = searchKey(text.slice(0, end)).length;
      for (let index = before; index < after; index++) composed[index] = { start: at, end };
      if (after <= before && after > 0 && composed[after - 1] !== undefined) composed[after - 1]!.end = end;
      at = end;
    }
    return composed;
  }
  return positions;
}

/** Build a CSS Custom Highlight controller for all rendered copies of matches. */
export function createSearchHighlighter(root: HTMLElement, changed: () => void = () => {}): SearchHighlighter {
  const view = root.ownerDocument.defaultView as HighlightWindow | null;
  const registry = view?.CSS?.highlights;
  const HighlightConstructor = view?.Highlight;
  const name = resourceId("search");
  // The rule lives in a constructed stylesheet adopted onto the viewer's scope,
  // not in a `<style>` element: a strict `style-src` blocks style elements
  // without a nonce, and aared cannot know the page's nonce. CSSOM is allowed.
  // The root realm's own constructor, so a viewer inside an iframe works.
  const realm = root.ownerDocument.defaultView as (Window & typeof globalThis) | null;
  const sheet = realm === null || typeof realm.CSSStyleSheet !== "function" ? null : new realm.CSSStyleSheet();
  sheet?.replaceSync(`::highlight(${name}){background-color:rgba(245,158,11,.22);color:inherit}`);
  let adoptedOn: Document | ShadowRoot | null = null;
  // ::highlight allows colors/text decoration, not borders (MDN ::highlight).
  // These non-interactive rectangles
  // outline the active Range without wrapping text or affecting its layout.
  const outline = root.ownerDocument.createElement("div");
  outline.dataset.aaredSearchActive = name;
  outline.setAttribute("aria-hidden", "true");
  outline.style.cssText = "position:fixed;inset:0;pointer-events:none;z-index:2147483646";
  const attachStyle = (): void => {
    const scope = root.getRootNode();
    const shadow = scope.nodeType === 11 && "host" in scope ? scope as ShadowRoot : null;
    const owner: Document | ShadowRoot = shadow ?? root.ownerDocument;
    if (adoptedOn !== owner) {
      if (adoptedOn !== null && sheet !== null) adoptedOn.adoptedStyleSheets = adoptedOn.adoptedStyleSheets.filter((item) => item !== sheet);
      adoptedOn = owner;
    }
    // Restore the lifecycle-owned sheet if an embedder replaced the list.
    // Do not cache it in constantSheet: each highlighter has a unique name.
    if (sheet !== null && !owner.adoptedStyleSheets.includes(sheet)) owner.adoptedStyleSheets = [...owner.adoptedStyleSheets, sheet];
    const overlayHost = shadow ?? root.ownerDocument.body;
    if (outline.parentNode !== overlayHost) overlayHost.append(outline);
  };
  attachStyle();
  root.setAttribute("data-aared-search-highlight", name);

  let queryKey: string | null = null;
  let destroyed = false;
  let ranges: Range[] = [];
  /** Parallel to `ranges`: the block each sits in and its key. */
  let owners: Array<{ block: Element; key: string }> = [];
  let activeIndex: number | null = null;
  let activeSelector: string | null = null;
  let frame = 0;

  const paintActive = (): void => {
    outline.replaceChildren();
    const range = activeIndex === null ? undefined : ranges[activeIndex];
    if (destroyed || !range?.startContainer.isConnected || view === null) return;
    let left = 0, top = 0, right = view.innerWidth, bottom = view.innerHeight;
    // Respect the viewer and each clipping/scrolling ancestor, so the outline
    // cannot float over the toolbar when its match scrolls out of view.
    for (let element: Element | null = root; element !== null; element = element.parentElement ?? (element.getRootNode() as ShadowRoot).host ?? null) {
      const css = view.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      if (element === root || css.overflowX !== "visible") { left = Math.max(left, rect.left); right = Math.min(right, rect.right); }
      if (element === root || css.overflowY !== "visible") { top = Math.max(top, rect.top); bottom = Math.min(bottom, rect.bottom); }
    }
    for (const rect of range.getClientRects()) {
      const x = Math.max(left, rect.left - 2), y = Math.max(top, rect.top - 1);
      const width = Math.min(right, rect.right + 2) - x, height = Math.min(bottom, rect.bottom + 1) - y;
      if (width <= 0 || height <= 0) continue;
      const box = root.ownerDocument.createElement("span");
      box.style.cssText = `position:fixed;box-sizing:border-box;left:${x}px;top:${y}px;width:${width}px;height:${height}px;border:2px solid #d97706;border-radius:3px;background:rgba(245,158,11,.28)`;
      outline.append(box);
    }
  };
  const queuePaint = (): void => {
    if (frame || destroyed || view === null) return;
    frame = view.requestAnimationFrame(() => { frame = 0; paintActive(); });
  };
  const scrollActive = (): void => {
    const range = activeIndex === null ? undefined : ranges[activeIndex];
    if (!range || view === null) return;
    const element = range.startContainer.parentElement;
    // Reveal a lazy page first, then scroll the actual occurrence rather than
    // the whole paragraph (which may span more than one viewport).
    element?.scrollIntoView({ block: "nearest", inline: "nearest" });
    for (let parent = element?.parentElement ?? null; parent !== null; parent = parent.parentElement) {
      const css = view.getComputedStyle(parent);
      if (!/(auto|scroll)/.test(css.overflowY) || parent.scrollHeight <= parent.clientHeight) continue;
      const box = parent.getBoundingClientRect();
      const rect = range.getBoundingClientRect();
      if (rect.top < box.top || rect.bottom > box.bottom) {
        parent.scrollBy({ top: (rect.top + rect.bottom - box.top - box.bottom) / 2, behavior: "smooth" });
      }
    }
    queuePaint();
  };

  const refresh = (): void => {
    if (destroyed || registry === undefined || HighlightConstructor === undefined || queryKey === null) {
      registry?.delete(name);
      ranges = [];
      owners = [];
      activeIndex = null;
      paintActive();
      return;
    }

    attachStyle();
    const previousCount = ranges.length;
    const previousIndex = activeIndex;
    const previous = activeIndex === null ? undefined : ranges[activeIndex];
    ranges = [];
    owners = [];
    const selectorRoot: ParentNode = root;
    const blocks = Array.from(selectorRoot.querySelectorAll(BLOCK_SELECTOR));
    for (const block of blocks) {
      const { text, segments } = readBlock(block);
      if (text.length === 0 || segments.length === 0) continue;
      const key = searchKey(text);
      if (key.length === 0 || !key.includes(queryKey)) continue;
      const offsets = sourceMap(text, key);
      let at = 0;
      while ((at = key.indexOf(queryKey, at)) !== -1) {
        const first = offsets[at];
        const last = offsets[at + queryKey.length - 1];
        if (first !== undefined && last !== undefined) {
          const start = boundaryAt(segments, first.start, false);
          const end = boundaryAt(segments, last.end, true);
          if (start !== undefined && end !== undefined) {
            const range = root.ownerDocument.createRange();
            range.setStart(start.node, start.offset);
            range.setEnd(end.node, end.offset);
            ranges.push(range);
            owners.push({ block, key });
          }
        }
        at += Math.max(1, queryKey.length);
      }
    }
    if (ranges.length === 0) registry.delete(name);
    else registry.set(name, new HighlightConstructor(...ranges));
    const selected = activeSelector === null
      ? ranges.findIndex((range) => previous?.startContainer === range.startContainer && previous.startOffset === range.startOffset)
      : ranges.findIndex((range) => (range.startContainer.parentElement?.closest(activeSelector!) ?? null) !== null);
    activeIndex = ranges.length === 0 || (activeSelector !== null && selected < 0) ? null : selected >= 0 ? selected : Math.min(activeIndex ?? 0, ranges.length - 1);
    paintActive();
    if (previousCount !== ranges.length || previousIndex !== activeIndex) changed();
  };

  // Style lives outside the render target so renderer replacements do not
  // remove it. Observe text rebuilds (XLSX row windows and DOCX split wraps).
  const observer = view === null ? undefined : new view.MutationObserver(() => refresh());
  observer?.observe(root, { childList: true, characterData: true, attributes: true, attributeFilter: ["style", "class"], subtree: true });
  root.ownerDocument.addEventListener("scroll", queuePaint, true);
  view?.addEventListener("resize", queuePaint);
  const resize = typeof ResizeObserver === "undefined" ? undefined : new ResizeObserver(queuePaint);
  resize?.observe(root);

  return {
    get count() { return ranges.length; },
    get index() { return activeIndex; },
    get occurrences() { return ranges.map((range, index) => ({ range, ...owners[index]! })); },
    setQuery(query: string): void {
      if (destroyed) return;
      const folded = searchKey(query);
      if (folded.length === 0) {
        queryKey = null;
        registry?.delete(name);
        ranges = [];
        activeIndex = null;
        activeSelector = null;
        paintActive();
        changed();
        return;
      }
      if (registry === undefined || HighlightConstructor === undefined) {
        throw new AaredError(
          "unsupported-feature",
          "aared/preview: search highlighting requires browser support for the CSS Custom Highlight API",
        );
      }
      if (folded !== queryKey) { activeIndex = null; activeSelector = null; ranges = []; }
      queryKey = folded;
      refresh();
    },
    activate(index: number, scroll = false): void {
      if (destroyed) return;
      refresh();
      activeSelector = null;
      activeIndex = index >= 0 && index < ranges.length ? index : null;
      paintActive();
      if (scroll) scrollActive();
      changed();
    },
    activateWithin(selector: string): void {
      if (destroyed) return;
      activeSelector = selector;
      refresh();
    },
    destroy(): void {
      if (destroyed) return;
      destroyed = true;
      observer?.disconnect();
      resize?.disconnect();
      root.ownerDocument.removeEventListener("scroll", queuePaint, true);
      view?.removeEventListener("resize", queuePaint);
      if (frame) view?.cancelAnimationFrame(frame);
      registry?.delete(name);
      if (adoptedOn !== null && sheet !== null) adoptedOn.adoptedStyleSheets = adoptedOn.adoptedStyleSheets.filter((item) => item !== sheet);
      adoptedOn = null;
      outline.remove();
      if (root.getAttribute("data-aared-search-highlight") === name) {
        root.removeAttribute("data-aared-search-highlight");
      }
      queryKey = null;
      ranges = [];
      activeIndex = null;
    },
  };
}
