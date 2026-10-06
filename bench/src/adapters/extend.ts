import type { Root } from "react-dom/client";
import { abortError, type OpenedDoc, type OutlineEntry } from "./types";
import { maxSectionWidthPx } from "./docx";

/**
 * Extend's viewer family (@extend-ai/react-pptx, react-docx, react-xlsx):
 * three React components over Rust/WASM parsers that run in their own
 * Workers — the one other three-format family with a single engine per
 * format, so the closest structural match to aared. Each column mounts the
 * component into its own React root and resolves at the library's own
 * first-paint signal, not at parse completion.
 */

/** A React root on `host`, plus the element factory, loaded lazily with the library. */
async function reactRoot(host: HTMLElement): Promise<{
  root: Root;
  createElement: typeof import("react").createElement;
  hooks: Pick<typeof import("react"), "useEffect">;
}> {
  const [react, { createRoot }] = await Promise.all([import("react"), import("react-dom/client")]);
  host.replaceChildren();
  return { root: createRoot(host), createElement: react.createElement, hooks: react };
}

/** Rejects on abort, otherwise settles with `run`; unmounts the root on either failure. */
function untilPainted<T>(
  root: Root,
  signal: AbortSignal,
  run: (resolve: (value: T) => void, reject: (error: unknown) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const onAbort = (): void => reject(abortError());
    signal.addEventListener("abort", onAbort, { once: true });
    run(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      },
    );
  }).catch((error: unknown) => {
    root.unmount();
    throw error;
  });
}

/** The viewport height the bench column gives a self-scrolling viewer. */
function viewportHeight(host: HTMLElement): number {
  return host.parentElement?.clientHeight || 640;
}

/**
 * PPTX: `ReactPptxViewer` in single-slide mode with its own chrome off, so
 * the bench pager steers it through the controller. Ready = controller
 * handed over and the first slide's DOM/SVG rendered.
 */
export async function openExtendPptx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const [{ ReactPptxViewer }] = await Promise.all([
    import("@extend-ai/react-pptx"),
    import("@extend-ai/react-pptx/styles.css"),
  ]);
  if (signal.aborted) throw abortError();
  const { root, createElement } = await reactRoot(host);

  type Controller = import("@extend-ai/react-pptx").PptxViewerController;
  const controller = await untilPainted<Controller>(root, signal, (resolve, reject) => {
    let ready: Controller | null = null;
    let painted = false;
    const settle = (): void => {
      if (ready !== null && painted) resolve(ready);
    };
    root.render(
      createElement(ReactPptxViewer, {
        source: bytes.slice(),
        mode: "slide",
        height: viewportHeight(host),
        showToolbar: false,
        showThumbnails: false,
        showNotes: false,
        onReady: (value) => {
          ready = value;
          settle();
        },
        onSlideRendered: () => {
          painted = true;
          settle();
        },
        onError: reject,
      }),
    );
  });

  const slideCount = controller.getDocument()?.slides.length ?? 0;
  if (slideCount <= 0) {
    root.unmount();
    throw new Error("This presentation has no slides to show.");
  }
  const ms = Math.round(performance.now() - t0);
  return {
    format: "pptx",
    pageCount: slideCount,
    ms,
    goToSlide: (index: number) => {
      void controller.goToSlide(index).catch(() => {});
    },
    destroy: () => {
      root.unmount();
      host.replaceChildren();
    },
  };
}

/** The laid-out page boxes `ReactDocxViewer` paints, in document order. */
const DOCX_PAGE_SELECTOR = "[data-page-index]";

/**
 * Headings from the parsed model (`style.headingLevel`): the read-only view
 * marks no headings in its DOM. Each entry scrolls to the first painted
 * element, in document order, whose own text is exactly the heading's.
 */
function docxOutline(
  model: import("@extend-ai/react-docx").DocModel,
  host: HTMLElement,
): OutlineEntry[] {
  const outline: OutlineEntry[] = [];
  const candidates = Array.from(host.querySelectorAll<HTMLElement>("div, p, span"));
  let from = 0;
  for (const node of model.nodes) {
    const level = node.type === "paragraph" ? node.style?.headingLevel : undefined;
    if (node.type !== "paragraph" || level === undefined) continue;
    const text = node.children
      .map((child) => (child.type === "text" ? child.text : ""))
      .join("")
      .trim();
    if (!text) continue;
    const at = candidates.findIndex((el, index) => index >= from && el.textContent?.trim() === text);
    if (at >= 0) from = at + 1;
    outline.push({ level, text, ...(at >= 0 ? { element: candidates[at] } : {}) });
  }
  return outline;
}

/**
 * DOCX: parse with `parseDocxForViewer` (the WASM import Worker), then mount
 * `ReactDocxViewer` on the parsed document. The component has no paint
 * callback, so ready = its first laid-out page box is in the DOM. Pages are
 * read back from the DOM it painted; headings come from its model.
 */
export async function openExtendDocx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { ReactDocxViewer, parseDocxForViewer } = await import("@extend-ai/react-docx");
  if (signal.aborted) throw abortError();
  const sectionWidth = maxSectionWidthPx(bytes);
  const parsed = await parseDocxForViewer(bytes.slice());
  if (signal.aborted) {
    parsed.dispose();
    throw abortError();
  }
  const { root, createElement } = await reactRoot(host);

  try {
    await untilPainted<void>(root, signal, (resolve) => {
      const observer = new MutationObserver(() => {
        if (host.querySelector(DOCX_PAGE_SELECTOR) === null) return;
        observer.disconnect();
        resolve();
      });
      observer.observe(host, { childList: true, subtree: true });
      signal.addEventListener("abort", () => observer.disconnect(), { once: true });
      root.render(createElement(ReactDocxViewer, { document: parsed, defaultZoom: 100 }));
    });
  } catch (error) {
    parsed.dispose();
    throw error;
  }

  const pages = Array.from(host.querySelectorAll<HTMLElement>(DOCX_PAGE_SELECTOR));
  const outline = docxOutline(parsed.model, host);

  const ms = Math.round(performance.now() - t0);
  return {
    format: "docx",
    pageCount: Math.max(pages.length, 1),
    pages: pages.map((page) => ({ width: page.offsetWidth, height: page.offsetHeight })),
    ms,
    outline,
    paperWidth: Math.max(pages[0]?.offsetWidth ?? 0, await sectionWidth),
    destroy: () => {
      root.unmount();
      parsed.dispose();
      host.replaceChildren();
    },
  };
}

/**
 * XLSX: `XlsxViewer` driven by a `useXlsxViewerController` the adapter
 * holds, read-only, with the library's defaults otherwise (its canvas grid
 * renderer included — that is what a consumer gets). Ready = the workbook's
 * tabs are known and loading has finished; sheets switch through the
 * controller so the bench's sheet pager and the viewer's own tabs agree.
 */
export async function openExtendXlsx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { XlsxViewer, useXlsxViewerController } = await import("@extend-ai/react-xlsx");
  if (signal.aborted) throw abortError();
  const { root, createElement, hooks } = await reactRoot(host);

  type Controller = import("@extend-ai/react-xlsx").XlsxViewerController;
  let latest: Controller | null = null;
  const file = bytes.slice().buffer as ArrayBuffer;
  const height = viewportHeight(host);

  const controller = await untilPainted<Controller>(root, signal, (resolve, reject) => {
    function Workbook(): ReturnType<typeof createElement> {
      const current = useXlsxViewerController({ file, readOnly: true });
      latest = current;
      hooks.useEffect(() => {
        if (current.error !== null) reject(current.error);
        else if (!current.isLoading && current.tabs.length > 0) resolve(current);
      }, [current.error, current.isLoading, current.tabs.length]);
      return createElement(XlsxViewer, { controller: current, height });
    }
    root.render(createElement(Workbook));
  });

  const sheets = controller.tabs.map((tab) => tab.name);
  const ms = Math.round(performance.now() - t0);
  return {
    format: "xlsx",
    pageCount: sheets.length,
    ms,
    sheets,
    renderSheet: (index: number) => {
      (latest ?? controller).setActiveTabIndex(index);
    },
    destroy: () => {
      root.unmount();
      host.replaceChildren();
    },
  };
}
