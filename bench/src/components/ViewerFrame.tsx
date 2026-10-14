import { useEffect, useRef, useState } from "react";
import type { UsePreviewReady } from "@aared/viewer/react";
import type { Dispatch, RefObject, SetStateAction } from "react";
import {
  AlertTriangle,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Download,
  List,
  Loader2,
  PanelLeft,
  Search,
  X,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { isAbort, type OpenedDoc, type OutlineEntry } from "../adapters";
import { FORMAT_BADGE } from "../lib/format";
import {
  clamp,
  formatSize,
  type DocMeta,
  type Renderer,
} from "../lib/doc";
import {
  cn,
  FileBadge,
  IconButton,
  IconLink,
  Separator,
  StageMessage,
} from "./chrome";

/** Repeated `/1.25`/`*1.25` steps land near but not exactly on 1 (float drift); snap back so "100%" reliably means "the product's own default", not a numeric zoom that merely rounds to it. */
function snapZoom(value: number): number {
  return Math.abs(value - 1) < 0.01 ? 1 : value;
}

/** Chrome reads display data; the mounting card owns teardown. */
type FrameDocument<T = OpenedDoc> = T extends OpenedDoc ? Omit<T, "destroy"> : never;

export type Status =
  | { state: "loading" }
  | { state: "ready"; handle: FrameDocument }
  | { state: "error"; message: string; ms: number };

/**
 * Everything the viewer chrome reads: derived rail/pager data, UI state,
 * and navigation. Produced by `useViewerFrame` from a `Status`, so the
 * baseline and aared cards share every interaction below the open step.
 */
export interface ViewerController {
  status: Status;
  hostRef: RefObject<HTMLDivElement | null>;
  scrollRef: RefObject<HTMLDivElement | null>;
  currentPage: number;
  pageCount: number;
  outline: OutlineEntry[];
  thumbs: string[];
  sheets: string[];
  hasRail: boolean;
  showPager: boolean;
  horizontal: boolean;
  zoom: number;
  setZoom: Dispatch<SetStateAction<number>>;
  railOpen: boolean;
  setRailOpen: Dispatch<SetStateAction<boolean>>;
  activeSheet: number;
  setActiveSheet: (index: number) => void;
  ms: number | null;
  failed: boolean;
  goToPage: (page: number) => void;
}

/**
 * Which viewer card owns window-level keyboard paging. Compare mode mounts
 * two cards with two global listeners; without this, one keypress pages both
 * viewers. Clicking (pointerdown) inside a card claims ownership, and each
 * mount claims once so a lone card (or the last-mounted one) just works.
 */
let activeFrameId: string | null = null;

function claimViewerFrame(id: string): void {
  activeFrameId = id;
}

export function useViewerFrame(args: {
  doc: DocMeta;
  /** Unique per card (`doc.key` + renderer); owns keyboard paging while claimed. */
  frameId: string;
  /** aared drives its own real zoom for DOCX (see `AaredViewerCard`'s
   * fit-width effect) instead of the bench-computed `paperWidth` fit below,
   * which stays for baselines with no zoom API of their own. */
  renderer: Renderer;
  status: Status;
  hostRef: RefObject<HTMLDivElement | null>;
  /** Called when a sheet or page fails to show; absent when the host reports it elsewhere (aared's `notice`). */
  reportError?: (message: string, ms: number) => void;
  viewer?: UsePreviewReady["viewer"] | null;
  index?: number;
}): ViewerController {
  const { doc, frameId, renderer, status, hostRef, reportError, viewer, index = 0 } = args;
  const [scrollPage, setScrollPage] = useState(1);
  const currentPage = viewer?.format === "pptx" ? index + 1 : scrollPage;
  const [zoom, setZoom] = useState(1);
  const [railOpen, setRailOpen] = useState(true);
  const [baselineSheet, setBaselineSheet] = useState(0);
  const activeSheet = viewer?.format === "xlsx" ? index : baselineSheet;
  const setActiveSheet = (next: number): void => {
    if (viewer?.format === "xlsx") void viewer.goTo(next).then(reportViewerSuccess, reportViewerFailure);
    else setBaselineSheet(next);
  };
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const reportViewerFailure = (error: unknown): void => {
    if (isAbort(error)) return;
    reportError?.(error instanceof Error ? error.message : "Could not go there.", status.state === "ready" ? status.handle.ms : 0);
  };
  const reportViewerSuccess = (): void => {
    reportError?.("", status.state === "ready" ? status.handle.ms : 0);
  };
  const paperWidthRef = useRef(0);

  // Fresh document: pager, zoom, and sheet selection restart. Theme is
  // viewer chrome only and must never reset navigation — appearance
  // forwards in place via `setAppearance()`, keeping the current page.
  useEffect(() => {
    setScrollPage(1);
    setZoom(1);
    setBaselineSheet(0);
  }, [doc]);

  // Fit-to-width is the default zoom for paged documents, the way the
  // product opens: measured once the handle reports its paper width.
  // aared drives its own real `fit-width` instead (see `renderer` above).
  useEffect(() => {
    if (renderer === "aared") return;
    if (status.state !== "ready" || status.handle.format !== "docx") return;
    if (status.handle.paperWidth > 0) {
      paperWidthRef.current = status.handle.paperWidth;
      const stageW = scrollRef.current?.clientWidth ?? 0;
      if (stageW > 0) {
        setZoom(clamp(stageW / status.handle.paperWidth, 0.25, 2));
      }
    } else {
      paperWidthRef.current = 0;
    }
  }, [status]);

  // XLSX tabs re-render through the handle; the open timing is untouched.
  useEffect(() => {
    if (!viewer && status.state === "ready" && status.handle.format === "xlsx") {
      void Promise.resolve(status.handle.renderSheet(activeSheet)).catch((error: unknown) => {
        if (isAbort(error)) return;
        const message = error instanceof Error ? error.message : "Could not show this sheet.";
        console.error(`[bench] ${doc.name} sheet ${activeSheet + 1} failed:`, message);
        reportError?.(message, status.handle.ms);
      });
    }
  }, [status, activeSheet, doc.name, reportError, viewer]);

  // The pager follows the scroll for paged documents.
  useEffect(() => {
    if (status.state !== "ready") return;
    if (
      status.handle.format !== "pdf" &&
      status.handle.format !== "docx"
    ) {
      return;
    }
    const root = scrollRef.current;
    if (!root) return;
    const observer = new IntersectionObserver(
      (entries) => {
        let best: { page: number; ratio: number } | null = null;
        for (const entry of entries) {
          const page = Number(
            (entry.target as HTMLElement).dataset.page ?? 0,
          );
          if (
            entry.isIntersecting &&
            page >= 1 &&
            (best === null || entry.intersectionRatio > best.ratio)
          ) {
            best = { page, ratio: entry.intersectionRatio };
          }
        }
        if (best) setScrollPage(best.page);
      },
      { root, threshold: [0, 0.25, 0.5, 0.75, 1] },
    );
    root.querySelectorAll("[data-page]").forEach((el) => observer.observe(el));
    return () => observer.disconnect();
  }, [status]);

  // Recalculate fit-to-width zoom when the rail toggles, so content stays
  // sized to the scroll area rather than overflowing. A no-op for aared:
  // `paperWidthRef` never gets set above for it.
  useEffect(() => {
    if (status.state !== "ready") return;
    const pw = paperWidthRef.current;
    if (pw <= 0) return;
    const stageW = scrollRef.current?.clientWidth ?? 0;
    if (stageW > 0) {
      setZoom(clamp(stageW / pw, 0.25, 2));
    }
  }, [railOpen, status]);

  const pageCount =
    status.state === "ready" ? status.handle.pageCount : 1;
  const outline =
    status.state === "ready" && status.handle.format === "docx"
      ? status.handle.outline
      : [];
  const thumbs =
    status.state === "ready" && status.handle.format === "pdf"
      ? status.handle.thumbs
      : [];
  const sheets =
    status.state === "ready" && status.handle.format === "xlsx"
      ? status.handle.sheets
      : [];
  const hasRail = thumbs.length > 0 || outline.length > 0;
  const showPager =
    status.state === "ready" &&
    status.handle.format !== "xlsx" &&
    pageCount > 1;
  const horizontal =
    status.state === "ready" && status.handle.format === "pptx";

  const goToPage = (page: number): void => {
    const next = clamp(page, 1, Math.max(pageCount, 1));
    if (viewer?.format === "pptx") {
      void viewer.goTo(next - 1).then(reportViewerSuccess, reportViewerFailure);
      return;
    }
    if (viewer?.format === "docx") {
      void viewer.docx.goToPage(next - 1).then(reportViewerSuccess, reportViewerFailure);
      return;
    }
    setScrollPage(next);
    if (status.state !== "ready") return;
    const handle = status.handle;
    if (handle.format === "pptx") {
      void Promise.resolve(handle.goToSlide(next - 1)).catch((error: unknown) => {
        const message = error instanceof Error ? error.message : "Could not show this slide.";
        console.error(`[bench] ${doc.name} slide ${next} failed:`, message);
        reportError?.(message, handle.ms);
      });
    } else if (handle.format === "pdf") {
      handle.scrollToPage(next);
    } else if (handle.format === "docx") {
      // Aared renders `section[data-aared-element="page"]` pages; the
      // baseline renders plain `section.docx` pages — prefer paged sections.
      const paged = hostRef.current?.querySelectorAll('section[data-aared-element="page"]');
      const sections = paged?.length
        ? paged
        : hostRef.current?.querySelectorAll("section");
      const el = sections?.[next - 1];
      if (el) el.scrollIntoView({ block: "start" });
    }
  };

  // Claiming on mount keeps single-card mode working with no clicks, and
  // in compare mode the last-mounted card (aared) starts owning the keys.
  useEffect(() => {
    claimViewerFrame(frameId);
  }, [frameId]);

  // Keyboard paging: `<`/`>` step any paged document, and the horizontal PPTX
  // pager also takes the left/right arrows. Ignored while typing, and ignored
  // when another card claimed the keys (compare mode: click a card to steer it).
  useEffect(() => {
    if (status.state !== "ready") return;
    const format = status.handle.format;
    if (format !== "pptx" && format !== "docx" && format !== "pdf") return;
    const horizontalKeys = format === "pptx";
    const onKeyDown = (event: KeyboardEvent): void => {
      if (activeFrameId !== frameId) return;
      if (event.defaultPrevented || event.metaKey || event.ctrlKey || event.altKey) return;
      const target = event.target as HTMLElement | null;
      if (target?.closest?.("input, textarea, select, [contenteditable='true']")) return;
      let delta = 0;
      if (event.key === "<") delta = -1;
      else if (event.key === ">") delta = 1;
      else if (horizontalKeys && event.key === "ArrowLeft") delta = -1;
      else if (horizontalKeys && event.key === "ArrowRight") delta = 1;
      if (delta === 0) return;
      event.preventDefault();
      goToPage(currentPage + delta);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [status, currentPage, frameId, viewer]);

  const ms =
    status.state === "error"
      ? status.ms
      : status.state === "ready"
        ? status.handle.ms
        : null;
  const failed = status.state === "error";

  return {
    status,
    hostRef,
    scrollRef,
    currentPage,
    pageCount,
    outline,
    thumbs,
    sheets,
    hasRail,
    showPager,
    horizontal,
    zoom,
    setZoom,
    railOpen,
    setRailOpen,
    activeSheet,
    setActiveSheet,
    ms,
    failed,
    goToPage,
  };
}

/**
 * One rendered document in the product's viewer chrome: badge-and-title
 * header, rail beside a continuous stage, footer with pager and zoom. The
 * renderer behind it is interchangeable — baseline library or aared — which
 * is the whole point of the bench.
 */
export function ViewerFrame({ doc, renderer, onClose, controller, notice, search }: {
  doc: DocMeta;
  renderer: Renderer;
  onClose: () => void;
  controller: ViewerController;
  /** A failure that leaves the view usable (a page that could not be shown): a bar above the stage, never the error screen. */
  notice?: string | null;
  search?: { open: boolean; onToggle: () => void } | null;
}) {
  const {
    status,
    hostRef,
    scrollRef,
    currentPage,
    pageCount,
    outline,
    thumbs,
    sheets,
    hasRail,
    showPager,
    horizontal,
    zoom,
    setZoom,
    railOpen,
    setRailOpen,
    activeSheet,
    setActiveSheet,
    ms,
    failed,
    goToPage,
  } = controller;
  const isAaredPptx = renderer === "aared" && status.state === "ready" && status.handle.format === "pptx";
  const isAaredDocx = renderer === "aared" && status.state === "ready" && status.handle.format === "docx";
  const isAaredXlsx = renderer === "aared" && status.state === "ready" && status.handle.format === "xlsx";
  const drivesOwnZoom = isAaredPptx || isAaredDocx || isAaredXlsx;
  const fitsStage = isAaredPptx && zoom === 1;
  const PreviousIcon = horizontal ? ChevronLeft : ChevronUp;
  const NextIcon = horizontal ? ChevronRight : ChevronDown;

  return (
    <section
      aria-label={`${doc.name} viewer`}
      // Clicking a card steers window-level keyboard paging to it (compare
      // mode mounts two global listeners; the claim keeps one keypress on one card).
      onPointerDownCapture={() => claimViewerFrame(`${doc.key}:${renderer}`)}
      className="flex h-full min-h-0 min-w-0 flex-1 flex-col overflow-hidden rounded-2xl border border-border bg-card text-foreground shadow-[0_1rem_2.5rem_-1.125rem_rgba(20,20,30,0.28)]"
    >
      <div className="flex shrink-0 items-center gap-2.5 border-b border-border/60 px-[0.9375rem] pb-[0.8125rem] pt-[0.875rem]">
        <div className="flex min-w-0 flex-1 flex-col gap-1.5">
          {renderer === "aared" ? (
            <p className="font-mono text-[0.625rem] font-bold uppercase tracking-[0.083rem] text-accent-blue">
              aared · usePreview()
            </p>
          ) : null}
          <div className="flex min-w-0 items-center gap-[0.5625rem]">
            <FileBadge
              className={FORMAT_BADGE[doc.format]}
              label={doc.format}
            />
            <h2
              className="min-w-0 flex-1 truncate text-[1.03125rem] font-extrabold leading-[1.2375rem] tracking-[-0.0156rem]"
              title={doc.name}
            >
              {doc.name}
            </h2>
          </div>
        </div>
        <div className="flex shrink-0 items-center gap-0.5">
          {search ? (
            <IconButton label="Search document" pressed={search.open} aria-expanded={search.open} onClick={search.onToggle}>
              <Search className="size-[1.0625rem]" />
            </IconButton>
          ) : null}
          <IconLink
            label={`Download ${doc.name}`}
            href={doc.url}
            download={doc.name}
          >
            <Download className="size-[1.0625rem]" />
          </IconLink>
          <IconButton label="Close document" onClick={onClose}>
            <X className="size-[1.0625rem]" />
          </IconButton>
        </div>
      </div>

      {sheets.length > 1 ? (
        <div
          role="tablist"
          aria-label="Sheets"
          className="flex shrink-0 items-center gap-1 overflow-x-auto border-b border-border/60 bg-card px-3 py-1.5"
        >
          {sheets.map((name, index) => (
            <button
              key={`${name}-${index}`}
              role="tab"
              aria-selected={index === activeSheet}
              type="button"
              onClick={() => setActiveSheet(index)}
              className={cn(
                "focus-ring shrink-0 cursor-pointer rounded-md px-2.5 py-1 text-xs font-semibold transition-colors",
                index === activeSheet
                  ? "bg-accent-blue/10 text-accent-blue"
                  : "text-muted-foreground hover:bg-muted hover:text-foreground",
              )}
            >
              {name}
            </button>
          ))}
        </div>
      ) : null}

      {notice ? (
        <div
          role="status"
          className="flex shrink-0 items-center gap-2 border-b border-border/60 bg-card px-[0.9375rem] py-1.5 text-xs text-muted-foreground"
        >
          <AlertTriangle className="size-3.5 shrink-0" aria-hidden />
          <p className="min-w-0 truncate" title={notice}>{notice}</p>
        </div>
      ) : null}

      <div className="flex min-h-0 flex-1">
        {hasRail && railOpen ? (
          <div className="w-36 shrink-0 overflow-y-auto border-r border-border/60 bg-card p-2">
            <p className="mb-2 px-1 font-mono text-[0.625rem] uppercase tracking-widest text-muted-foreground">
              {thumbs.length > 0 ? "Pages" : "Contents"}
            </p>
            <div className="grid gap-2">
              {thumbs.map((src, index) => (
                <button
                  key={index}
                  type="button"
                  onClick={() => goToPage(index + 1)}
                  aria-label={`Go to page ${index + 1}`}
                  aria-current={index + 1 === currentPage}
                  className={cn(
                    "focus-ring block w-full cursor-pointer overflow-hidden rounded border bg-white transition-opacity",
                    index + 1 === currentPage
                      ? "border-accent-blue shadow-[0_0_0.375rem_rgba(10,130,242,0.25)]"
                      : "border-black/10 opacity-60 hover:opacity-100",
                  )}
                >
                  <img
                    alt={`Page ${index + 1} thumbnail`}
                    className="block w-full"
                    loading="lazy"
                    src={src}
                  />
                </button>
              ))}
              {outline.map((entry, index) => (
                <button
                  key={`${entry.text}-${index}`}
                  type="button"
                  onClick={() =>
                    entry.onSelect ? entry.onSelect() : entry.element?.scrollIntoView({ block: "start" })
                  }
                  className="focus-ring block w-full cursor-pointer rounded px-1.5 py-1 text-left text-xs leading-snug text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
                  style={{ paddingLeft: 6 + (entry.level - 1) * 10 }}
                  title={entry.text}
                >
                  <span className="line-clamp-2">{entry.text}</span>
                </button>
              ))}
            </div>
          </div>
        ) : null}

        <div
          ref={scrollRef}
          className="relative min-w-0 flex-1 overflow-auto bg-background"
        >
          {/* aared drives its own zoom in every format (real pixel resize via
              `setZoom`, wired in `AaredViewerCard`; XLSX scales its grid and
              keeps its own scroller and fill) instead of the bench's
              CSS `zoom` trick: CSS `zoom` on a `width: 100%` box never grows
              past its container, so the browser reports no overflow and a
              zoomed-in slide or page clips with nothing to scroll to. At the
              unzoomed fit, PPTX's own `fit-page` default sizes the slide to
              the host's height, so the host fills the stage — a height to
              fit to, as any embedder has to give; DOCX's own `fit-width`
              sizes pages to the host's width instead (a page keeps scrolling
              vertically past one screenful, unlike a slide). Zoomed in
              either way, the host grows with its real content size and the
              stage scrolls. Baselines keep the CSS `zoom` trick: they have no
              real zoom API of their own to drive instead. */}
          <div className={fitsStage ? "box-border h-full p-4 md:p-6" : "p-4 md:p-6"}>
            <div
              style={drivesOwnZoom ? undefined : { zoom }}
              // aared's own box is `w-full`: `fit-width`/`fit-page` measure
              // and size themselves against the real container edge, so the
              // box needs to actually reach that edge. A baseline library
              // paints at its own natural (or CSS-`zoom`ed) size with no such
              // measurement, so `w-fit max-w-full` avoids stretching it to
              // fill space it never asked to fill.
              className={renderer === "aared" ? `mx-auto w-full${fitsStage ? " h-full" : ""}` : "mx-auto w-fit max-w-full"}
            >
              {status.state !== "error" ? (
                <div ref={hostRef} key={`${doc.key}-${renderer}`} className={fitsStage ? "h-full" : undefined} />
              ) : null}
              {status.state === "loading" ? (
                <div className="pointer-events-none absolute inset-0 flex">
                  <StageMessage
                    className="bg-transparent"
                    icon={
                      <Loader2 className="size-4 animate-spin motion-reduce:animate-none" />
                    }
                  >
                    Opening {doc.name}…
                  </StageMessage>
                </div>
              ) : null}
              {status.state === "error" ? (
                <StageMessage
                  icon={<AlertTriangle className="size-4 text-warning" />}
                >
                  {status.message} The file itself is one click away in the
                  header.
                </StageMessage>
              ) : null}
            </div>
          </div>
        </div>
      </div>

      <div className="flex shrink-0 items-center justify-between border-t border-border/60 bg-card px-2.5 py-2">
        <div className="flex items-center gap-[0.1875rem]">
          {hasRail ? (
            outline.length > 0 ? (
              <button
                type="button"
                onClick={() => setRailOpen((value) => !value)}
                aria-pressed={railOpen}
                className={cn(
                  "focus-ring flex h-[1.875rem] cursor-pointer items-center gap-[0.4375rem] rounded-lg px-1 transition-colors hover:bg-muted",
                  railOpen ? "text-foreground" : "text-muted-foreground",
                )}
              >
                <List aria-hidden="true" className="size-4" />
                <span className="text-[0.78125rem] font-semibold tracking-[-0.004em]">
                  Contents
                </span>
              </button>
            ) : (
              <IconButton
                label={
                  railOpen ? "Hide page thumbnails" : "Show page thumbnails"
                }
                onClick={() => setRailOpen((value) => !value)}
                pressed={railOpen}
                size="footer"
              >
                <PanelLeft className="size-4" />
              </IconButton>
            )
          ) : null}
          {showPager ? (
            <>
              <Separator />
              <IconButton
                label={horizontal ? "Previous slide" : "Previous page"}
                onClick={() => goToPage(currentPage - 1)}
                disabled={currentPage <= 1}
                size="footer"
              >
                <PreviousIcon className="size-4" />
              </IconButton>
              <p
                aria-live="polite"
                className="min-w-[2.625rem] text-center font-mono text-[0.6875rem] text-muted-foreground"
              >
                <span className="sr-only">
                  {horizontal ? "Slide " : "Page "}
                </span>
                {currentPage}
                <span className="text-neutral-soft"> / {pageCount}</span>
              </p>
              <IconButton
                label={horizontal ? "Next slide" : "Next page"}
                onClick={() => goToPage(currentPage + 1)}
                disabled={currentPage >= pageCount}
                size="footer"
              >
                <NextIcon className="size-4" />
              </IconButton>
            </>
          ) : null}
          {status.state === "ready" &&
          status.handle.format === "xlsx" ? (
            <p className="px-1 font-mono text-[0.6875rem] text-muted-foreground">
              {status.handle.sheets.length}{" "}
              {status.handle.sheets.length === 1 ? "sheet" : "sheets"}
            </p>
          ) : null}
        </div>

        <div className="flex items-center gap-[0.1875rem]">
          <IconButton
            label="Zoom out"
            onClick={() => setZoom((z) => snapZoom(clamp(z / 1.25, 0.25, 3)))}
            disabled={zoom <= 0.25}
            size="footer"
          >
            <ZoomOut className="size-4" />
          </IconButton>
          <p className="min-w-[2.75rem] text-center font-mono text-[0.6875rem] text-muted-foreground">
            <span className="sr-only">Zoom </span>
            {Math.round(zoom * 100)}%
          </p>
          <IconButton
            label="Zoom in"
            onClick={() => setZoom((z) => snapZoom(clamp(z * 1.25, 0.25, 3)))}
            disabled={zoom >= 3}
            size="footer"
          >
            <ZoomIn className="size-4" />
          </IconButton>
          <Separator />
          <span
            title={`${doc.name} · ${formatSize(doc.sizeBytes)} · ${renderer}`}
            className={cn(
              "rounded-md bg-muted px-2 py-1 font-mono text-[11px] whitespace-nowrap",
              failed ? "text-warning" : "text-muted-foreground",
            )}
          >
            {ms === null ? "…" : `${ms.toLocaleString("en-US")} ms`}
          </span>
        </div>
      </div>
    </section>
  );
}
