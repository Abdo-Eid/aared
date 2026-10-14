import { useEffect, useMemo, useRef, useState } from "react";
import { usePreview } from "@aared/viewer/react";
import { isAbort, messageFor } from "@aared/viewer/core";
import { defaultResolver } from "aared-fonts";
import type { BenchTheme } from "../adapters";
import type { DocMeta } from "../lib/doc";
import { ViewerFrame, useViewerFrame, type Status } from "./ViewerFrame";
import { ViewerSearch } from "./ViewerSearch";

/**
 * `?fonts=aared` passes aared-fonts' resolver, so a device without Office's
 * fonts (a phone) paints the calibrated substitutes. Off by default: the
 * bench's timings stay free of font downloads.
 */
const FONT_RESOLVER = new URLSearchParams(location.search).get("fonts") === "aared" ? defaultResolver : undefined;

/**
 * The bench's appearance: its theme, and no XLSX sheet tabs (on by default in
 * embeds), since the frame draws its own tab strip. The hook applies the
 * full appearance when its values change.
 */
function benchAppearance(theme: BenchTheme) {
  return { theme, spreadsheet: { sheetTabs: false, formulaBar: true } };
}

/**
 * The aared column, built the way a consumer app builds it: one
 * `usePreview()` hook owning detection, open timing, navigation guards,
 * and teardown. The bench adds only what its UI needs on top: the
 * open-speed `ms`, an explicit theme (never `"system"`, so the preview
 * follows the bench chrome instead of the OS preference), and DOCX layout
 * mode. Layout pages, the heading outline, and DOCX zoom fit all come
 * straight from `viewer.docx` (`@aared/viewer/preview`'s real API) — not scraped
 * from the painted DOM, which only ever reflected render output, never
 * the parsed document itself.
 */
export function AaredViewerCard({ doc, theme, onClose }: {
  doc: DocMeta;
  theme: BenchTheme;
  onClose: () => void;
}) {
  const [searchOpen, setSearchOpen] = useState(false);
  const [actionError, setActionError] = useState<string | null>(null);
  useEffect(() => setActionError(null), [doc]);
  // The bytes the bench already holds: stable per document, so the preview
  // mounts once (aared copies them for the Worker itself).
  const t0 = useMemo(() => performance.now(), [doc.bytes]);
  // Theme changes flow through the hook without reopening the document.
  const hook = usePreview({
    source: doc.bytes,
    appearance: benchAppearance(theme),
    pagination: { mode: "layout" },
    ...(FONT_RESOLVER === undefined ? {} : { fontResolver: FONT_RESOLVER }),
  });
  useEffect(() => setActionError(null), [hook.viewer, hook.index]);

  const status: Status = useMemo(() => {
    if (hook.status === "error") {
      const ms = Math.round(performance.now() - t0);
      return { state: "error", message: hook.error.message, ms };
    }
    if (hook.status !== "ready") return { state: "loading" };
    const { viewer } = hook;
    const info = viewer.info;
    const ms = Math.round(performance.now() - t0);
    if (viewer.format === "pptx") {
      return {
        state: "ready",
        handle: {
          format: "pptx",
          pageCount: info.count,
          ms,
          goToSlide: (index) => viewer.goTo(index),
        },
      };
    }
    if (viewer.format === "xlsx") {
      return {
        state: "ready",
        handle: {
          format: "xlsx",
          pageCount: info.count,
          ms,
          sheets: [...info.titles],
          renderSheet: (index) => viewer.goTo(index),
        },
      };
    }
    const docx = viewer.docx;
    // aared reports inches; the bench's pager works in CSS pixels like the other adapters.
    const pages = docx.pages.map((page) => ({ width: page.width * 96, height: page.height * 96 }));
    return {
      state: "ready",
      handle: {
        format: "docx",
        pageCount: docx.pageCount,
        pages,
        ms,
        outline: docx.headings.map((heading) => ({
          level: heading.level,
          text: heading.title,
          onSelect: () => void docx.goToHeading(heading.index).then(
            () => setActionError(null),
            (error: unknown) => { if (!isAbort(error)) setActionError(messageFor(error)); },
          ),
        })),
        paperWidth: pages.length > 0 ? Math.max(...pages.map((page) => page.width)) : 0,
      },
    };
    // Not `hook` itself: it also changes with the index and zoom, which would
    // re-time the open on every move.
  },[hook.status, hook.viewer, hook.status === "error" ? hook.error : null, t0]);

  const firedKey = useRef<string | null>(null);
  useEffect(() => {
    if (status.state === "loading") {
      firedKey.current = null;
      return;
    }
    if (status.state === "ready") {
      const key = `ready:${status.handle.format}:${status.handle.ms}`;
      if (firedKey.current === key) return;
      firedKey.current = key;
      console.info(
        `[bench] ${doc.name} via aared: ${status.handle.ms} ms, ${status.handle.pageCount} pages`,
      );
    } else {
      const key = `error:${status.ms}:${status.message}`;
      if (firedKey.current === key) return;
      firedKey.current = key;
      console.error(`[bench] ${doc.name} via aared failed:`, status.message);
    }
  }, [status, doc]);

  // Imperative actions are app-owned; built-in actions land in hook.error.
  const notice = actionError ?? (hook.status === "ready" ? hook.error?.message ?? null : null);

  const controller = useViewerFrame({
    doc,
    frameId: `${doc.key}:aared`,
    renderer: "aared",
    status,
    hostRef: hook.ref,
    viewer: hook.viewer,
    index: hook.index ?? 0,
    reportError: (message) => setActionError(message || null),
  });

  // The pager's zoom buttons drive the product's own zoom for every format
  // (see ViewerFrame's `drivesOwnZoom`), not a bench-side CSS `zoom`. The
  // unzoomed pager reading of 100% means the zoom the format opens at
  // (`resetZoom()`): for PPTX and DOCX a numeric `setZoom(1)` would instead
  // pin the *natural* (unfitted) size, reintroducing the overflow the default
  // exists to avoid in a host narrower or shorter than the content; XLSX's
  // default is 100%, like Excel. This must always call the viewer, even back
  // at 100%: once a numeric zoom has been set, the product stays there until
  // told otherwise.
  const { viewer } = hook;
  useEffect(() => {
    if (viewer === null) return;
    if (controller.zoom === 1) viewer.resetZoom();
    else viewer.setZoom(controller.zoom);
  }, [controller.zoom, viewer]);

  return (
    <div className="flex h-full min-h-0 min-w-0 flex-1 flex-col gap-2">
      {hook.status === "ready" && searchOpen ? (
        <ViewerSearch viewer={hook.viewer} index={hook.index} />
      ) : null}
      <ViewerFrame doc={doc} renderer="aared" onClose={onClose} controller={controller} notice={notice}
        search={hook.status === "ready" ? { open: searchOpen, onToggle: () => setSearchOpen((open) => !open) } : null}
      />
    </div>
  );
}
