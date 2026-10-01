import { useCallback, useEffect, useMemo, useRef } from "react";
import { usePreview } from "aared/react";
import { defaultResolver } from "aared-fonts";
import type { BenchTheme } from "../adapters";
import type { DocMeta } from "../lib/doc";
import { ViewerFrame, useViewerFrame, type Status } from "./ViewerFrame";

/**
 * `?fonts=aared` passes aared-fonts' resolver, so a device without Office's
 * fonts (a phone) paints the calibrated substitutes. Off by default: the
 * bench's timings stay free of font downloads. Module-level, so the hook
 * sees one stable resolver.
 */
const FONT_RESOLVER = new URLSearchParams(location.search).get("fonts") === "aared" ? defaultResolver : undefined;

/** What the pager's "100%" means per format: the zoom each format opens at. */
const PAGER_DEFAULT_ZOOM = { pptx: "fit-page", docx: "fit-width", xlsx: 1 } as const;

/**
 * The aared column, built the way a consumer app builds it: one
 * `usePreview()` hook owning detection, open timing, navigation guards,
 * and teardown. The bench adds only what its UI needs on top: the
 * open-speed `ms`, an explicit theme (never `"system"`, so the preview
 * follows the bench chrome instead of the OS preference), and DOCX layout
 * mode. Layout pages, the heading outline, and DOCX zoom fit all come
 * straight from `viewer.docx` (`aared/preview`'s real API) — not scraped
 * from the painted DOM, which only ever reflected render output, never
 * the parsed document itself.
 */
export function AaredViewerCard({ doc, theme, onClose }: {
  doc: DocMeta;
  theme: BenchTheme;
  onClose: () => void;
}) {
  // Stable byte identity: a fresh ArrayBuffer per render would remount the
  // preview in a loop, so the copy is memoized (the hook documents this).
  const buffer = useMemo(() => doc.bytes.slice().buffer as ArrayBuffer, [doc.bytes]);
  const t0 = useMemo(() => performance.now(), [buffer]);
  const hook = usePreview({
    source: buffer,
    appearance: { theme },
    pagination: { mode: "layout" },
    ...(FONT_RESOLVER === undefined ? {} : { fontResolver: FONT_RESOLVER }),
  });

  const status: Status = useMemo(() => {
    if (hook.status === "loading" || hook.status === "idle") return { state: "loading" };
    if (hook.status === "error") {
      const ms = Math.round(performance.now() - t0);
      const message = hook.error instanceof Error ? hook.error.message : "Could not open this file.";
      return { state: "error", message, ms };
    }
    const viewer = hook.viewer;
    const info = hook.info;
    if (viewer === null || info === null) return { state: "loading" };
    const ms = Math.round(performance.now() - t0);
    // Teardown is hook-owned; `destroy` below only satisfies the shared
    // chrome's handle union and is never called on this path.
    if (viewer.format === "pptx") {
      return {
        state: "ready",
        handle: {
          format: "pptx",
          pageCount: info.count,
          ms,
          goToSlide: (index) => viewer.goTo(index),
          destroy: () => viewer.destroy(),
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
          destroy: () => viewer.destroy(),
        },
      };
    }
    const docx = viewer.docx;
    if (docx === null) return { state: "loading" };
    const pages = docx.pages.map((page) => ({ width: page.width, height: page.height }));
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
          // A failure fires the viewer's `error` (the notice); the rejection itself needs no handling here.
          onSelect: () => void docx.goToHeading(heading.index).catch(() => {}),
        })),
        paperWidth: pages.length > 0 ? Math.max(...pages.map((page) => page.width)) : 0,
        destroy: () => viewer.destroy(),
      },
    };
  }, [hook.status, hook.viewer, hook.info, hook.error, t0]);

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

  // A slide or sheet that cannot be shown: the frame logs the rejected
  // `goTo()` and reports it here, but the handle's `error` event has already
  // put it in `hook.navigationError` (the notice below), and the unit on screen
  // stays. So this reports nothing more, and never swaps in the error screen.
  const reportError = useCallback(() => {}, []);
  const navigationError = hook.navigationError;
  const notice = navigationError === null
    ? null
    : `Could not go there: ${navigationError instanceof Error ? navigationError.message : String(navigationError)}`;

  const controller = useViewerFrame({
    doc,
    theme,
    frameId: `${doc.key}:aared`,
    renderer: "aared",
    status,
    hostRef: hook.ref,
    reportError,
  });

  // The pager's zoom buttons drive the product's own zoom for every format
  // (see ViewerFrame's `drivesOwnZoom`), not a bench-side CSS `zoom`. The
  // unzoomed pager reading of 100% asks for the format's own default
  // explicitly — for PPTX and DOCX a numeric `setZoom(1)` would instead pin
  // the *natural* (unfitted) size, reintroducing the overflow the default
  // exists to avoid in a host narrower or shorter than the content; XLSX's
  // default is 100%, like Excel. This must always call `setZoom`, even back
  // at 100%: once a numeric zoom has been set, the product stays there until
  // told otherwise.
  useEffect(() => {
    const format = hook.viewer?.format;
    if (format === undefined) return;
    hook.setZoom(controller.zoom === 1 ? PAGER_DEFAULT_ZOOM[format] : controller.zoom);
  }, [controller.zoom, hook]);

  // XLSX sheet tabs: `hook.index` already mirrors the real active sheet
  // (forwarded from `sheetchange`), so it is the source of truth — pushing
  // it into the bench's own `activeSheet` state (rather than letting a tab
  // click own that state outright) keeps the two from drifting if the
  // product's index ever changes some other way.
  useEffect(() => {
    if (hook.viewer?.format !== "xlsx") return;
    controller.setActiveSheet(hook.index);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hook.index, hook.viewer, controller.setActiveSheet]);

  return <ViewerFrame doc={doc} renderer="aared" onClose={onClose} controller={controller} notice={notice} />;
}
