/**
 * `aared/react`: React hooks over `aared/preview`.
 *
 * `usePreview()` mounts `preview(source, target)` inside an effect and owns
 * the parts every embed otherwise rewrites: the `AbortController`, the
 * loading/ready/error state, `index`/`zoom` mirroring from handle events,
 * and full teardown (`destroy()`) on unmount or source change. It is the
 * same open → render → navigate → destroy boundary as `aared/preview` —
 * search, range reads, cell navigation, and media stay on the headless
 * format entrypoints.
 *
 * `usePptx()`, `useXlsx()`, and `useDocx()` are the same hook with `type`
 * asserted, so detection is skipped and only that format chunk is fetched
 * at runtime (bundlers still emit all three) — plus a narrowed `info` type.
 * They add no lifecycle of their own.
 *
 * Import-time safe: importing this module touches no DOM or `Worker`
 * globals (SSR-safe imports) — everything runs inside effects. It lives
 * behind its own subpath with `react` as an optional peer, so vanilla
 * consumers never pay for it.
 */

import { useCallback, useEffect, useRef, useState } from "react";
import type { RefObject } from "react";
import { isAbort } from "../core/index.js";
import type { Appearance, DocumentFormat, DocumentInfo, Source, ZoomValue } from "../core/index.js";
import type { DocxInfo } from "../docx/index.js";
import type { PptxInfo } from "../pptx/index.js";
import type { XlsxInfo } from "../xlsx/index.js";
import { preview } from "../preview/index.js";
import type { PreviewHandle, PreviewOptions } from "../preview/index.js";

/**
 * Options for `usePreview()`: `PreviewOptions` plus the mounted source.
 * `signal` is hook-owned (one `AbortController` per mount) and therefore
 * omitted — pass nothing abort-related. `worker` is read at mount: a new
 * value (or a new inline factory identity) applies at the next mount
 * (`source`/`type`/`index`/`pagination` change, or `retry()`), never
 * remounting on its own. `zoom` is the starting zoom; change it after
 * mount with `setZoom()`.
 */
export interface UsePreviewOptions extends Omit<PreviewOptions, "signal"> {
  /** Previewed document. Must be referentially stable: a new identity
   * re-runs the mount, so memoize `ArrayBuffer`s (don't build one inline
   * per render). `File`, `Blob`, and URL strings are naturally stable. */
  source: Source;
}

/** Lifecycle phase of a mounted preview. */
export type UsePreviewPhase = "idle" | "loading" | "ready" | "error";

/** Controller returned by `usePreview()` and the per-format hooks. */
export interface UsePreviewResult {
  /** Attach to the host element: `<div ref={ref} />`. The hook owns what
   * is painted inside while mounted and clears it on teardown. */
  ref: RefObject<HTMLDivElement | null>;
  status: UsePreviewPhase;
  /** The mounted handle once `status` is `"ready"`; otherwise `null`. Per-format
   * controls (`viewer.xlsx`, `viewer.docx`) hang off it. */
  viewer: PreviewHandle | null;
  /** The detected (or asserted) format once mounted; otherwise `null`. */
  format: DocumentFormat | null;
  /** Document metadata once mounted; otherwise `null`. */
  info: DocumentInfo | null;
  /** Currently rendered unit index, mirrored from `indexchange`. */
  index: number;
  /** Active zoom, mirrored from `zoomchange`. */
  zoom: ZoomValue;
  /** Rejection once `status` is `"error"`; aborted mounts never land here. */
  error: unknown;
  /** Render another unit; resolves when painted. Silent no-op before
   * anything is mounted. Out-of-range indexes reject with `not-found`. */
  goTo: (index: number) => Promise<void>;
  /** Change zoom; silent no-op before anything is mounted. */
  setZoom: (zoom: ZoomValue) => void;
  /** Replace the appearance in place without remounting; silent no-op before mount. */
  setAppearance: (appearance: Appearance) => void;
  /** Re-run the current mount with the latest options. */
  retry: () => void;
}

/** Options for `usePptx()`: `UsePreviewOptions` without `type`/`pagination`. */
export interface UsePptxOptions extends Omit<UsePreviewOptions, "type" | "pagination"> {}

/** Options for `useXlsx()`: `UsePreviewOptions` without `type`/`pagination`. */
export interface UseXlsxOptions extends Omit<UsePreviewOptions, "type" | "pagination"> {}

/** Options for `useDocx()`: `UsePreviewOptions` without `type`. */
export interface UseDocxOptions extends Omit<UsePreviewOptions, "type"> {}

function isCancellation(error: unknown, signal: AbortSignal): boolean {
  return signal.aborted || isAbort(error);
}

export function usePreview({
  source,
  type,
  index: initialIndex = 0,
  appearance,
  zoom: initialZoom,
  fontResolver,
  mediaResolver,
  worker,
  pagination,
}: UsePreviewOptions): UsePreviewResult {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const viewerRef = useRef<PreviewHandle | null>(null);
  const [status, setStatus] = useState<UsePreviewPhase>("idle");
  const [viewer, setViewer] = useState<PreviewHandle | null>(null);
  const [format, setFormat] = useState<DocumentFormat | null>(null);
  const [info, setInfo] = useState<DocumentInfo | null>(null);
  const [unitIndex, setUnitIndex] = useState<number>(initialIndex);
  const [zoom, setZoomValue] = useState<ZoomValue>(initialZoom ?? "fit-width");
  const [error, setError] = useState<unknown>(null);
  const [nonce, setNonce] = useState(0);

  // Plain-data mount options are keyed by value so inline literals
  // (`pagination={{ mode: "layout" }}`) don't remount every render.
  // `appearance` is deliberately excluded: theme changes forward in place via
  // `setAppearance()` so the unit index, zoom, and Worker session survive.
  // `source` and the function resolvers are identity-keyed and must be
  // stable (see `UsePreviewOptions.source`). `initialIndex` and the starting
  // zoom ride in the key.
  const requestKey = JSON.stringify({ type, index: initialIndex, zoom: initialZoom, pagination });
  // Worker construction is transport, not content: read at mount through a
  // ref, never a dependency, so an inline factory can't remount every render.
  const workerRef = useRef(worker);
  workerRef.current = worker;
  // Value-key for the live appearance so inline literals
  // (`appearance={{ theme }}`) forward without remounting.
  const appearanceKey = JSON.stringify(appearance ?? null);
  const appearanceRef = useRef(appearance);
  appearanceRef.current = appearance;
  // Key the mount already painted with: the forwarding effect below skips
  // while they match, so mounting doesn't repaint once for free.
  const mountedAppearanceKey = useRef<string | null>(null);

  useEffect(() => {
    const host = hostRef.current;
    if (host === null) {
      setStatus("idle");
      return;
    }
    mountedAppearanceKey.current = JSON.stringify(appearanceRef.current ?? null);
    const controller = new AbortController();
    const signal = controller.signal;
    let handle: PreviewHandle | null = null;
    const attached: Array<{ type: string; listener: () => void }> = [];
    const detach = (): void => {
      if (handle !== null) {
        for (const { type, listener } of attached) handle.removeEventListener(type, listener);
        attached.length = 0;
      }
    };
    setStatus("loading");
    setViewer(null);
    viewerRef.current = null;
    setFormat(null);
    setInfo(null);
    setError(null);
    setUnitIndex(initialIndex);
    setZoomValue(initialZoom ?? "fit-width");
    const workerInit = workerRef.current;
    preview(source, host, {
      ...(type === undefined ? {} : { type }),
      index: initialIndex,
      ...(appearance === undefined ? {} : { appearance }),
      ...(initialZoom === undefined ? {} : { zoom: initialZoom }),
      ...(fontResolver === undefined ? {} : { fontResolver }),
      ...(mediaResolver === undefined ? {} : { mediaResolver }),
      ...(workerInit === undefined ? {} : { worker: workerInit }),
      ...(pagination === undefined ? {} : { pagination }),
      signal,
    }).then(
      (next) => {
        if (signal.aborted) {
          next.destroy();
          return;
        }
        handle = next;
        viewerRef.current = next;
        const onIndex = (): void => setUnitIndex(next.index);
        const onZoom = (): void => setZoomValue(next.zoom);
        next.addEventListener("indexchange", onIndex);
        next.addEventListener("zoomchange", onZoom);
        attached.push({ type: "indexchange", listener: onIndex }, { type: "zoomchange", listener: onZoom });
        setViewer(next);
        setFormat(next.format);
        setInfo(next.info);
        setUnitIndex(next.index);
        setZoomValue(next.zoom);
        setStatus("ready");
      },
      (failure: unknown) => {
        if (isCancellation(failure, signal)) return;
        setError(failure);
        setStatus("error");
      },
    );
    return () => {
      controller.abort();
      detach();
      handle?.destroy();
      if (viewerRef.current === handle) viewerRef.current = null;
    };
    // `initialIndex`, the starting zoom and `pagination` are read through
    // `requestKey`; `source` must be stable. `appearance` is read once for the
    // initial mount and forwarded after; `worker` is read through its ref.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, requestKey, fontResolver, mediaResolver, nonce]);

  // Live appearance: theme/chrome knobs apply to the mounted handle without
  // re-opening the document, so pager state survives dark/light toggles.
  // Runs on `viewer` too so an appearance change that landed mid-mount
  // still applies once the handle resolves. Skips while the viewer already
  // paints this key (initial mount), avoiding one free repaint per mount.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (viewer === null || appearanceKey === mountedAppearanceKey.current) return;
    const current = appearanceRef.current;
    if (current === undefined) return;
    mountedAppearanceKey.current = appearanceKey;
    viewer.setAppearance(current);
  }, [viewer, appearanceKey]);

  const goTo = useCallback((index: number): Promise<void> => {
    const current = viewerRef.current;
    if (current === null) return Promise.resolve();
    return current.goTo(index);
  }, []);

  const setZoom = useCallback((zoom: ZoomValue): void => {
    viewerRef.current?.setZoom(zoom);
  }, []);

  const setAppearance = useCallback((next: Appearance): void => {
    viewerRef.current?.setAppearance(next);
  }, []);

  const retry = useCallback((): void => {
    setNonce((n) => n + 1);
  }, []);

  return { ref: hostRef, status, viewer, format, info, index: unitIndex, zoom, error, goTo, setZoom, setAppearance, retry };
}

/**
 * Preview a presentation with the `usePreview()` lifecycle and `type`
 * asserted to `"pptx"` (detection skipped, only the PPTX chunk loads).
 * `info` is narrowed to `PptxInfo`.
 */
export function usePptx(options: UsePptxOptions): UsePreviewResult & { info: PptxInfo | null } {
  const result = usePreview({ ...options, type: "pptx" satisfies DocumentFormat });
  return { ...result, info: result.info as PptxInfo | null };
}

/**
 * Preview a workbook with the `usePreview()` lifecycle and `type`
 * asserted to `"xlsx"` (detection skipped, only the XLSX chunk loads).
 * `info` is narrowed to `XlsxInfo`. Sheet switching is `goTo()` by sheet
 * index; moving within the sheet is `viewer.xlsx` (`goToCell`, `scrollTo`).
 */
export function useXlsx(options: UseXlsxOptions): UsePreviewResult & { info: XlsxInfo | null } {
  const result = usePreview({ ...options, type: "xlsx" satisfies DocumentFormat });
  return { ...result, info: result.info as XlsxInfo | null };
}

/**
 * Preview a document with the `usePreview()` lifecycle and `type`
 * asserted to `"docx"` (detection skipped, only the DOCX chunk loads).
 * `info` is narrowed to `DocxInfo`. A document renders as one unit
 * (`info.count` is 1): `goTo(0)` is the only valid index and any other
 * rejects `not-found`. Moving within it is `viewer.docx` — `goToHeading()`
 * in both modes, `goToPage()` with `pagination: { mode: "layout" }`.
 */
export function useDocx(options: UseDocxOptions): UsePreviewResult & { info: DocxInfo | null } {
  const result = usePreview({ ...options, type: "docx" satisfies DocumentFormat });
  return { ...result, info: result.info as DocxInfo | null };
}
