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
 * asserted, so detection is skipped and only that format chunk loads —
 * plus a narrowed `info` type. They add no lifecycle of their own.
 *
 * Import-time safe: importing this module touches no DOM or `Worker`
 * globals (SSR-safe imports) — everything runs inside effects. It lives
 * behind its own subpath with `react` as an optional peer, so vanilla
 * consumers never pay for it.
 */

import { useCallback, useEffect, useRef, useState } from "react";
import type { RefObject } from "react";
import { isAaredError } from "../core/index.js";
import type { DocumentFormat, DocumentInfo, Source, ZoomValue } from "../core/index.js";
import type { DocxInfo } from "../docx/index.js";
import type { PptxInfo } from "../pptx/index.js";
import type { XlsxInfo } from "../xlsx/index.js";
import { preview } from "../preview/index.js";
import type { PreviewHandle, PreviewOptions } from "../preview/index.js";

/**
 * Options for `usePreview()`: `PreviewOptions` plus the mounted source.
 * `signal` is hook-owned (one `AbortController` per mount) and therefore
 * omitted — pass nothing abort-related.
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
  /** The mounted handle once `status` is `"ready"`; otherwise `null`. */
  viewer: PreviewHandle | null;
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
  /** Re-run the current mount with the latest options. */
  retry: () => void;
}

/** Options for `usePptx()`: `UsePreviewOptions` without `type`/`docx`. */
export interface UsePptxOptions extends Omit<UsePreviewOptions, "type" | "docx"> {}

/** Options for `useXlsx()`: `UsePreviewOptions` without `type`/`docx`. */
export interface UseXlsxOptions extends Omit<UsePreviewOptions, "type" | "docx"> {}

/** Options for `useDocx()`: `UsePreviewOptions` without `type`. */
export interface UseDocxOptions extends Omit<UsePreviewOptions, "type"> {}

function isCancellation(error: unknown, signal: AbortSignal): boolean {
  if (signal.aborted) return true;
  if (error instanceof DOMException && error.name === "AbortError") return true;
  return isAaredError(error) && error.code === "aborted";
}

export function usePreview({
  source,
  type,
  index: initialIndex = 0,
  appearance,
  fontResolver,
  mediaResolver,
  docx,
}: UsePreviewOptions): UsePreviewResult {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const viewerRef = useRef<PreviewHandle | null>(null);
  const [status, setStatus] = useState<UsePreviewPhase>("idle");
  const [viewer, setViewer] = useState<PreviewHandle | null>(null);
  const [info, setInfo] = useState<DocumentInfo | null>(null);
  const [unitIndex, setUnitIndex] = useState<number>(initialIndex);
  const [zoom, setZoomValue] = useState<ZoomValue>("fit-width");
  const [error, setError] = useState<unknown>(null);
  const [nonce, setNonce] = useState(0);

  // Plain-data options are keyed by value so inline literals
  // (`appearance={{ theme }}`) don't remount every render. `source` and
  // the function resolvers are identity-keyed and must be stable (see
  // `UsePreviewOptions.source`). `initialIndex` rides inside the key.
  const requestKey = JSON.stringify({ type, index: initialIndex, appearance, docx });

  useEffect(() => {
    const host = hostRef.current;
    if (host === null) {
      setStatus("idle");
      return;
    }
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
    setInfo(null);
    setError(null);
    setUnitIndex(initialIndex);
    setZoomValue("fit-width");
    preview(source, host, {
      ...(type === undefined ? {} : { type }),
      index: initialIndex,
      ...(appearance === undefined ? {} : { appearance }),
      ...(fontResolver === undefined ? {} : { fontResolver }),
      ...(mediaResolver === undefined ? {} : { mediaResolver }),
      ...(docx === undefined ? {} : { docx }),
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
    // `initialIndex` is read through `requestKey`; `source` must be stable.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, requestKey, fontResolver, mediaResolver, nonce]);

  const goTo = useCallback((index: number): Promise<void> => {
    const current = viewerRef.current;
    if (current === null) return Promise.resolve();
    return current.goTo(index);
  }, []);

  const setZoom = useCallback((zoom: ZoomValue): void => {
    viewerRef.current?.setZoom(zoom);
  }, []);

  const retry = useCallback((): void => {
    setNonce((n) => n + 1);
  }, []);

  return { ref: hostRef, status, viewer, info, index: unitIndex, zoom, error, goTo, setZoom, retry };
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
 * `info` is narrowed to `XlsxInfo`. Sheet switching is `goTo()` by
 * sheet index; cell-level navigation stays on `aared/xlsx`.
 */
export function useXlsx(options: UseXlsxOptions): UsePreviewResult & { info: XlsxInfo | null } {
  const result = usePreview({ ...options, type: "xlsx" satisfies DocumentFormat });
  return { ...result, info: result.info as XlsxInfo | null };
}

/**
 * Preview a document with the `usePreview()` lifecycle and `type`
 * asserted to `"docx"` (detection skipped, only the DOCX chunk loads).
 * `info` is narrowed to `DocxInfo`. DOCX renders as one unit, so `goTo()`
 * is a no-op here — heading navigation stays on `aared/docx`.
 */
export function useDocx(options: UseDocxOptions): UsePreviewResult & { info: DocxInfo | null } {
  const result = usePreview({ ...options, type: "docx" satisfies DocumentFormat });
  return { ...result, info: result.info as DocxInfo | null };
}
