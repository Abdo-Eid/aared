/**
 * `aared/react`: React hooks over `aared/preview`.
 *
 * `usePreview()` mounts `preview({ source, target })` inside an effect and owns
 * the parts every embed otherwise rewrites: the `AbortController`, the
 * loading/ready/error state, `index`/`zoom` mirroring from handle events,
 * and full teardown (`destroy()`) on unmount or source change. It is the
 * same open → render → navigate → destroy boundary as `aared/preview`:
 * `viewer.search()` finds places in the open document, and range reads and
 * media stay on the headless format entrypoints.
 *
 * The result is a union on `status`: once `status` is `"ready"`, `viewer`
 * (typed by format), `index` and `zoom` are there; once it is `"error"`,
 * `error` is an `AaredError`; before either, they are `null`. Destructuring
 * keeps the narrowing (`const { status, viewer } = ...`).
 *
 * `index` and `zoom` are start values; change them with `viewer.goTo()` and
 * `viewer.setZoom()`, and read their current values back from the result.
 * `appearance` updates in place when its values change, without reopening.
 * An imperative `setAppearance()` lasts until the option's values change.
 * Only `source`, `type`, `pagination` and `retry()` re-open the document.
 *
 * `usePptx()`, `useXlsx()`, and `useDocx()` are the same hook with `type`
 * asserted, so detection is skipped and only that format chunk is fetched
 * at runtime (bundlers still emit all three) — and `viewer` is that format's
 * handle. They add no lifecycle of their own.
 *
 * Import-time safe: importing this module touches no DOM or `Worker`
 * globals (SSR-safe imports) — everything runs inside effects. It lives
 * behind its own subpath with `react` as an optional peer, so vanilla
 * consumers never pay for it.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { RefObject } from "react";
import { AaredError, isAbort } from "../core/index.js";
import type { DocumentFormat, Source, ZoomValue } from "../core/index.js";
import { preview } from "../preview/index.js";
import type { DocxPreview, PptxPreview, PreviewHandle, PreviewOptions, XlsxPreview } from "../preview/index.js";

/**
 * Options for `usePreview()`: `PreviewOptions` plus the mounted source.
 * `signal` is hook-owned (one `AbortController` per mount) and therefore
 * omitted. `index` and `zoom` are start values, read at mount:
 * a later change does nothing until the next mount (a `source` change or
 * `retry()`); change what is on screen with `viewer.goTo()`,
 * `viewer.setZoom()`. `appearance` changes update the mounted viewer in place,
 * including resetting defaults when removed. `type` and `pagination`
 * re-open the document when their value changes. The function options
 * (`fontResolver`, `mediaResolver`, `worker`) never re-open it on their own,
 * so inline functions are fine: a media resolver is called through the latest
 * one given, and `fontResolver` and `worker` are read at the next mount.
 */
export interface UsePreviewOptions extends Omit<PreviewOptions, "signal"> {
  /** Previewed document. Must be referentially stable: a new identity
   * re-runs the mount, so pass the bytes you hold (a `Uint8Array` in state is
   * stable) rather than a copy made inline per render. `File`, `Blob`, and URL
   * strings are naturally stable. */
  source: Source;
}

/** Lifecycle phase of a mounted preview: the result's `status`. */
export type UsePreviewPhase = "idle" | "loading" | "ready" | "error";

/**
 * The mounted viewer as the hook hands it out: the `preview()` handle
 * without `destroy()`, since teardown is the hook's (unmount, a new
 * `source`, `retry()`); a viewer destroyed under it would leave the result
 * `"ready"` over nothing. Typed per format, so checking `format` still
 * narrows it (`viewer.xlsx` for a workbook).
 */
export type UsePreviewViewer<V extends PreviewHandle = PreviewHandle> = V extends PreviewHandle ? Omit<V, "destroy"> : never;

/**
 * What the result carries in every state: the host ref and `retry`. Changes
 * to what is on screen go through `viewer` (`goTo`, `setZoom`,
 * `setAppearance`), there once `status` is `"ready"`.
 */
interface UsePreviewActions {
  /** Attach to the host element: `<div ref={ref} />`. The hook owns what
   * is painted inside while mounted and clears it on teardown. */
  ref: RefObject<HTMLDivElement | null>;
  /** Re-run the current mount with the latest options (start values included). */
  retry: () => void;
}

/** Before a document is on screen (`"idle"` with no host yet, then `"loading"`): nothing to read. */
export interface UsePreviewPending extends UsePreviewActions {
  status: "idle" | "loading";
  viewer: null;
  index: null;
  zoom: null;
  error: null;
  navigationError: null;
}

/** A document on screen: the viewer, typed by format, and what it shows. */
export interface UsePreviewReady<V extends PreviewHandle = PreviewHandle> extends UsePreviewActions {
  status: "ready";
  /** The mounted handle (`viewer.info`, `viewer.format`, and `viewer.xlsx` / `viewer.docx` for those formats), without `destroy()`. */
  viewer: UsePreviewViewer<V>;
  /** Currently rendered unit index, mirrored from `indexchange`. */
  index: number;
  /** Active zoom, mirrored from `zoomchange`. The painted scale is `viewer.scale` (it changes on resize, with no event). */
  zoom: ZoomValue;
  error: null;
  /**
   * The last failed navigation, or `null`: whatever the mounted handle
   * reported through its `error` event, so `viewer.goTo()` and moves inside
   * the unit (`viewer.xlsx.goToCell`, `viewer.docx.goToHeading`, ...) all
   * land. What is on screen stays, so a host can say "that could not be
   * shown" without losing the view. Cleared when the unit changes and on
   * every mount; a cancelled or overtaken call never lands here.
   */
  navigationError: AaredError | null;
}

/** The mount failed (an aborted mount never lands here). */
export interface UsePreviewFailed extends UsePreviewActions {
  status: "error";
  viewer: null;
  index: null;
  zoom: null;
  /** Why; `retry()` mounts again. */
  error: AaredError;
  navigationError: null;
}

/**
 * What `usePreview()` and the per-format hooks return: one shape per
 * `status`. The object keeps its identity until something in it changes, so
 * it can be an effect dependency.
 */
export type UsePreviewResult<V extends PreviewHandle = PreviewHandle> = UsePreviewPending | UsePreviewReady<V> | UsePreviewFailed;

/** Options for `usePptx()`: `UsePreviewOptions` without `type`/`pagination`. */
export interface UsePptxOptions extends Omit<UsePreviewOptions, "type" | "pagination"> {}

/** Options for `useXlsx()`: `UsePreviewOptions` without `type`/`pagination`. */
export interface UseXlsxOptions extends Omit<UsePreviewOptions, "type" | "pagination"> {}

/** Options for `useDocx()`: `UsePreviewOptions` without `type`. */
export interface UseDocxOptions extends Omit<UsePreviewOptions, "type"> {}

function isCancellation(error: unknown, signal: AbortSignal): boolean {
  return signal.aborted || isAbort(error);
}

/** A failure as the result promises it: an `AaredError`, anything else wrapped as `internal` with itself as `cause`. */
function asAaredError(failure: unknown): AaredError {
  if (failure instanceof AaredError) return failure;
  const detail = failure instanceof Error ? failure.message : String(failure);
  return new AaredError("internal", `aared/react: ${detail}`, { cause: failure });
}

export function usePreview(options: UsePreviewOptions): UsePreviewResult {
  const { source, type, mediaResolver, pagination } = options;
  const hostRef = useRef<HTMLDivElement | null>(null);
  const [status, setStatus] = useState<UsePreviewPhase>("idle");
  const [viewer, setViewer] = useState<PreviewHandle | null>(null);
  const [shownIndex, setShownIndex] = useState(0);
  const [shownZoom, setShownZoom] = useState<ZoomValue | null>(null);
  const [error, setError] = useState<AaredError | null>(null);
  const [navigationError, setNavigationError] = useState<AaredError | null>(null);
  const [nonce, setNonce] = useState(0);

  // Everything a mount reads but that never re-opens the document by itself:
  // the start values (`index`, `zoom`) and the function options. Appearance
  // also uses the latest value when a viewer finishes loading.
  const latest = useRef(options);
  latest.current = options;

  // Structural options re-open the document. Keyed by value, so inline
  // literals (`pagination={{ mode: "layout" }}`) don't remount every render;
  // a media resolver re-opens it only when one appears or goes away (its
  // identity is read through `latest`).
  const mountKey = JSON.stringify({ type, pagination, media: mediaResolver !== undefined });
  // Compare fields, not object identity or property order: inline appearance
  // objects must not overwrite imperative changes on an unrelated re-render.
  const appearanceKey = JSON.stringify([
    options.appearance?.theme,
    options.appearance?.spreadsheet?.headers,
    options.appearance?.spreadsheet?.sheetTabs,
    options.appearance?.spreadsheet?.fill,
    options.appearance?.spreadsheet?.selection,
    options.appearance?.spreadsheet?.gridlines,
  ]);

  useEffect(() => {
    const host = hostRef.current;
    if (host === null) {
      setStatus("idle");
      return;
    }
    const start = latest.current;
    const controller = new AbortController();
    const signal = controller.signal;
    let handle: PreviewHandle | null = null;
    const attached: Array<{ type: string; listener: (event: Event) => void }> = [];
    setStatus("loading");
    setViewer(null);
    setError(null);
    setNavigationError(null);
    // A media resolver goes through the latest one given (the one this mount
    // started with, should it be removed while a picture is still loading).
    const startMedia = start.mediaResolver;
    preview({
      source,
      target: host,
      ...(start.type === undefined ? {} : { type: start.type }),
      index: start.index ?? 0,
      ...(start.appearance === undefined ? {} : { appearance: start.appearance }),
      ...(start.zoom === undefined ? {} : { zoom: start.zoom }),
      ...(start.fontResolver === undefined ? {} : { fontResolver: start.fontResolver }),
      ...(startMedia === undefined ? {} : { mediaResolver: (mediaId: string) => (latest.current.mediaResolver ?? startMedia)(mediaId) }),
      ...(start.worker === undefined ? {} : { worker: start.worker }),
      ...(start.pagination === undefined ? {} : { pagination: start.pagination }),
      signal,
    }).then(
      (next) => {
        if (signal.aborted) {
          next.destroy();
          return;
        }
        handle = next;
        // A unit change clears a failed navigation; the handle's `error` event
        // reports one, whatever made it (`viewer.goTo()`, a move on `viewer.xlsx`).
        const onIndex = (): void => {
          setShownIndex(next.index);
          setNavigationError(null);
        };
        const onZoom = (): void => setShownZoom(next.zoom);
        const onError = (event: Event): void => {
          setNavigationError(asAaredError((event as CustomEvent<{ cause: unknown }>).detail.cause));
        };
        for (const [eventType, listener] of [["indexchange", onIndex], ["zoomchange", onZoom], ["error", onError]] as const) {
          next.addEventListener(eventType, listener);
          attached.push({ type: eventType, listener });
        }
        setShownIndex(next.index);
        setShownZoom(next.zoom);
        setViewer(next);
        setStatus("ready");
      },
      (failure: unknown) => {
        if (isCancellation(failure, signal)) return;
        setError(asAaredError(failure));
        setStatus("error");
      },
    );
    return () => {
      controller.abort();
      if (handle !== null) {
        for (const { type, listener } of attached) handle.removeEventListener(type, listener);
      }
      handle?.destroy();
    };
    // `type`, `pagination` and whether there is a media resolver are read
    // through `mountKey`; `source` must be stable; everything else (the start
    // values included) is read through `latest` at mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, mountKey, nonce]);

  useEffect(() => {
    viewer?.setAppearance(latest.current.appearance ?? {});
  }, [viewer, appearanceKey]);

  const retry = useCallback((): void => {
    setNonce((n) => n + 1);
  }, []);

  return useMemo((): UsePreviewResult => {
    const actions = { ref: hostRef, retry };
    if (status === "ready" && viewer !== null) {
      return { ...actions, status, viewer, index: shownIndex, zoom: shownZoom ?? viewer.zoom, error: null, navigationError };
    }
    if (status === "error" && error !== null) {
      return { ...actions, status, viewer: null, index: null, zoom: null, error, navigationError: null };
    }
    return {
      ...actions,
      status: status === "idle" ? "idle" : "loading",
      viewer: null,
      index: null,
      zoom: null,
      error: null,
      navigationError: null,
    };
  }, [status, viewer, shownIndex, shownZoom, error, navigationError, retry]);
}

/**
 * Preview a presentation with the `usePreview()` lifecycle and `type`
 * asserted to `"pptx"` (detection skipped, only the PPTX chunk loads).
 * Once ready, `viewer` is a `PptxPreview`.
 */
export function usePptx(options: UsePptxOptions): UsePreviewResult<PptxPreview> {
  // An asserted type fails the mount on anything but a presentation.
  return usePreview({ ...options, type: "pptx" satisfies DocumentFormat }) as UsePreviewResult<PptxPreview>;
}

/**
 * Preview a workbook with the `usePreview()` lifecycle and `type`
 * asserted to `"xlsx"` (detection skipped, only the XLSX chunk loads).
 * Once ready, `viewer` is an `XlsxPreview`. Sheet switching is `viewer.goTo()`
 * by sheet index (or a sheet tab); moving within the sheet is `viewer.xlsx`
 * (`goToCell`, `scrollTo`).
 */
export function useXlsx(options: UseXlsxOptions): UsePreviewResult<XlsxPreview> {
  return usePreview({ ...options, type: "xlsx" satisfies DocumentFormat }) as UsePreviewResult<XlsxPreview>;
}

/**
 * Preview a document with the `usePreview()` lifecycle and `type`
 * asserted to `"docx"` (detection skipped, only the DOCX chunk loads).
 * Once ready, `viewer` is a `DocxPreview`. A document renders as one unit
 * (`info.count` is 1): `viewer.goTo(0)` is the only valid index and any other
 * rejects `not-found`. Moving within it is `viewer.docx` — `goToHeading()`
 * in both modes, `goToPage()` with `pagination: { mode: "layout" }`.
 */
export function useDocx(options: UseDocxOptions): UsePreviewResult<DocxPreview> {
  return usePreview({ ...options, type: "docx" satisfies DocumentFormat }) as UsePreviewResult<DocxPreview>;
}
