import { StrictMode } from "react";
import type { ReactElement } from "react";
import { createRoot } from "react-dom/client";
import type { Root } from "react-dom/client";
import { useDocx, usePreview, usePptx, useXlsx } from "../src/react/index.js";
import type { UsePreviewResult } from "../src/react/index.js";
import type { Appearance, ViewerTheme, ZoomValue } from "../src/core/index.js";

type HookWindow = Window &
  typeof globalThis & {
    probe?:
      | {
          status: string;
          format: string | null;
          count: number | null;
          index: number | null;
          zoom: unknown;
          error: string | null;
          navigationError: string | null;
        }
      | undefined;
    /** Distinct viewers a probe has seen (a remount makes a new one) and distinct result objects. */
    seen?: { viewers: number; results: number } | undefined;
    setStartProps?: (props: StartProps) => void;
    hookRetry?: (() => void) | undefined;
    mountHook?: (kind: string, source: string, strict?: boolean) => void;
    unmountHook?: () => number;
    viewerGoTo?: (index: number) => Promise<unknown>;
    viewerGoToFailure?: (index: number) => Promise<{ code: unknown; message: string } | null>;
    hookViewer?: UsePreviewResult["viewer"] | undefined;
    viewerSetZoom?: (zoom: ZoomValue) => Promise<unknown>;
    rerenderHook?: () => void;
    workerFactoryCalls?: number;
    snapshot?: () => unknown;
  };

const w = window as HookWindow;

let lastViewer: unknown = null;
let lastResult: unknown = null;

function publish(result: UsePreviewResult): void {
  w.probe = {
    status: result.status,
    format: result.viewer?.format ?? null,
    count: result.viewer?.info.count ?? null,
    index: result.index,
    zoom: result.zoom,
    // Typed by the union: an AaredError when failed, null otherwise; no casts.
    error: result.error?.message ?? null,
    navigationError: result.navigationError?.message ?? null,
  };
  const seen = w.seen ?? { viewers: 0, results: 0 };
  if (result.viewer !== null && result.viewer !== lastViewer) seen.viewers += 1;
  if (result !== lastResult) seen.results += 1;
  lastViewer = result.viewer ?? lastViewer;
  lastResult = result;
  w.seen = seen;
  w.hookRetry = result.retry;
  w.hookViewer = result.viewer;
}

function PptxProbe({ source }: { source: string }) {
  const result = usePptx({ source });
  // Compile-time check: destructuring keeps the narrowing, and the viewer is a PPTX one.
  const { status, viewer } = result;
  if (status === "ready") {
    const format: "pptx" = viewer.info.format;
    void format;
    // @ts-expect-error teardown is the hook's: its viewer has no destroy().
    void viewer.destroy;
    // The hook's viewer searches its deck, with PPTX matches. Type-checked,
    // never called: a search started on each render could still be reading
    // slides at unmount, and it would reject with `closed`, unhandled.
    void (async () => (await viewer.search("x"))[0]?.slide.index satisfies number | undefined);
  }
  publish(result);
  return <div ref={result.ref} />;
}

function XlsxProbe({ source }: { source: string }) {
  const result = useXlsx({ source });
  // Compile-time check: a ready XLSX result has non-null cell controls.
  if (result.status === "ready") void result.viewer.xlsx.selection;
  publish(result);
  return <div ref={result.ref} />;
}

function DocxProbe({ source }: { source: string }) {
  const result = useDocx({ source, pagination: { mode: "layout" } });
  // Compile-time check: a ready DOCX result has non-null page controls.
  if (result.status === "ready") void result.viewer.docx.pageCount;
  publish(result);
  return <div ref={result.ref} />;
}

interface StartProps {
  zoom?: ZoomValue;
  index?: number;
  theme?: ViewerTheme;
  appearance?: Appearance;
  /** Pass a media resolver written inline: a new function every render. */
  inlineMedia?: boolean;
}

/** Index/zoom start values, reactive appearance, and stable inline resolvers. */
function StartProbe({ source, props }: { source: string; props: StartProps }) {
  const appearance = props.appearance ?? (props.theme === undefined ? undefined : { theme: props.theme });
  const result = usePreview({
    source,
    ...(props.zoom === undefined ? {} : { zoom: props.zoom }),
    ...(props.index === undefined ? {} : { index: props.index }),
    ...(appearance === undefined ? {} : { appearance }),
    ...(props.inlineMedia === true ? { mediaResolver: () => new Promise<ArrayBuffer>(() => {}) } : {}),
  });
  publish(result);
  return <div ref={result.ref} />;
}

function AutoProbe({ source }: { source: string }) {
  const result = usePreview({ source });
  publish(result);
  return <div ref={result.ref} />;
}

/** An inline worker factory: a new identity every render, which must not remount. */
function WorkerProbe({ source }: { source: string }) {
  const result = usePreview({
    source,
    worker: () => {
      w.workerFactoryCalls = (w.workerFactoryCalls ?? 0) + 1;
      return new Worker("/custom-worker.mjs", { type: "module" });
    },
  });
  publish(result);
  return <div ref={result.ref} />;
}

let root: Root | null = null;
let mounted: { node: ReactElement; strict: boolean } | null = null;

function unmountHook(): number {
  root?.unmount();
  root = null;
  w.probe = undefined;
  w.hookRetry = undefined;
  w.hookViewer = undefined;
  w.seen = undefined;
  lastViewer = null;
  lastResult = null;
  return document.querySelector("#hook-root > div")?.childElementCount ?? 0;
}

w.unmountHook = unmountHook;

w.mountHook = (kind: string, source: string, strict?: boolean) => {
  unmountHook();
  const stage = document.querySelector("#stage");
  if (stage === null) throw new Error("react harness: missing #stage");
  stage.replaceChildren();
  const holder = document.createElement("div");
  holder.id = "hook-root";
  stage.append(holder);
  root = createRoot(holder);
  const node = kind === "pptx"
    ? <PptxProbe source={source} />
    : kind === "xlsx"
    ? <XlsxProbe source={source} />
    : kind === "docx"
    ? <DocxProbe source={source} />
    : kind === "worker"
    ? <WorkerProbe source={source} />
    : kind === "start"
    ? <StartProbe source={source} props={{ zoom: "fit-page" }} />
    : <AutoProbe source={source} />;
  mounted = { node, strict: strict === true };
  root.render(mounted.strict ? <StrictMode>{node}</StrictMode> : node);
};

/** Re-render the start probe with new options (a parent changing its props). */
w.setStartProps = (props: StartProps) => {
  if (root === null || mounted === null) return;
  const source = (mounted.node.props as { source: string }).source;
  mounted = { ...mounted, node: <StartProbe source={source} props={props} /> };
  root.render(mounted.strict ? <StrictMode>{mounted.node}</StrictMode> : mounted.node);
};

/** Render the mounted probe again with the same props (a parent re-render). */
w.rerenderHook = () => {
  if (root === null || mounted === null) return;
  const node = { ...mounted.node };
  root.render(mounted.strict ? <StrictMode>{node}</StrictMode> : node);
};

/** Move through the hook's viewer (`viewer.goTo`), the one way to change the unit. */
w.viewerGoTo = async (index: number) => {
  await w.hookViewer?.goTo(index);
  return w.probe;
};

/** `viewer.goTo(index)` expected to fail: the rejection's code and message, or null if it resolved. */
w.viewerGoToFailure = async (index: number) => {
  try {
    await w.hookViewer?.goTo(index);
    return null;
  } catch (error) {
    return { code: (error as { code?: unknown }).code, message: String((error as Error).message) };
  }
};

w.viewerSetZoom = async (zoom: ZoomValue) => {
  w.hookViewer?.setZoom(zoom);
  return w.probe;
};

w.snapshot = () => {
  const host = document.querySelector("#hook-root");
  return {
    text: host?.textContent ?? "",
    children: host?.querySelector("div")?.childElementCount ?? 0,
    canvasCount: host?.querySelectorAll("canvas").length ?? 0,
    cells: host?.querySelectorAll('td[data-aared-element="cell"]').length ?? 0,
    pages: host?.querySelectorAll('section[data-aared-element="page"]').length ?? 0,
    probe: w.probe ?? null,
  };
};
