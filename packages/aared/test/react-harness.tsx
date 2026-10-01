import { StrictMode } from "react";
import type { ReactElement } from "react";
import { createRoot } from "react-dom/client";
import type { Root } from "react-dom/client";
import { useDocx, usePreview, usePptx, useXlsx } from "../src/react/index.js";
import type { UsePreviewResult } from "../src/react/index.js";
import type { ZoomValue } from "../src/core/index.js";

type HookWindow = Window &
  typeof globalThis & {
    probe?:
      | {
          status: string;
          format: string | null;
          count: number | null;
          index: number;
          zoom: unknown;
          error: string | null;
          navigationError: string | null;
        }
      | undefined;
    hookActions?:
      | {
          goTo: (index: number) => Promise<void>;
          setZoom: (zoom: ZoomValue) => void;
          retry: () => void;
        }
      | undefined;
    mountHook?: (kind: string, source: string, strict?: boolean) => void;
    unmountHook?: () => number;
    hookGoTo?: (index: number) => Promise<unknown>;
    hookGoToFailure?: (index: number) => Promise<{ code: unknown; message: string } | null>;
    viewerGoToFailure?: (index: number) => Promise<{ code: unknown; message: string } | null>;
    hookViewer?: UsePreviewResult["viewer"] | undefined;
    goToWhileLoading?: () => Promise<{ code: unknown; message: string } | "mounted-too-fast">;
    hookSetZoom?: (zoom: ZoomValue) => Promise<unknown>;
    rerenderHook?: () => void;
    workerFactoryCalls?: number;
    snapshot?: () => unknown;
  };

const w = window as HookWindow;

function publish(result: UsePreviewResult): void {
  w.probe = {
    status: result.status,
    format: result.format,
    count: result.info?.count ?? null,
    index: result.index,
    zoom: result.zoom,
    error: result.error == null
      ? null
      : String((result.error as Error)?.message ?? result.error),
    navigationError: result.navigationError == null
      ? null
      : String((result.navigationError as Error)?.message ?? result.navigationError),
  };
  w.hookActions = { goTo: result.goTo, setZoom: result.setZoom, retry: result.retry };
  w.hookViewer = result.viewer;
}

function PptxProbe({ source }: { source: string }) {
  const result = usePptx({ source });
  publish(result);
  return <div ref={result.ref} />;
}

function XlsxProbe({ source }: { source: string }) {
  const result = useXlsx({ source });
  publish(result);
  return <div ref={result.ref} />;
}

function DocxProbe({ source }: { source: string }) {
  const result = useDocx({ source, pagination: { mode: "layout" } });
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
  w.hookActions = undefined;
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
    : <AutoProbe source={source} />;
  mounted = { node, strict: strict === true };
  root.render(mounted.strict ? <StrictMode>{node}</StrictMode> : node);
};

/** Render the mounted probe again with the same props (a parent re-render). */
w.rerenderHook = () => {
  if (root === null || mounted === null) return;
  const node = { ...mounted.node };
  root.render(mounted.strict ? <StrictMode>{node}</StrictMode> : node);
};

w.hookGoTo = async (index: number) => {
  await w.hookActions?.goTo(index);
  return w.probe;
};

/** `goTo(index)` expected to fail: the rejection's code and message, or null if it resolved. */
w.hookGoToFailure = async (index: number) => {
  try {
    await w.hookActions?.goTo(index);
    return null;
  } catch (error) {
    return { code: (error as { code?: unknown }).code, message: String((error as Error).message) };
  }
};

/** The same, through the handle itself (`viewer.goTo`), bypassing the hook's `goTo`. */
w.viewerGoToFailure = async (index: number) => {
  try {
    await w.hookViewer?.goTo(index);
    return null;
  } catch (error) {
    return { code: (error as { code?: unknown }).code, message: String((error as Error).message) };
  }
};

/** Call `goTo(0)` the moment the hook reports "loading", before any viewer exists. */
w.goToWhileLoading = async () => {
  while (w.probe?.status !== "loading") {
    if (w.probe?.status === "ready") return "mounted-too-fast";
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  return (await w.hookGoToFailure!(0)) ?? { code: "resolved", message: "" };
};

w.hookSetZoom = async (zoom: ZoomValue) => {
  w.hookActions?.setZoom(zoom);
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
