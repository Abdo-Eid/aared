import { StrictMode } from "react";
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
    hookSetZoom?: (zoom: ZoomValue) => Promise<unknown>;
    snapshot?: () => unknown;
  };

const w = window as HookWindow;

function publish(result: UsePreviewResult): void {
  w.probe = {
    status: result.status,
    format: result.viewer?.format ?? null,
    count: result.info?.count ?? null,
    index: result.index,
    zoom: result.zoom,
    error: result.error == null
      ? null
      : String((result.error as Error)?.message ?? result.error),
  };
  w.hookActions = { goTo: result.goTo, setZoom: result.setZoom, retry: result.retry };
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
  const result = useDocx({ source, docx: { pagination: { mode: "layout" } } });
  publish(result);
  return <div ref={result.ref} />;
}

function AutoProbe({ source }: { source: string }) {
  const result = usePreview({ source });
  publish(result);
  return <div ref={result.ref} />;
}

let root: Root | null = null;

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
    : <AutoProbe source={source} />;
  root.render(strict === true ? <StrictMode>{node}</StrictMode> : node);
};

w.hookGoTo = async (index: number) => {
  await w.hookActions?.goTo(index);
  return w.probe;
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
