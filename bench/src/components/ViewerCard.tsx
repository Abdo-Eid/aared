import { useCallback, useEffect, useRef, useState } from "react";
import { ADAPTERS, BASELINES, DEFAULT_BASELINE, isAbort } from "../adapters";
import type { OpenedDoc } from "../adapters";
import type { BenchTheme } from "../adapters";
import type { DocMeta, Renderer } from "../lib/doc";
import { ViewerFrame, useViewerFrame, type Status } from "./ViewerFrame";
import { AaredViewerCard } from "./AaredViewerCard";

interface ViewerCardProps {
  doc: DocMeta;
  renderer: Renderer;
  theme: BenchTheme;
  onClose: () => void;
}

/**
 * One rendered document: the aared column demonstrates the consumer
 * integration (`usePreview()` in `AaredViewerCard`), while baselines keep
 * their adapter open effect below. Both share every interaction through
 * `ViewerFrame`, so the timing comparison stays apples-to-apples.
 */
export function ViewerCard({ doc, renderer, theme, onClose }: ViewerCardProps) {
  if (renderer === "aared") {
    return <AaredViewerCard doc={doc} theme={theme} onClose={onClose} />;
  }
  return <BaselineViewerCard doc={doc} renderer={renderer} theme={theme} onClose={onClose} />;
}

function BaselineViewerCard({ doc, renderer, theme, onClose }: {
  doc: DocMeta;
  renderer: Renderer;
  theme: BenchTheme;
  onClose: () => void;
}) {
  const [status, setStatus] = useState<Status>({ state: "loading" });
  const hostRef = useRef<HTMLDivElement | null>(null);
  const reportError = useCallback((message: string, ms: number) => {
    setStatus({ state: "error", message, ms });
  }, []);
  const controller = useViewerFrame({ doc, theme, frameId: `${doc.key}:${renderer}`, renderer, status, hostRef, reportError });

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const ctrl = new AbortController();
    const t0 = performance.now();
    let handle: OpenedDoc | null = null;
    setStatus({ state: "loading" });

    const baselines = BASELINES[doc.format];
    const adapter =
      (renderer === "baseline"
        ? baselines[DEFAULT_BASELINE[doc.format]]
        : baselines[renderer]) ?? ADAPTERS[doc.format];
    adapter(doc.bytes, host, ctrl.signal).then(
      (opened) => {
        if (ctrl.signal.aborted) {
          opened.destroy();
          return;
        }
        handle = opened;
        setStatus({ state: "ready", handle: opened });
        console.info(
          `[bench] ${doc.name} via baseline: ${opened.ms} ms, ${opened.pageCount} pages`,
        );
      },
      (error: unknown) => {
        if (ctrl.signal.aborted || isAbort(error)) return;
        const ms = Math.round(performance.now() - t0);
        const message =
          error instanceof Error ? error.message : "Could not open this file.";
        console.error(`[bench] ${doc.name} via baseline failed:`, message);
        setStatus({ state: "error", message, ms });
      },
    );

    return () => {
      ctrl.abort();
      handle?.destroy();
    };
  }, [doc, renderer]);

  return <ViewerFrame doc={doc} renderer={renderer} onClose={onClose} controller={controller} />;
}
