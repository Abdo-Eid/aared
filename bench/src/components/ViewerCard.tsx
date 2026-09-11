import { useCallback, useEffect, useRef, useState } from "react";
import { ADAPTERS, isAbort } from "../adapters";
import type { OpenedDoc } from "../adapters";
import type { BenchTheme } from "../adapters";
import type { DocMeta, Renderer, RunRecord } from "../lib/doc";
import { ViewerFrame, useViewerFrame, type Status } from "./ViewerFrame";
import { AaredViewerCard } from "./AaredViewerCard";

interface ViewerCardProps {
  doc: DocMeta;
  renderer: Renderer;
  theme: BenchTheme;
  onClose: () => void;
  onDone: (record: RunRecord) => void;
}

/**
 * One rendered document: the aared column demonstrates the consumer
 * integration (`usePreview()` in `AaredViewerCard`), while baselines keep
 * their adapter open effect below. Both share every interaction through
 * `ViewerFrame`, so the timing comparison stays apples-to-apples.
 */
export function ViewerCard({ doc, renderer, theme, onClose, onDone }: ViewerCardProps) {
  if (renderer === "aared") {
    return <AaredViewerCard doc={doc} theme={theme} onClose={onClose} onDone={onDone} />;
  }
  return <BaselineViewerCard doc={doc} theme={theme} onClose={onClose} onDone={onDone} />;
}

function BaselineViewerCard({ doc, theme, onClose, onDone }: {
  doc: DocMeta;
  theme: BenchTheme;
  onClose: () => void;
  onDone: (record: RunRecord) => void;
}) {
  const [status, setStatus] = useState<Status>({ state: "loading" });
  const hostRef = useRef<HTMLDivElement | null>(null);
  const reportError = useCallback((message: string, ms: number) => {
    setStatus({ state: "error", message, ms });
  }, []);
  const controller = useViewerFrame({ doc, theme, status, hostRef, reportError });

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const ctrl = new AbortController();
    const t0 = performance.now();
    let handle: OpenedDoc | null = null;
    setStatus({ state: "loading" });

    ADAPTERS[doc.format](doc.bytes, host, ctrl.signal).then(
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
        onDone({
          key: `${doc.key}:baseline`,
          name: doc.name,
          format: doc.format,
          renderer: "baseline",
          ms: opened.ms,
          sizeBytes: doc.sizeBytes,
          at: Date.now(),
        });
      },
      (error: unknown) => {
        if (ctrl.signal.aborted || isAbort(error)) return;
        const ms = Math.round(performance.now() - t0);
        const message =
          error instanceof Error ? error.message : "Could not open this file.";
        console.error(`[bench] ${doc.name} via baseline failed:`, message);
        setStatus({ state: "error", message, ms });
        onDone({
          key: `${doc.key}:baseline`,
          name: doc.name,
          format: doc.format,
          renderer: "baseline",
          ms,
          sizeBytes: doc.sizeBytes,
          error: message,
          at: Date.now(),
        });
      },
    );

    return () => {
      ctrl.abort();
      handle?.destroy();
    };
  }, [doc, onDone]);

  return <ViewerFrame doc={doc} renderer="baseline" onClose={onClose} controller={controller} />;
}
