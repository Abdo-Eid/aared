import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { usePreview } from "aared/react";
import type { OutlineEntry } from "../adapters";
import type { BenchTheme } from "../adapters";
import type { DocMeta, RunRecord } from "../lib/doc";
import { ViewerFrame, useViewerFrame, type Status } from "./ViewerFrame";

interface DocxExtra {
  outline: OutlineEntry[];
  pages: Array<{ width: number; height: number }>;
  paperWidth: number;
}

/**
 * The aared column, built the way a consumer app builds it: one
 * `usePreview()` hook owning detection, open timing, navigation guards,
 * and teardown. The bench adds only what its UI needs on top: the
 * open-speed `ms`, an explicit theme (never `"system"`, so the preview
 * follows the bench chrome instead of the OS preference), DOCX layout
 * mode, and the outline/page geometry the rail and pager read.
 */
export function AaredViewerCard({ doc, theme, onClose, onDone }: {
  doc: DocMeta;
  theme: BenchTheme;
  onClose: () => void;
  onDone: (record: RunRecord) => void;
}) {
  // Stable byte identity: a fresh ArrayBuffer per render would remount the
  // preview in a loop, so the copy is memoized (the hook documents this).
  const buffer = useMemo(() => doc.bytes.slice().buffer as ArrayBuffer, [doc.bytes]);
  const t0 = useMemo(() => performance.now(), [buffer]);
  const hook = usePreview({
    source: buffer,
    appearance: { theme },
    docx: { pagination: { mode: "layout" } },
  });

  const [docxExtra, setDocxExtra] = useState<DocxExtra | null>(null);
  const [tabError, setTabError] = useState<{ message: string; ms: number } | null>(null);

  // Fresh bytes restart per-tab error state alongside the hook's own reset.
  useEffect(() => {
    setTabError(null);
  }, [buffer]);

  // Layout pages and headings are derived render output, so they are scraped
  // from the painted DOM once ready — the same read the old adapter did.
  useEffect(() => {
    if (hook.status !== "ready" || hook.viewer?.format !== "docx") {
      setDocxExtra(null);
      return;
    }
    const host = hook.ref.current;
    if (host === null) return;
    const outline: OutlineEntry[] = [];
    host.querySelectorAll("h1,h2,h3,h4,h5,h6").forEach((element) => {
      const level = Number(element.tagName.slice(1));
      const text = (element.textContent ?? "").trim();
      if (Number.isInteger(level) && text !== "") outline.push({ level, text, element });
    });
    // Layout pages carry their geometry as inline styles, so this reads the
    // authored page size before any bench zoom scaling applies.
    const pages = [...host.querySelectorAll<HTMLElement>('section[data-aared-element="page"]')]
      .map((element) => ({
        width: Number.parseFloat(element.style.width) || 0,
        height: Number.parseFloat(element.style.height) || 0,
      }));
    setDocxExtra({
      outline,
      pages,
      paperWidth: pages.length > 0 ? Math.max(...pages.map((page) => page.width)) : 0,
    });
  }, [hook.status, hook.viewer, hook.ref]);

  const status: Status = useMemo(() => {
    if (tabError !== null) return { state: "error", message: tabError.message, ms: tabError.ms };
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
    if (docxExtra === null) return { state: "loading" };
    return {
      state: "ready",
      handle: {
        format: "docx",
        pageCount: docxExtra.pages.length,
        pages: docxExtra.pages,
        ms,
        outline: docxExtra.outline,
        paperWidth: docxExtra.paperWidth,
        destroy: () => viewer.destroy(),
      },
    };
  }, [hook.status, hook.viewer, hook.info, hook.error, docxExtra, tabError, t0]);

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
      onDone({
        key: `${doc.key}:aared`,
        name: doc.name,
        format: doc.format,
        renderer: "aared",
        ms: status.handle.ms,
        sizeBytes: doc.sizeBytes,
        at: Date.now(),
      });
    } else {
      const key = `error:${status.ms}:${status.message}`;
      if (firedKey.current === key) return;
      firedKey.current = key;
      console.error(`[bench] ${doc.name} via aared failed:`, status.message);
      onDone({
        key: `${doc.key}:aared`,
        name: doc.name,
        format: doc.format,
        renderer: "aared",
        ms: status.ms,
        sizeBytes: doc.sizeBytes,
        error: status.message,
        at: Date.now(),
      });
    }
  }, [status, doc, onDone]);

  const reportError = useCallback((message: string, ms: number) => {
    setTabError({ message, ms });
  }, []);

  const controller = useViewerFrame({
    doc,
    theme,
    status,
    hostRef: hook.ref,
    reportError,
  });
  return <ViewerFrame doc={doc} renderer="aared" onClose={onClose} controller={controller} />;
}
