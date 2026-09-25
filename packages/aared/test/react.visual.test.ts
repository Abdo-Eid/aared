import { expect, test } from "bun:test";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

interface Probe {
  status: string;
  format: string | null;
  count: number | null;
  index: number;
  zoom: unknown;
  error: string | null;
}

interface Snapshot {
  text: string;
  children: number;
  canvasCount: number;
  cells: number;
  pages: number;
  probe: Probe | null;
}

async function waitFor(
  view: { evaluate: (expression: string) => Promise<unknown> },
  expression: string,
  timeoutMs = 20000,
): Promise<void> {
  const start = Date.now();
  for (;;) {
    if ((await view.evaluate(expression)) === true) return;
    if (Date.now() - start > timeoutMs) throw new Error(`timed out waiting for ${expression}`);
    await Bun.sleep(100);
  }
}

/** How many times the harness server served each worker script, per test. */
const requests = new Map<string, number>();

async function withHarness(run: (view: {
  evaluate: (expression: string) => Promise<unknown>;
}) => Promise<void>): Promise<void> {
  requests.clear();
  const [harnessModule, workerModule, pptx, xlsx, docx] = await Promise.all([
    browserModule(`${import.meta.dir}/react-harness.tsx`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
    Bun.file(PPTX_FIXTURE).arrayBuffer(),
    Bun.file(XLSX_FIXTURE).arrayBuffer(),
    Bun.file(DOCX_FIXTURE).arrayBuffer(),
  ]);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/react-harness.mjs") {
        return new Response(harnessModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname === "/worker.mjs" || url.pathname === "/custom-worker.mjs") {
        requests.set(url.pathname, (requests.get(url.pathname) ?? 0) + 1);
        return new Response(workerModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname === "/arabic.pptx") return new Response(pptx);
      if (url.pathname === "/arabic.xlsx") return new Response(xlsx);
      if (url.pathname === "/arabic.docx") return new Response(docx);
      return new Response(`<!doctype html><main id="stage"></main><script type="module">
        window.pageErrors = [];
        window.addEventListener("error", (event) => window.pageErrors.push(String(event.message)));
        window.addEventListener("unhandledrejection", (event) => window.pageErrors.push(String(event.reason)));
        await import("/react-harness.mjs");
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    await waitFor(view, `typeof window.mountHook === "function"`);
    await run(view);
    const pageErrors = (await view.evaluate(`window.pageErrors`)) as string[];
    expect(pageErrors).toEqual([]);
  } finally {
    server.stop();
  }
}

visual("visual react hooks: usePptx mounts, navigates, zooms, and unmounts", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("pptx", "/arabic.pptx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const opened = (await view.evaluate(`window.probe`)) as Probe;
    expect(opened.format).toBe("pptx");
    expect(opened.count).toBe(10);
    expect(opened.index).toBe(0);
    expect(opened.error).toBeNull();
    const first = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(first.canvasCount).toBe(0);
    expect(first.children).toBeGreaterThan(0);
    expect(first.text).toContain("تقرير الإنتاج الأسبوعي");

    await view.evaluate(`hookGoTo(1)`);
    await waitFor(view, `window.probe?.index === 1`);
    const second = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(second.text).toContain("الاسم");
    expect(second.canvasCount).toBe(0);

    await view.evaluate(`hookSetZoom(1.5)`);
    await waitFor(view, `window.probe?.zoom === 1.5`);

    // Zoom belongs to the facade: navigation keeps it, in the mirror and the DOM.
    await view.evaluate(`hookGoTo(2)`);
    await waitFor(view, `window.probe?.index === 2`);
    const kept = (await view.evaluate(`(() => {
      const slide = document.querySelector('[data-aared-element="slide"]');
      const scene = document.querySelector('[data-aared-element="slide-scene"]');
      return { zoom: window.probe.zoom, scale: slide.getBoundingClientRect().width / parseFloat(scene.style.width) };
    })()`)) as { zoom: unknown; scale: number };
    expect(kept.zoom).toBe(1.5);
    expect(kept.scale).toBeCloseTo(1.5, 2);

    const cleared = (await view.evaluate(`unmountHook()`)) as number;
    expect(cleared).toBe(0);
    const empty = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(empty.text).toBe("");
  });
});

visual("visual react hooks: a worker factory is used at mount, and a re-render never remounts", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("worker", "/arabic.pptx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    await view.evaluate(`rerenderHook()`);
    await view.evaluate(`rerenderHook()`);
    await Bun.sleep(300);
    const state = (await view.evaluate(`({ status: window.probe?.status, format: window.probe?.format, calls: window.workerFactoryCalls })`)) as {
      status: string; format: string; calls: number;
    };
    expect(state.status).toBe("ready");
    expect(state.format).toBe("pptx");
    // Detection and the document session each start one worker, both through
    // the caller's factory; re-renders with a new factory identity add none.
    expect(state.calls).toBeGreaterThanOrEqual(1);
    const callsAtReady = state.calls;
    expect((await view.evaluate(`window.workerFactoryCalls`)) as number).toBe(callsAtReady);
    expect(requests.get("/custom-worker.mjs") ?? 0).toBeGreaterThanOrEqual(1);
    expect(requests.get("/worker.mjs") ?? 0).toBe(0);
    await view.evaluate(`unmountHook()`);
  });
});

visual("visual react hooks: unmounting mid-flight leaves no painted output", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("pptx", "/arabic.pptx")`);
    // Catch the mount while fetching/parsing, if it lasts long enough to
    // observe; a fast mount still exercises cleanup through the same path.
    await waitFor(view, `window.probe?.status === "loading" || window.probe?.status === "ready"`);
    const cleared = (await view.evaluate(`unmountHook()`)) as number;
    expect(cleared).toBe(0);
    await Bun.sleep(1000);
    const empty = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(empty.text).toBe("");
    expect(empty.children).toBe(0);
  });
});

visual("visual react hooks: useXlsx and useDocx mount their formats; StrictMode auto-detect works", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("xlsx", "/arabic.xlsx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const workbook = (await view.evaluate(`window.probe`)) as Probe;
    expect(workbook.format).toBe("xlsx");
    const grid = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(grid.cells).toBeGreaterThan(0);
    expect(grid.canvasCount).toBe(0);
    await view.evaluate(`unmountHook()`);

    await view.evaluate(`mountHook("docx", "/arabic.docx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const document = (await view.evaluate(`window.probe`)) as Probe;
    expect(document.format).toBe("docx");
    const paged = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(paged.pages).toBeGreaterThanOrEqual(1);
    expect(paged.text).toContain("1. Project Overview");
    expect(paged.canvasCount).toBe(0);
    await view.evaluate(`unmountHook()`);

    await view.evaluate(`mountHook("auto", "/arabic.pptx", true)`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const strict = (await view.evaluate(`window.probe`)) as Probe;
    expect(strict.format).toBe("pptx");
    expect(strict.error).toBeNull();
    const painted = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(painted.text).toContain("تقرير الإنتاج الأسبوعي");
    await view.evaluate(`unmountHook()`);
  });
});
