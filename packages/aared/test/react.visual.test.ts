import { expect, test } from "bun:test";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;

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

async function browserModule(entrypoint: string): Promise<string> {
  const build = await Bun.build({ entrypoints: [entrypoint], target: "browser", format: "esm" });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  return build.outputs[0]!.text();
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

async function withHarness(run: (view: {
  evaluate: (expression: string) => Promise<unknown>;
}) => Promise<void>): Promise<void> {
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
      if (url.pathname === "/worker.mjs") {
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
    expect(opened.count).toBe(4);
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

    const cleared = (await view.evaluate(`unmountHook()`)) as number;
    expect(cleared).toBe(0);
    const empty = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(empty.text).toBe("");
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
