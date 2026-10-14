import { expect, test } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

interface Probe {
  status: string;
  format: string | null;
  count: number | null;
  index: number | null;
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
  await withVisualPage({
    modules: {
      "/react-harness.mjs": harnessModule,
      "/worker.mjs": workerModule,
      "/custom-worker.mjs": workerModule,
    },
    files: { "/arabic.pptx": pptx, "/slow.pptx": pptx, "/arabic.xlsx": xlsx, "/arabic.docx": docx },
    page: `<!doctype html><main id="stage"></main><script type="module">
        window.pageErrors = [];
        window.addEventListener("error", (event) => window.pageErrors.push(String(event.message)));
        window.addEventListener("unhandledrejection", (event) => window.pageErrors.push(String(event.reason)));
        await import("/react-harness.mjs");
      </script>`,
    onRequest: async (url) => {
      if (url.pathname === "/worker.mjs" || url.pathname === "/custom-worker.mjs") {
        requests.set(url.pathname, (requests.get(url.pathname) ?? 0) + 1);
      } else if (url.pathname === "/slow.pptx") {
        await Bun.sleep(300);
      }
      return undefined;
    },
    width: 1280,
    height: 720,
  }, async (view) => {
    await waitFor(view, `typeof window.mountHook === "function"`);
    await run(view);
    const pageErrors = (await view.evaluate(`window.pageErrors`)) as string[];
    expect(pageErrors).toEqual([]);
  });
}

visual("visual react hooks: usePptx mounts, navigates and zooms through its viewer, and unmounts", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("pptx", "/arabic.pptx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const opened = (await view.evaluate(`window.probe`)) as Probe;
    expect(opened.format).toBe("pptx");
    expect(opened.count).toBe(20);
    expect(opened.index).toBe(0);
    expect(opened.error).toBeNull();
    const first = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(first.canvasCount).toBe(0);
    expect(first.children).toBeGreaterThan(0);
    expect(first.text).toContain("تقرير الإنتاج الأسبوعي");

    await view.evaluate(`viewerGoTo(1)`);
    await waitFor(view, `window.probe?.index === 1`);
    const second = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(second.text).toContain("الاسم");
    expect(second.canvasCount).toBe(0);

    await view.evaluate(`viewerSetZoom(1.5)`);
    await waitFor(view, `window.probe?.zoom === 1.5`);

    // Zoom belongs to the facade: navigation keeps it, in the mirror and the DOM.
    await view.evaluate(`viewerGoTo(2)`);
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

visual("visual react hooks: a failed public call keeps the page and rejects only to its caller", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("pptx", "/arabic.pptx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    expect(((await view.evaluate(`window.probe`)) as Probe).error).toBeNull();

    // A unit that cannot be shown rejects to the app without changing hook state.
    const failed = (await view.evaluate(`viewerGoToFailure(99)`)) as { code: unknown; message: string } | null;
    expect(failed?.code).toBe("not-found");
    const after = (await view.evaluate(`snapshot()`)) as Snapshot;
    expect(after.probe?.status).toBe("ready");
    expect(after.probe?.error).toBeNull();
    expect(after.probe?.index).toBe(0);
    expect(after.text).toContain("تقرير الإنتاج الأسبوعي");

    // The next successful move still works.
    await view.evaluate(`viewerGoTo(1)`);
    await waitFor(view, `window.probe?.index === 1 && window.probe?.error === null`);
    await view.evaluate(`unmountHook()`);
  });
});

visual("visual react hooks: a failed built-in sheet tab lands in error while keeping the view ready", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`(() => {
      const NativeWorker = window.Worker;
      window.Worker = class extends NativeWorker {
        constructor(...args) { super(...args); window.testWorker = this; }
      };
    })()`);
    await view.evaluate(`mountHook("xlsx", "/arabic.xlsx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const before = await view.evaluate(`snapshot()`) as Snapshot;
    await view.evaluate(`(() => {
      window.testWorker.terminate();
      window.testWorker.dispatchEvent(new ErrorEvent('error', {message:'test Worker failure'}));
      document.querySelectorAll('[data-aared-element="sheet-tab"]')[1].click();
    })()`);
    await waitFor(view, `window.probe?.error !== null`);
    const after = await view.evaluate(`snapshot()`) as Snapshot;
    expect(after.probe?.status).toBe("ready");
    expect(after.probe?.index).toBe(0);
    expect(after.probe?.error).toContain("test Worker failure");
    expect(after.text).toBe(before.text);
    expect(after.cells).toBe(before.cells);
    await view.evaluate(`unmountHook()`);
  });
});

visual("visual react hooks: borrowed content has no lifecycle and ends at unmount for every format", async () => {
  await withHarness(async (view) => {
    for (const format of ["pptx", "xlsx", "docx"]) {
      await view.evaluate(`mountHook(${JSON.stringify(format)}, '/arabic.${format}')`);
      await waitFor(view, `window.probe?.status === "ready"`);
      expect(await view.evaluate(`(() => {
        const viewer = window.hookViewer;
        const content = viewer.content;
        window.readContent = viewer.format === 'pptx' ? content.slides.get(0).read
          : viewer.format === 'xlsx' ? content.sheets.get(0).read : content.sections.list;
        return !('document' in viewer) && !('close' in content) && !('search' in content);
      })()`)).toBe(true);
      await view.evaluate(`unmountHook()`);
      expect(await view.evaluate(`Promise.resolve().then(() => window.readContent()).then(() => 'resolved', error => error.code)`)).toBe("closed");
    }
  });
});

visual("visual react hooks: index and zoom are start values; appearance updates without remounting", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("start", "/arabic.pptx")`);
    await waitFor(view, `window.probe?.status === "ready"`);
    const opened = (await view.evaluate(`({ probe: window.probe, seen: window.seen })`)) as { probe: Probe; seen: { viewers: number } };
    expect(opened.probe).toMatchObject({ format: "pptx", index: 0, zoom: "fit-page" });
    expect(opened.seen.viewers).toBe(1);
    const scale = `(() => {
      const slide = document.querySelector('[data-aared-element="slide"]');
      const scene = document.querySelector('[data-aared-element="slide-scene"]');
      return slide.getBoundingClientRect().width / parseFloat(scene.style.width);
    })()`;
    const theme = `document.querySelector('[data-aared-element="slide"]')?.dataset.aaredTheme`;
    const fitScale = (await view.evaluate(scale)) as number;
    expect(await view.evaluate(theme)).toBe("light");

    // Index/zoom remain start values; appearance changes on the same viewer.
    await view.evaluate(`setStartProps({ zoom: 1.5, index: 2, theme: "dark" })`);
    await Bun.sleep(300);
    const kept = (await view.evaluate(`({ probe: window.probe, seen: window.seen })`)) as { probe: Probe; seen: { viewers: number } };
    expect(kept.probe).toMatchObject({ status: "ready", index: 0, zoom: "fit-page" });
    expect(kept.seen.viewers).toBe(1);
    expect((await view.evaluate(scale)) as number).toBeCloseTo(fitScale, 2);
    expect(await view.evaluate(theme)).toBe("dark");

    // The viewer's methods are what change it, and the result mirrors them.
    await view.evaluate(`viewerSetZoom(1.5)`);
    await waitFor(view, `window.probe?.zoom === 1.5`);
    expect((await view.evaluate(scale)) as number).toBeCloseTo(1.5, 2);

    // retry() mounts again, starting from the latest start values.
    await view.evaluate(`hookRetry()`);
    await waitFor(view, `window.seen?.viewers === 2 && window.probe?.status === "ready"`);
    expect(((await view.evaluate(`window.probe`)) as Probe)).toMatchObject({ index: 2, zoom: 1.5 });
    expect(await view.evaluate(theme)).toBe("dark");

    // An inline media resolver: adding one re-opens once (the media route changes),
    // then re-renders with a new function each time re-open nothing, and an
    // unchanged result keeps its identity.
    await view.evaluate(`setStartProps({ zoom: 1.5, index: 2, theme: "dark", inlineMedia: true })`);
    await waitFor(view, `window.seen?.viewers === 3 && window.probe?.status === "ready"`);
    expect(((await view.evaluate(`window.probe`)) as Probe).index).toBe(2);
    const before = (await view.evaluate(`window.seen`)) as { viewers: number; results: number };
    await view.evaluate(`setStartProps({ zoom: 1.5, index: 2, theme: "dark", inlineMedia: true })`);
    await view.evaluate(`setStartProps({ zoom: 1.5, index: 2, theme: "dark", inlineMedia: true })`);
    await Bun.sleep(300);
    expect((await view.evaluate(`window.seen`)) as { viewers: number; results: number }).toEqual(before);
    await view.evaluate(`unmountHook()`);
  });
});

visual("visual react hooks: appearance follows values across formats and preserves viewer state", async () => {
  await withHarness(async (view) => {
    for (const format of ["pptx", "xlsx", "docx"]) {
      await view.evaluate(`mountHook("start", "/arabic.${format}", true)`);
      await waitFor(view, `window.probe?.status === "ready"`);
      if (format !== "docx") await view.evaluate(`viewerGoTo(1)`);
      await view.evaluate(`viewerSetZoom(1.25)`);
      if (format === "xlsx") await view.evaluate(`hookViewer.xlsx.goToCell("B2")`);
      await view.evaluate(`(() => { window.savedViewer = hookViewer; window.savedSeen = seen.viewers; })()`);
      const workersBefore = requests.get("/worker.mjs") ?? 0;
      await view.evaluate(`setStartProps({ appearance: { theme: "dark", spreadsheet: { sheetTabs: false, headers: false } } })`);
      await waitFor(view, `document.querySelector('[data-aared-theme]')?.dataset.aaredTheme === "dark"`);
      expect(await view.evaluate(`hookViewer === savedViewer && seen.viewers === savedSeen`)).toBe(true);
      expect(await view.evaluate(`probe.index`)).toBe(format === "docx" ? 0 : 1);
      expect(await view.evaluate(`probe.zoom`)).toBe(1.25);
      if (format === "xlsx") {
        expect(await view.evaluate(`hookViewer.xlsx.selection?.address`)).toBe("B2");
        expect(await view.evaluate(`document.querySelector('[data-aared-element="sheet-tabs"]') === null`)).toBe(true);
        expect(await view.evaluate(`document.querySelector('[data-aared-element="column-headers"]') === null`)).toBe(true);
        // The chrome flag flows through the same key: toggling it paints
        // and removes the bar on the same viewer, never remounting.
        await view.evaluate(`setStartProps({ appearance: { theme: "dark", spreadsheet: { sheetTabs: false, headers: false, formulaBar: true } } })`);
        await waitFor(view, `document.querySelector('[data-aared-element="sheet-formula-bar"]') !== null`);
        expect(await view.evaluate(`hookViewer === savedViewer && seen.viewers === savedSeen`)).toBe(true);
        await view.evaluate(`setStartProps({ appearance: { theme: "dark", spreadsheet: { sheetTabs: false, headers: false } } })`);
        await waitFor(view, `document.querySelector('[data-aared-element="sheet-formula-bar"]') === null`);
        expect(await view.evaluate(`hookViewer === savedViewer && seen.viewers === savedSeen`)).toBe(true);
      }
      // Unchanged values, even with different property order, must not undo
      // a consumer's imperative appearance change on an unrelated render.
      await view.evaluate(`hookViewer.setAppearance({ theme: "light" })`);
      await view.evaluate(`setStartProps({ appearance: { spreadsheet: { headers: false, sheetTabs: false }, theme: "dark" } })`);
      await Bun.sleep(100);
      expect(await view.evaluate(`document.querySelector('[data-aared-theme]')?.dataset.aaredTheme`)).toBe("light");
      await view.evaluate(`setStartProps({ appearance: { theme: "dark" } })`);
      await waitFor(view, `document.querySelector('[data-aared-theme]')?.dataset.aaredTheme === "dark"`);
      await view.evaluate(`setStartProps({})`);
      await waitFor(view, `document.querySelector('[data-aared-theme]')?.dataset.aaredTheme === "light"`);
      if (format === "xlsx") {
        expect(await view.evaluate(`document.querySelector('[data-aared-element="sheet-tabs"]') !== null`)).toBe(true);
        expect(await view.evaluate(`document.querySelector('[data-aared-element="column-headers"]') !== null`)).toBe(true);
      }
      expect(await view.evaluate(`hookViewer === savedViewer && seen.viewers === savedSeen`)).toBe(true);
      expect(requests.get("/worker.mjs") ?? 0).toBe(workersBefore);
      await view.evaluate(`unmountHook()`);
    }
  });
});

visual("visual react hooks: appearance changed while loading applies to the ready viewer", async () => {
  await withHarness(async (view) => {
    await view.evaluate(`mountHook("start", "/slow.pptx")`);
    await waitFor(view, `probe?.status === "loading"`);
    await view.evaluate(`setStartProps({ theme: "dark" })`);
    await waitFor(view, `probe?.status === "ready" && document.querySelector('[data-aared-theme]')?.dataset.aaredTheme === "dark"`);
    expect(await view.evaluate(`seen.viewers`)).toBe(1);
    await view.evaluate(`unmountHook()`);
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

visual("visual react hooks: unmounting the moment it is ready leaves nothing running", async () => {
  await withHarness(async (view) => {
    // Unmount in the same tick the result turns ready, so whatever the viewer
    // or the host started on that render is still in flight when the document
    // closes. Polling from the test would only land there by chance; a
    // rejection nobody handles shows up in the page errors `withHarness` checks.
    for (let round = 0; round < 3; round++) {
      await view.evaluate(`new Promise((done) => {
        mountHook("pptx", "/arabic.pptx");
        const tick = () => window.probe?.status === "ready" ? done(unmountHook()) : setTimeout(tick, 0);
        tick();
      })`);
    }
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
