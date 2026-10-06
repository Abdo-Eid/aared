import { expect, test } from "bun:test";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURES = `${import.meta.dir}/../../../fixtures`;

// A page whose OS colour scheme the test controls: `matchMedia` is replaced
// before any module loads, `flip("dark")` changes the answer and fires every
// listener, and `listeners()` says how many are still attached (a viewer that
// follows the OS must drop its listener when it stops, or the page leaks one).
const FAKE_OS = `
  window.scheme = "light";
  const subscribed = new Set();
  window.matchMedia = (query) => ({
    media: query,
    get matches() { return query === "(prefers-color-scheme: dark)" && window.scheme === "dark"; },
    addEventListener: (type, listener) => { if (type === "change") subscribed.add(listener); },
    removeEventListener: (type, listener) => { subscribed.delete(listener); },
  });
  window.listeners = () => subscribed.size;
  window.flip = (scheme) => {
    window.scheme = scheme;
    for (const listener of [...subscribed]) listener({ matches: scheme === "dark", media: "(prefers-color-scheme: dark)" });
  };
  window.frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
`;

/** Every `data-aared-theme` a renderer or the element wrote under `root` (and on it). */
const MARKS = `
  window.marks = (root) => [root, ...root.querySelectorAll("[data-aared-theme]")]
    .filter((element) => element.dataset?.aaredTheme !== undefined)
    .map((element) => element.dataset.aaredTheme);
`;

async function serve(entry: "preview" | "element", body: string, script: string) {
  const [module, worker, pptx, xlsx, docx] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/${entry}/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
    Bun.file(`${FIXTURES}/pptx-coverage.pptx`).arrayBuffer(),
    Bun.file(`${FIXTURES}/xlsx-coverage.xlsx`).arrayBuffer(),
    Bun.file(`${FIXTURES}/docx-coverage.docx`).arrayBuffer(),
  ]);
  const js = { headers: { "content-type": "text/javascript; charset=utf-8" } };
  return Bun.serve({
    port: 0,
    fetch(request) {
      const { pathname } = new URL(request.url);
      if (pathname === `/${entry}.mjs`) return new Response(module, js);
      if (pathname === "/worker.mjs") return new Response(worker, js);
      if (pathname === "/a.pptx") return new Response(pptx);
      if (pathname === "/a.xlsx") return new Response(xlsx);
      if (pathname === "/a.docx") return new Response(docx);
      return new Response(
        `<!doctype html><body style="margin:0"><script>${FAKE_OS}</script>${body}<script type="module">${MARKS}${script}</script>`,
        { headers: { "content-type": "text/html; charset=utf-8" } },
      );
    },
  });
}

interface Reading { marks: string[]; index: number; zoom: unknown }

visual("the system theme follows the OS in preview() for every format, and an explicit theme never does", async () => {
  const server = await serve(
    "preview",
    `<main id="host" style="height:600px"></main>`,
    `import { preview } from "/preview.mjs";
     window.openDoc = async (source, theme) => {
       window.handle = await preview({ source, target: document.querySelector("#host"), appearance: { theme } });
     };
     window.read = () => ({ marks: marks(document.querySelector("#host")), index: handle.index, zoom: handle.zoom });`,
  );
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    const read = async (): Promise<Reading> => await view.evaluate(`frame().then(() => read())`) as Reading;
    const flip = async (scheme: string): Promise<void> => void await view.evaluate(`flip(${JSON.stringify(scheme)})`);
    /** Statements, run in the page (an evaluate takes one expression). */
    const run = async (code: string): Promise<void> => void await view.evaluate(`(async () => { ${code} })()`);
    /** How many OS-scheme listeners the page still has attached. */
    const attached = async (): Promise<number> => await view.evaluate(`listeners()`) as number;

    for (const format of ["pptx", "xlsx", "docx"]) {
      // `system`: follows the OS now and on every flip, with the unit and zoom left alone.
      await run(`window.scheme = "light"; await openDoc("/a.${format}", "system");`);
      const start = await read();
      expect(start.marks.length, format).toBeGreaterThan(0);
      expect(new Set(start.marks), format).toEqual(new Set(["light"]));
      expect(await attached(), format).toBe(1);

      await flip("dark");
      const dark = await read();
      expect(new Set(dark.marks), format).toEqual(new Set(["dark"]));
      expect({ index: dark.index, zoom: dark.zoom }, format).toEqual({ index: start.index, zoom: start.zoom });
      await flip("light");
      expect(new Set((await read()).marks), format).toEqual(new Set(["light"]));

      // The host takes the theme over: the OS stops mattering, and the listener is dropped.
      await view.evaluate(`handle.setAppearance({ theme: "dark" })`);
      expect(await attached(), format).toBe(0);
      await flip("light");
      expect(new Set((await read()).marks), format).toEqual(new Set(["dark"]));
      await flip("dark");

      // Asking for `system` again resolves against the OS as it is now, and listens again.
      await run(`window.scheme = "light"; handle.setAppearance({ theme: "system" });`);
      expect(new Set((await read()).marks), format).toEqual(new Set(["light"]));
      expect(await attached(), format).toBe(1);

      // Destroying the viewer drops the listener too.
      await view.evaluate(`handle.destroy()`);
      expect(await attached(), format).toBe(0);
    }

    // An explicit theme, light or dark, is never moved by the OS and never listens.
    for (const theme of ["light", "dark"]) {
      await run(`window.scheme = ${JSON.stringify(theme === "light" ? "dark" : "light")}; await openDoc("/a.pptx", ${JSON.stringify(theme)});`);
      expect(await attached(), theme).toBe(0);
      await flip(theme === "light" ? "light" : "dark");
      expect(new Set((await read()).marks), theme).toEqual(new Set([theme]));
      await view.evaluate(`handle.destroy()`);
    }
  } finally {
    server.stop(true);
  }
});

visual("<aared-viewer theme=system> follows the OS in its content and its chrome, and stops when the theme is set", async () => {
  const server = await serve(
    "element",
    `<aared-viewer id="viewer" style="height:600px"></aared-viewer>`,
    `import "/element.mjs";
     const viewer = document.querySelector("#viewer");
     window.openSource = (source) => new Promise((resolve, reject) => {
       viewer.addEventListener("load", () => resolve(), { once: true });
       viewer.addEventListener("error", (event) => reject(new Error(event.detail.message)), { once: true });
       viewer.src = source;
     });
     window.read = () => ({
       chrome: viewer.dataset.aaredTheme,
       content: marks(viewer.shadowRoot),
       index: viewer.viewer?.index,
       zoom: viewer.viewer?.zoom,
     });`,
  );
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    type ElementReading = { chrome: string; content: string[]; index: number; zoom: unknown };
    const read = async (): Promise<ElementReading> => await view.evaluate(`frame().then(() => read())`) as ElementReading;
    const run = async (code: string): Promise<void> => void await view.evaluate(`(async () => { ${code} })()`);
    const attached = async (): Promise<number> => await view.evaluate(`listeners()`) as number;

    // The element's own chrome listener is attached from the moment it connects: one listener.
    expect(await attached()).toBe(1);

    await run(`document.querySelector("#viewer").theme = "system"; await openSource("/a.pptx");`);
    const start = await read();
    expect(start.chrome).toBe("light");
    expect(new Set(start.content)).toEqual(new Set(["light"]));
    // The element's chrome listener, and the viewer's own.
    expect(await attached()).toBe(2);

    await view.evaluate(`flip("dark")`);
    const dark = await read();
    expect(dark.chrome).toBe("dark");
    expect(new Set(dark.content)).toEqual(new Set(["dark"]));
    expect({ index: dark.index, zoom: dark.zoom }).toEqual({ index: start.index, zoom: start.zoom });

    // The host sets the theme: the viewer stops following (the chrome listener stays while connected).
    await view.evaluate(`document.querySelector("#viewer").theme = "light"`);
    expect(await attached()).toBe(1);
    await view.evaluate(`flip("dark")`);
    const fixed = await read();
    expect(fixed.chrome).toBe("light");
    expect(new Set(fixed.content)).toEqual(new Set(["light"]));

    // `system` again: both follow, from the OS as it is now.
    await view.evaluate(`document.querySelector("#viewer").theme = "system"`);
    expect((await read()).chrome).toBe("dark");
    expect(new Set((await read()).content)).toEqual(new Set(["dark"]));
    expect(await attached()).toBe(2);

    // Removing the element drops both.
    await view.evaluate(`document.querySelector("#viewer").remove()`);
    expect(await attached()).toBe(0);
  } finally {
    server.stop(true);
  }
});
