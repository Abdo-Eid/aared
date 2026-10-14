import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

visual("a PPTX or DOCX render handle's setAppearance() changes only what it sets", async () => {
  const formats = ["pptx", "docx"] as const;
  const entries = new Map(await Promise.all(formats.map(async (format) => [format,
    await browserModule(`${import.meta.dir}/../src/${format}/index.ts`)] as const)));
  const fixtures = new Map(await Promise.all(formats.map(async (format) => [format,
    await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer()] as const)));
  const worker = await browserModule(`${import.meta.dir}/../src/core/worker.ts`);
  const modules: Record<string, string> = { "/worker.mjs": worker };
  const files: Record<string, ArrayBuffer> = {};
  for (const format of formats) {
    modules[`/${format}.mjs`] = entries.get(format)!;
    files[`/fixture.${format}`] = fixtures.get(format)!;
  }
  await withVisualPage({
    modules,
    files,
    page: `<!doctype html><main id="host" style="width:900px;height:500px"></main><script type="module">
      const host = document.querySelector('#host');
      // Re-read the theme marker the render painted, not the handle's state.
      const mark = () => host.querySelector('[data-aared-theme]')?.dataset.aaredTheme ?? null;
      window.run = async (format) => {
        const { open, render } = await import('/' + format + '.mjs');
        const doc = await open('/fixture.' + format);
        try {
          const input = format === 'pptx' ? { slide: doc.slides.get(0) } : { document: doc };
          const handle = await render({ ...input, target: host, appearance: { theme: 'dark' } });
          const start = mark();
          handle.setAppearance({});
          const afterEmpty = mark();
          handle.setAppearance({ spreadsheet: { headers: false } });
          const afterOther = mark();
          handle.setAppearance({ theme: 'light' });
          const afterLight = mark();
          handle.destroy();
          return { start, afterEmpty, afterOther, afterLight };
        } finally { doc.close(); }
      };
    </script>`,
    width: 1000,
    height: 700,
  }, async (view) => {
    for (const format of formats) {
      // A theme left out stays (it used to go back to light); passing one changes it.
      expect(await view.evaluate(`run(${JSON.stringify(format)})`) as Record<string, string | null>, format).toEqual({
        start: "dark", afterEmpty: "dark", afterOther: "dark", afterLight: "light",
      });
    }
  });
}, 60000);
