import { expect } from "bun:test";
import { join } from "node:path";
import { visual, withVisualPage, within } from "./harness.js";

visual("visual shipped ESM defers preview adapters and loads XLSX effects only when painted", async () => {
  const requests: string[] = [];
  let rejectEffects = false;
  let effectRequested: (() => void) | undefined;
  let effectGate: Promise<void> | undefined;
  const server = Bun.serve({ port: 0, async fetch(request) {
    const path = new URL(request.url).pathname;
    requests.push(path);
    if (/\/effect-paint-/.test(path)) {
      effectRequested?.();
      if (rejectEffects) return new Response("painter unavailable", { status: 503 });
      await effectGate;
    }
    if (path.startsWith("/dist/")) return new Response(Bun.file(join(import.meta.dir, "..", path.slice(1))));
    if (path.startsWith("/fixture.")) {
      const format = path.slice("/fixture.".length);
      return new Response(Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`));
    }
    return new Response(`<!doctype html><main id="host" style="width:1100px;height:650px"></main><script>
      window.importFormat = async format => { window.api = await import('/dist/' + format + '/index.mjs'); };
      window.openSheet = async title => {
        window.doc ??= await api.open('/fixture.xlsx');
        window.paint?.destroy();
        window.paint = await api.render({ sheet: doc.sheets.get(title), target: document.querySelector('#host') });
        const picture = document.querySelector('[data-aared-element="picture"]');
        return { text: document.querySelector('#host').textContent,
          transform: picture && getComputedStyle(picture).transform,
          filter: picture && getComputedStyle(picture).filter,
          fallback: picture?.dataset.aaredEffectFallback };
      };
      window.previewFormat = async format => {
        window.paint?.destroy(); window.doc?.close();
        const viewer = await api.preview({ source: '/fixture.' + format, target: document.querySelector('#host') });
        const result = { format: viewer.format, text: document.querySelector('#host').textContent,
          canvases: document.querySelectorAll('canvas').length };
        viewer.destroy();
        return { ...result, empty: document.querySelector('#host').childElementCount === 0 };
      };
    </script>`, { headers: { "content-type": "text/html" } });
  } });
  try {
    for (const format of ["pptx", "xlsx", "docx"]) {
      requests.length = 0;
      // One server for the whole test: the request log and the effect-painter
      // gate are per-server state the loop mutates, so it must outlive a page.
      await withVisualPage({ server, keepServer: true, width: 1200, height: 800 }, async (view) => {
      await within("importing the shipped format entry", view.evaluate(`importFormat(${JSON.stringify(format)})`), 30_000);
      expect(requests.some((path) => /\/(?:embed-viewer|auto-viewer|viewer)-/.test(path))).toBe(false);
      expect(requests.some((path) => /\/effect-paint-/.test(path))).toBe(format === "pptx");
      for (const other of ["pptx", "xlsx", "docx"].filter((other) => other !== format)) {
        expect(requests.some((path) => path.includes(`/core/${other}-`))).toBe(false);
      }
      if (format === "xlsx") {
        const plain = await within("painting a plain sheet", view.evaluate(`openSheet('Summary')`), 30_000) as { text: string };
        expect(plain.text).toContain("Wireless Charger");
        expect(requests.some((path) => /\/effect-paint-/.test(path))).toBe(false);
        let releaseEffect!: () => void;
        effectGate = new Promise<void>((resolve) => { releaseEffect = resolve; });
        const requested = new Promise<void>((resolve) => { effectRequested = resolve; });
        await view.evaluate(`(() => { window.painting = openSheet('Notes'); return true; })()`);
        try {
          await within("requesting the effect painter", requested, 30_000);
          const pending = await view.evaluate(`(() => {
            const picture = document.querySelector('[data-aared-element="picture"]');
            return { present: !!picture, fallback: picture?.dataset.aaredEffectFallback ?? null };
          })()`);
          expect(pending).toEqual({ present: true, fallback: null });
        } finally { releaseEffect(); effectRequested = undefined; effectGate = undefined; }
        const effected = await within("painting the native SVG picture and shadow", view.evaluate(`window.painting`), 30_000) as {
          text: string; transform: string; filter: string; fallback?: string;
        };
        expect(effected.transform).not.toBe("none");
        expect(effected.filter).toContain("drop-shadow");
        expect(effected.fallback).toBeUndefined();
        expect(requests.filter((path) => /\/effect-paint-/.test(path))).toHaveLength(1);
      }
      const result = await within("opening the lazy format preview", view.evaluate(`previewFormat(${JSON.stringify(format)})`), 30_000) as {
        format: string; text: string; canvases: number; empty: boolean;
      };
      expect(result.format).toBe(format);
      expect(result.text.length).toBeGreaterThan(0);
      expect(result.canvases).toBe(0);
      expect(result.empty).toBe(true);
      expect(requests.some((path) => /\/embed-viewer-/.test(path))).toBe(true);
      expect(requests.some((path) => /\/viewer-/.test(path))).toBe(true);
      });
    }
    rejectEffects = true;
    await withVisualPage({ server, keepServer: true, width: 1200, height: 800 }, async (view) => {
      await view.evaluate(`importFormat('xlsx')`);
      const failed = await within("reporting a failed effect painter", view.evaluate(`openSheet('Notes')`), 30_000) as {
        text: string; transform: string; filter: string; fallback: string;
      };
      expect(failed.text).toContain("Studio handbook");
      expect(failed.transform).not.toBe("none");
      expect(failed.filter).toBe("none");
      expect(failed.fallback).toBe("painter-unavailable");
    });
  } finally { server.stop(); }
}, 120_000);
