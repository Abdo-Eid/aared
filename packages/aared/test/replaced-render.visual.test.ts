import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

const FIXTURES = {
  pptx: `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`,
  xlsx: `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`,
  docx: `${import.meta.dir}/../../../fixtures/docx-coverage.docx`,
};

/**
 * A render that a newer render into the same target replaced tears itself
 * down in every format (`onReplaced`): it reads `destroyed`, and stops
 * observing, refitting and holding media URLs without the caller tracking it.
 * The newer render keeps working, and the target gets its own layout back
 * once the last XLSX render on it is destroyed. A render whose `signal` is
 * aborted before it paints never touches the target, so a stale render
 * cannot replace a newer one in the first place; one aborted after it painted
 * but before it resolved takes its own nodes back out. Either way an aborted
 * render never hands back a handle.
 */
visual("visual DOM output: a replaced render destroys itself and a cancelled one paints nothing, in every format", async () => {
  const [pptx, xlsx, docx, worker, deck, book, doc] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/pptx/index.ts`),
    browserModule(`${import.meta.dir}/../src/xlsx/index.ts`),
    browserModule(`${import.meta.dir}/../src/docx/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
    Bun.file(FIXTURES.pptx).arrayBuffer(),
    Bun.file(FIXTURES.xlsx).arrayBuffer(),
    Bun.file(FIXTURES.docx).arrayBuffer(),
  ]);
  const scripts: Record<string, string> = { "/pptx.mjs": pptx, "/xlsx.mjs": xlsx, "/docx.mjs": docx, "/worker.mjs": worker };
  const files: Record<string, ArrayBuffer> = { "/deck.pptx": deck, "/book.xlsx": book, "/doc.docx": doc };
  await withVisualPage({
    modules: scripts,
    files,
    page: `<!doctype html><body style="margin:0"><main id="host" style="width:900px;height:600px"></main><script type="module">
        import * as pptx from "/pptx.mjs";
        import * as xlsx from "/xlsx.mjs";
        import * as docx from "/docx.mjs";
        const host = document.querySelector("#host");
        const settle = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        window.replace = async () => {
          const deck = await pptx.open("/deck.pptx");
          const book = await xlsx.open("/book.xlsx");
          const doc = await docx.open("/doc.docx");
          const out = {};
          // Each render goes into the same target with no destroy() in between.
          const slideA = await pptx.render({ slide: deck.slides.get(0), target: host });
          const slideB = await pptx.render({ slide: deck.slides.get(1), target: host });
          await settle();
          out.pptx = [slideA.destroyed, slideB.destroyed];
          const flow = await docx.render({ document: doc, target: host });
          await settle();
          out.pptxByDocx = slideB.destroyed;
          const paged = await docx.render({ document: doc, target: host, pagination: { mode: "layout" } });
          await settle();
          out.docx = [flow.destroyed, paged.destroyed];
          // The replaced flow render no longer moves anything.
          await flow.goToHeading(0);
          const sheetA = await xlsx.render({ sheet: book.sheets.get(0), target: host });
          await settle();
          out.docxByXlsx = paged.destroyed;
          const sheetB = await xlsx.render({ sheet: book.sheets.get(1), target: host, appearance: { spreadsheet: { sheetTabs: true } } });
          await settle();
          out.xlsx = [sheetA.destroyed, sheetB.destroyed];
          // The newer sheet still owns the target's layout and still zooms.
          out.display = host.style.display;
          sheetB.setZoom(2);
          out.zoomed = sheetB.scale;
          // Switching sheets inside one render (the tab strip) is not a replacement.
          await sheetB.goToSheet(0);
          await settle();
          out.switched = sheetB.destroyed;
          // Another format over the sheet: the sheet goes, and the target gets its own layout back.
          const last = await docx.render({ document: doc, target: host });
          await settle();
          out.xlsxByDocx = sheetB.destroyed;
          out.restored = host.style.display;
          last.destroy();
          deck.close();
          book.close();
          doc.close();
          return out;
        };
        // A render cancelled before it paints leaves the target as it was and
        // rejects aborted; one cancelled after it painted is simply on screen.
        window.cancel = async () => {
          const deck = await pptx.open("/deck.pptx");
          const book = await xlsx.open("/book.xlsx");
          const doc = await docx.open("/doc.docx");
          const out = {};
          const keep = document.createElement("p");
          keep.id = "kept";
          const attempt = async (start) => {
            host.replaceChildren(keep);
            const controller = new AbortController();
            const pending = start(controller.signal);
            controller.abort();
            const code = await pending.then(() => "resolved", (error) => error.code);
            await settle();
            return { code, kept: host.childElementCount === 1 && host.firstElementChild === keep, display: host.style.display };
          };
          out.pptx = await attempt((signal) => pptx.render({ slide: deck.slides.get(0), target: host, signal }));
          out.xlsx = await attempt((signal) => xlsx.render({ sheet: book.sheets.get(0), target: host, signal }));
          out.docx = await attempt((signal) => docx.render({ document: doc, target: host, signal }));
          out.docxLayout = await attempt((signal) => docx.render({ document: doc, target: host, pagination: { mode: "layout" }, signal }));
          out.already = await attempt(() => pptx.render({ slide: deck.slides.get(0), target: host, signal: AbortSignal.abort() }));
          const controller = new AbortController();
          const painted = await pptx.render({ slide: deck.slides.get(0), target: host, signal: controller.signal });
          controller.abort();
          out.afterPaint = { destroyed: painted.destroyed, slides: host.querySelectorAll('[data-aared-element="slide"]').length };
          painted.destroy();
          deck.close();
          book.close();
          doc.close();
          return out;
        };
        // Aborted after it painted but before it resolved (it is waiting for
        // fonts): the render takes its own nodes back out, gives the target its
        // layout back and rejects aborted. Only its own nodes: what something
        // else put in the target meanwhile stays.
        window.cancelPainted = async () => {
          const deck = await pptx.open("/deck.pptx");
          const book = await xlsx.open("/book.xlsx");
          const doc = await docx.open("/doc.docx");
          const out = {};
          const attempt = async (start, replaceFirst) => {
            host.replaceChildren();
            const controller = new AbortController();
            let painted = false;
            const kept = document.createElement("p");
            // A mutation is reported once the paint has run, before the render's font wait can end.
            const observer = new MutationObserver(() => {
              if (painted || host.childElementCount === 0) return;
              painted = true;
              if (replaceFirst) host.replaceChildren(kept);
              controller.abort();
            });
            observer.observe(host, { childList: true });
            const code = await start(controller.signal).then(() => "resolved", (error) => error.code);
            observer.disconnect();
            await settle();
            return {
              painted,
              code,
              children: host.childElementCount,
              kept: host.firstElementChild === kept,
              display: host.style.display,
            };
          };
          out.pptx = await attempt((signal) => pptx.render({ slide: deck.slides.get(0), target: host, signal }), false);
          out.xlsx = await attempt((signal) => xlsx.render({ sheet: book.sheets.get(0), target: host, signal }), false);
          out.docx = await attempt((signal) => docx.render({ document: doc, target: host, signal }), false);
          out.docxLayout = await attempt((signal) => docx.render({ document: doc, target: host, pagination: { mode: "layout" }, signal }), false);
          out.othersKept = await attempt((signal) => pptx.render({ slide: deck.slides.get(1), target: host, signal }), true);
          host.replaceChildren();
          deck.close();
          book.close();
          doc.close();
          return out;
        };
      </script>`,
    width: 1000,
    height: 700,
  }, async (view) => {
    const out = await view.evaluate(`replace()`) as Record<string, unknown>;
    expect(out).toEqual({
      pptx: [true, false],
      pptxByDocx: true,
      docx: [true, false],
      docxByXlsx: true,
      xlsx: [true, false],
      display: "flex",
      zoomed: 2,
      switched: false,
      xlsxByDocx: true,
      restored: "",
    });
    const cancelled = { code: "aborted", kept: true, display: "" };
    expect(await view.evaluate(`cancel()`) as Record<string, unknown>).toEqual({
      pptx: cancelled,
      xlsx: cancelled,
      docx: cancelled,
      docxLayout: cancelled,
      already: cancelled,
      afterPaint: { destroyed: false, slides: 1 },
    });
    const takenBack = { painted: true, code: "aborted", children: 0, kept: false, display: "" };
    expect(await view.evaluate(`cancelPainted()`) as Record<string, unknown>).toEqual({
      pptx: takenBack,
      xlsx: takenBack,
      docx: takenBack,
      docxLayout: takenBack,
      othersKept: { painted: true, code: "aborted", children: 1, kept: true, display: "" },
    });
  });
}, 60000);
