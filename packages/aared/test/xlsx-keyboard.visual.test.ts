import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;

/**
 * Keyboard navigation must survive virtualization: a move that scrolls
 * re-renders the row window, which destroys the focused cell, so the render
 * restores focus; otherwise focus would fall to <body> after the first move
 * and every later keydown would miss the grid (the listener lives below the cell).
 */
visual("visual DOM output: XLSX arrow-key navigation keeps focus and moves selection repeatedly", async () => {
  const pkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
  const workbook = await readWorkbook(pkg);
  const sheet = await parseXlsxSheet(pkg, workbook, 0);

  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/xlsx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px;height:400px}</style><main id="host"></main><script type="module">
        import { render } from "/xlsx.mjs";
        window.viewer = await render({ sheet: ${JSON.stringify(sheet)}, target: document.querySelector("#host") });
        window.selections = [];
        window.viewer.addEventListener("selectionchange", (event) => window.selections.push(event.detail.selection.address));
        window.clickCell = (row, col) => {
          const td = document.querySelector('td[data-aared-row="' + row + '"][data-aared-col="' + col + '"]');
          td.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
          td.focus();
        };
        window.pressKey = (key) => {
          document.activeElement.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
        };
        window.focusInGrid = () => !!document.activeElement?.closest?.("#host");
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 800 });
    await view.navigate(server.url.href);
    await view.evaluate(`new Promise((resolve) => { const poll = () => (window.viewer ? resolve(true) : setTimeout(poll, 20)); poll(); })`);

    await view.evaluate(`clickCell(2, 1)`);
    expect(await view.evaluate<string>(`window.viewer.selection?.address`)).toBe("B3");

    for (const address of ["B4", "B5", "B6"]) {
      await view.evaluate(`pressKey("ArrowDown")`);
      expect(await view.evaluate<string>(`window.viewer.selection?.address`)).toBe(address);
      expect(await view.evaluate<boolean>(`focusInGrid()`)).toBe(true);
    }
    for (const address of ["C6", "D6"]) {
      await view.evaluate(`pressKey("ArrowRight")`);
      expect(await view.evaluate<string>(`window.viewer.selection?.address`)).toBe(address);
      expect(await view.evaluate<boolean>(`focusInGrid()`)).toBe(true);
    }

    expect(await view.evaluate<string[]>(`window.selections`)).toEqual(["B3", "B4", "B5", "B6", "C6", "D6"]);
  } finally {
    server.stop(true);
  }
});
