/**
 * Excel's gradient and pattern cell fills paint as a layer behind the cell's
 * text, through the shared fill painter. Re-reads the rendered DOM of the
 * coverage workbook's Fills sheet (authored through Excel's COM object model):
 * the layer, its paint server and the server's stops or tile.
 */
import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

const XLSX_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

interface PaintedFill {
  layer: boolean;
  server: string | null;
  units: string | null;
  stops: string[];
  patternBox: number[] | null;
  tileCells: number;
  background: string;
  fallback: string | null;
  text: string;
  textAbove: boolean;
}

visual("visual DOM output: XLSX gradient and pattern cell fills paint behind the text", async () => {
  const [xlsxModule, workerModule] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/xlsx/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
  ]);
  await withVisualPage({
    modules: { "/xlsx.mjs": xlsxModule, "/worker.mjs": workerModule },
    files: { "/fixture.xlsx": await Bun.file(XLSX_PATH).arrayBuffer() },
    page: `<!doctype html><style>html,body{margin:0}#host{width:900px;height:500px}</style><main id="host"></main><script type="module">
        import { open, render } from "/xlsx.mjs";
        window.run = async () => {
          const workbook = await open("/fixture.xlsx");
          const host = document.querySelector("#host");
          const view = await render({ sheet: workbook.sheets.get("Fills"), target: host });
          const read = (row) => {
            const cell = host.querySelector('td[data-aared-row="' + row + '"][data-aared-col="1"]');
            const svg = cell?.querySelector(':scope > svg[data-aared-element="cell-fill"]');
            // The layer's own rect; a pattern server in <defs> holds rects of its own.
            const url = svg?.querySelector(":scope > rect")?.getAttribute("fill") ?? "";
            const server = url.startsWith("url(#") ? svg.querySelector("#" + CSS.escape(url.slice(5, -1))) : null;
            const text = cell?.querySelector('[data-aared-element="cell-text"]');
            return {
              layer: svg != null,
              server: server?.localName ?? null,
              units: server?.getAttribute("gradientUnits") ?? server?.getAttribute("patternUnits") ?? null,
              stops: [...(server?.querySelectorAll("stop") ?? [])].map((stop) => stop.getAttribute("stop-color")),
              patternBox: server?.localName === "pattern" ? ["x", "y", "width", "height"].map((name) => Number(server.getAttribute(name))) : null,
              tileCells: server?.localName === "pattern" ? server.querySelectorAll("rect").length - 1 : 0,
              background: cell ? getComputedStyle(cell).backgroundColor : "",
              fallback: cell?.getAttribute("data-aared-fill-fallback") ?? null,
              text: text?.textContent ?? "",
              // The text paints over the layer: the topmost element at the text's centre is the text's.
              textAbove: (() => {
                if (!text) return false;
                const box = text.getBoundingClientRect();
                const hit = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2);
                return hit === text || text.contains(hit);
              })(),
            };
          };
          const painted = [1, 2, 3, 4, 5, 6].map(read);
          view.destroy();
          workbook.close();
          return painted;
        };
      </script>`,
    width: 1000,
    height: 600,
  }, async (view) => {
    const [linear90, linear0, center, darkGrid, gray125, lightDown] = await view.evaluate("run()") as PaintedFill[];
    for (const [cell, text] of [[linear90, "1200"], [linear0, "850"], [center, "430"], [darkGrid, "300"], [gray125, "125"], [lightDown, "75"]] as const) {
      expect(cell!.layer, text).toBe(true);
      expect(cell!.fallback, text).toBeNull();
      // No solid background stands in for the layer, and the value still reads over it.
      expect(cell!.background, text).toBe("rgba(0, 0, 0, 0)");
      expect(cell!.text.replace(/\s/g, ""), text).toBe(text);
      expect(cell!.textAbove, text).toBe(true);
    }
    // Excel blends cell gradients straight in sRGB: two authored stops, none interpolated.
    expect(linear90!.server).toBe("linearGradient");
    expect(linear90!.stops).toEqual(["#FFFFFF", "#F4B183"]);
    expect(linear0!.server).toBe("linearGradient");
    expect(linear0!.stops).toEqual(["#DDEBF7", "#2F5597"]);
    // From the centre: the rectangular path gradient, drawn as four sides in a pattern.
    expect(center!.server).toBe("pattern");
    // Excel's patterns: anchored at the cell's own top-left, one CSS pixel per bitmap cell.
    expect(darkGrid!.server).toBe("pattern");
    expect(darkGrid!.patternBox).toEqual([0, 0, 8, 8]);
    expect(darkGrid!.tileCells).toBeGreaterThan(0);
    expect(gray125!.patternBox).toEqual([0, 0, 8, 8]);
    expect(lightDown!.patternBox).toEqual([0, 0, 8, 8]);
  });
}, 60000);
