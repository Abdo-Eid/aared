import { expect } from "bun:test";
import { join } from "node:path";
import { open } from "../../packages/aared/src/xlsx/index.js";
import type { SheetModel } from "../../packages/aared/src/core/model.js";
import { buildXlsxRenderer, hostPage } from "../../packages/aared/test/xlsx-review.js";
import { visual, within, withVisualPage } from "../../packages/aared/test/harness.js";

visual("visual XLSX review bands keep RTL and deep LTR cells inside the capture viewport", async () => {
  const workbook = await open(await Bun.file(join(import.meta.dir, "../../fixtures/xlsx-coverage.xlsx")).arrayBuffer());
  const plans = [
    { sheet: 0, range: "A1:E2", firstRow: 0, lastRow: 1, direction: "rtl" },
    { sheet: 2, range: "A60:I70", firstRow: 59, lastRow: 69, direction: "ltr" },
  ] as const;
  const models: SheetModel[] = [];
  try {
    for (const plan of plans) models.push(await workbook.sheets.get(plan.sheet).read());
  } finally { workbook.close(); }
  const module = await buildXlsxRenderer();
  await withVisualPage({
    modules: { "/xlsx.mjs": module },
    page: hostPage(false),
    width: 810,
    height: 1218,
  }, async (view) => {
    for (const [index, plan] of plans.entries()) {
      await within("rendering the review sheet", view.evaluate(`renderFull(${JSON.stringify(models[index])}, {})`), 30_000);
      await within("isolating the review band", view.evaluate(`bandPart(${JSON.stringify(models[index])}, ${JSON.stringify(plan.range)})`), 30_000);
      const painted = await view.evaluate(`(() => {
        const grid = document.querySelector('[data-aared-element="grid"]');
        const box = grid.getBoundingClientRect();
        const viewport = document.querySelector('#viewport').getBoundingClientRect();
        return { direction: grid.dataset.aaredSheetDirection, left: box.left, right: box.right,
          viewportLeft: viewport.left, viewportRight: viewport.right,
          rows: [...grid.querySelectorAll('tbody tr[data-aared-row]')].filter(row => getComputedStyle(row).display !== 'none').map(row => Number(row.dataset.aaredRow)),
          text: [...grid.querySelectorAll('tbody tr')].filter(row => getComputedStyle(row).display !== 'none').map(row => row.textContent).join(' ') };
      })()`) as { direction: string; left: number; right: number; viewportLeft: number; viewportRight: number; rows: number[]; text: string };
      expect(painted.direction).toBe(plan.direction);
      expect(painted.left).toBeGreaterThanOrEqual(painted.viewportLeft);
      expect(painted.right).toBeLessThanOrEqual(painted.viewportRight + 1);
      expect(painted.rows).toEqual(Array.from({ length: plan.lastRow - plan.firstRow + 1 }, (_, i) => plan.firstRow + i));
      expect(painted.text.trim().length).toBeGreaterThan(20);
    }
  });
}, 90_000);
