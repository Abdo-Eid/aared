import { expect } from "bun:test";
import { pngSize, serveVisual, visual, within } from "./harness.js";
import { createHash } from "node:crypto";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { open, readSheetModel } from "../src/xlsx/renderer.js";
import type { SheetModel } from "../src/core/model.js";
import { ABSENT_FONTS, absentFontsModule } from "./absent-fonts.js";
import { buildXlsxRenderer, hostPage, writeReview, type AaredPart, type AaredSheet } from "./xlsx-review.js";
import { renderExcel } from "../../../tools/review/office-export.js";
import { partSlug } from "../../../tools/shared/office.js";

const RETAIN_ARTIFACTS = process.env.AARED_VISUAL_ARTIFACTS === "1";
const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;

interface PartSpec {
  name: string;
  range: string;
}

interface FixturePlan {
  file: string;
  path: string;
  /** Sheet index (0-based) to curated print-area parts. */
  parts: Record<number, PartSpec[]>;
}

const PLANS: FixturePlan[] = [
  {
    file: "xlsx-coverage.xlsx",
    path: COVERAGE_PATH,
    parts: {
      0: [
        { name: "title-block", range: "A1:E2" },
        { name: "totals", range: "A9:E11" },
      ],
      1: [{ name: "sales-summary", range: "A1:G11" }],
      // Deep slice forces real scrolling + virtualized re-render (the exact
      // path that once shipped stale pixels past a green gate).
      2: [{ name: "deep-rows", range: "A60:I70" }],
      4: [{ name: "three-colour-scale", range: "A1:J13" }],
    },
  },
  {
    file: "post-production-tracker.xlsx",
    path: TRACKER_PATH,
    parts: {
      0: [
        { name: "title-kpis", range: "A1:H6" },
        { name: "episode-table", range: "A11:H14" },
        { name: "chart", range: "A18:H32" },
      ],
      1: [
        { name: "head", range: "A1:I8" },
        { name: "mid", range: "A9:I16" },
      ],
    },
  },
];

visual("visual review: XLSX pairs Excel print with aared full sheets and parts", async () => {
  const xlsx = await buildXlsxRenderer();
  const absentFonts = ABSENT_FONTS ? await absentFontsModule() : "";
  const server = serveVisual({
    modules: { "/xlsx.mjs": xlsx, "/absent-fonts.mjs": absentFonts },
    page: hostPage(ABSENT_FONTS),
  });
  try {
    await using view = new Bun.WebView({ width: 810, height: 1218 });
    await within("loading the WebView fixture host", view.navigate(server.url.href), 30_000);
    for (const plan of PLANS) {
      const specs: Array<{ sheet: number; name: string; range: string }> = [];
      for (const [sheet, parts] of Object.entries(plan.parts)) {
        for (const part of parts) specs.push({ sheet: Number(sheet) + 1, ...part });
      }
      const reference = await renderExcel(plan.path, specs);
      try {
        const fixtureHash = createHash("sha256").update(new Uint8Array(await Bun.file(plan.path).arrayBuffer())).digest("hex");
        const workbook = await open(await Bun.file(plan.path).arrayBuffer());
        const models: unknown[] = [];
        const names: string[] = [];
        let media: Record<string, string>;
        try {
          for (const handle of workbook.sheets.list()) {
            models.push(await readSheetModel(handle));
          }
          names.push(...workbook.info.titles);
          expect(workbook.info.format).toBe("xlsx");
          expect(reference.sheetCount).toBe(workbook.info.count);
          const ids = new Set<string>();
          for (const match of JSON.stringify(models).matchAll(/"(?:mediaId|pictureMediaId)":"([^"]+)"/g)) ids.add(match[1]!);
          media = {};
          for (const id of ids) media[id] = Buffer.from(await workbook.media.get(id)).toString("base64");
        } finally {
          workbook.close();
        }
        const sheets: AaredSheet[] = [];
        for (const [index, model] of models.entries()) {
          const full = await within(`rendering ${plan.file} sheet ${index + 1}`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${JSON.stringify(media)})`), 30_000) as {
            direction: string | null; canvasCount: number; scale: number; totalHeight: number;
          };
          expect(full.canvasCount).toBe(0);
          // The rendered grid must agree with Excel on sheet direction.
          expect(full.direction).toBe(reference.sheets[index]?.direction ?? null);
          // Any overfull sheet genuinely shrinks (tracker Dashboard is wider
          // than printable A4, as are the deep log and notes); the RTL grid
          // grid fits at full size; the bilingual sales summary may shrink to
          // fit its long notes column on the page.
          expect(full.scale).toBeGreaterThan(0);
          expect(full.scale).toBeLessThanOrEqual(1);
          if (plan.path === COVERAGE_PATH && index === 0) {
            expect(full.scale).toBe(1);
          }
          if (plan.path === COVERAGE_PATH && names[index] === "CF Rules") {
            const painted = await within("reading the three-colour scale from the DOM", view.evaluate(`
              [...document.querySelectorAll('[data-aared-element="cell"][data-aared-col="4"]')]
                .filter(cell => Number(cell.dataset.aaredRow) >= 1 && Number(cell.dataset.aaredRow) <= 11)
                .map(cell => ({ text: cell.textContent, fill: getComputedStyle(cell).backgroundColor }))
            `), 30_000) as Array<{ text: string; fill: string }>;
            expect(painted.map((cell) => cell.text)).toEqual(["0", "10", "20", "30", "40", "50", "60", "70", "80", "90", "100"]);
            const sheet = model as SheetModel;
            const fills = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 4 && typeof entry.cell.value === "number").map((entry) => entry.cell.conditional!.fill!));
            expect(painted.map((cell) => cell.fill)).toEqual(fills.map((fill) => {
              const channels = [1, 3, 5].map((offset) => Number.parseInt(fill.slice(offset, offset + 2), 16));
              return `rgb(${channels.join(", ")})`;
            }));
            expect(painted[0]!.fill).toBe("rgb(248, 105, 107)");
            expect(painted[5]!.fill).toBe("rgb(255, 235, 132)");
            expect(painted[10]!.fill).toBe("rgb(99, 190, 123)");
            const expression = await within("reading the expression rule from the DOM", view.evaluate(`
              [...document.querySelectorAll('[data-aared-element="cell"][data-aared-col="5"]')]
                .filter(cell => Number(cell.dataset.aaredRow) >= 1 && Number(cell.dataset.aaredRow) <= 11)
                .map(cell => ({ text: cell.textContent, fill: getComputedStyle(cell).backgroundColor }))
            `), 30_000) as Array<{ text: string; fill: string }>;
            expect(expression.map((cell) => cell.text)).toEqual(["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"]);
            // Even values carry the native pink fill; odds keep no conditional fill (transparent).
            for (const [i, cell] of expression.entries()) {
              const value = i + 1;
              if (value % 2 === 0) expect(cell.fill).not.toBe("rgba(0, 0, 0, 0)");
              else expect(cell.fill).toBe("rgba(0, 0, 0, 0)");
            }
            // G/H/I/J carry the closed formula gap: TODAY, COUNTIF duplicates,
            // SEARCH text and the x14 cross-sheet rule (§111).
            const due = await within("reading the TODAY rule from the DOM", view.evaluate(`
              [...document.querySelectorAll('[data-aared-element="cell"][data-aared-col="6"]')]
                .filter(cell => Number(cell.dataset.aaredRow) >= 1 && Number(cell.dataset.aaredRow) <= 12)
                .map(cell => ({ text: cell.textContent, fill: getComputedStyle(cell).backgroundColor }))
            `), 30_000) as Array<{ text: string; fill: string }>;
            expect(due).toHaveLength(12);
            for (const [i, cell] of due.entries()) {
              if (i < 6) expect(cell.fill).not.toBe("rgba(0, 0, 0, 0)");
              else expect(cell.fill).toBe("rgba(0, 0, 0, 0)");
            }
            const dupes = await within("reading the COUNTIF rule from the DOM", view.evaluate(`
              [...document.querySelectorAll('[data-aared-element="cell"][data-aared-col="7"]')]
                .filter(cell => Number(cell.dataset.aaredRow) >= 1 && Number(cell.dataset.aaredRow) <= 12)
                .map(cell => ({ text: cell.textContent, fill: getComputedStyle(cell).backgroundColor }))
            `), 30_000) as Array<{ text: string; fill: string }>;
            expect(dupes.map((cell) => cell.text)).toEqual(["apple", "banana", "Apple", "cherry", "banana", "date", "apple", "fig", "grape", "banana", "kiwi", "apple"]);
            for (const [i, cell] of dupes.entries()) {
              if ([0, 1, 2, 4, 6, 9, 11].includes(i)) expect(cell.fill).not.toBe("rgba(0, 0, 0, 0)");
              else expect(cell.fill).toBe("rgba(0, 0, 0, 0)");
            }
            const search = await within("reading the SEARCH rule from the DOM", view.evaluate(`
              [...document.querySelectorAll('[data-aared-element="cell"][data-aared-col="8"]')]
                .filter(cell => Number(cell.dataset.aaredRow) >= 1 && Number(cell.dataset.aaredRow) <= 8)
                .map(cell => ({ text: cell.textContent, fill: getComputedStyle(cell).backgroundColor }))
            `), 30_000) as Array<{ text: string; fill: string }>;
            expect(search.map((cell) => cell.text)).toEqual(["apple pie", "APPLE", "banana", "pineapple", "Avocado", "crabapple", "Cherry", "green apple"]);
            for (const [i, cell] of search.entries()) {
              if ([0, 1, 3, 5, 7].includes(i)) expect(cell.fill).not.toBe("rgba(0, 0, 0, 0)");
              else expect(cell.fill).toBe("rgba(0, 0, 0, 0)");
            }
            const cross = await within("reading the x14 cross-sheet rule from the DOM", view.evaluate(`
              [...document.querySelectorAll('[data-aared-element="cell"][data-aared-col="9"]')]
                .filter(cell => Number(cell.dataset.aaredRow) >= 1 && Number(cell.dataset.aaredRow) <= 12)
                .map(cell => ({ text: cell.textContent, fill: getComputedStyle(cell).backgroundColor }))
            `), 30_000) as Array<{ text: string; fill: string }>;
            expect(cross.map((cell) => cell.text)).toEqual(["100", "200", "300", "400", "500", "600", "700", "800", "900", "1000", "1100", "1200"]);
            for (const [i, cell] of cross.entries()) {
              if (i >= 4) expect(cell.fill).toBe("rgb(255, 0, 0)");
              else expect(cell.fill).toBe("rgba(0, 0, 0, 0)");
            }
          }
          const fullFile = `aared-sheet-${index + 1}.png`;
          if (RETAIN_ARTIFACTS) {
            const screenshot = await within(`capturing ${plan.file} sheet ${index + 1}`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000);
            expect(pngSize(screenshot).width / pngSize(screenshot).height).toBeCloseTo(Math.SQRT1_2, 2);
            await mkdir(reference.outputDir, { recursive: true });
            await writeFile(join(reference.outputDir, fullFile), screenshot);
          }
          if (ABSENT_FONTS) {
            await within(`rendering ${plan.file} sheet ${index + 1} without Office fonts`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${JSON.stringify(media)}, true)`), 60_000);
            const screenshot = await within(`capturing ${plan.file} sheet ${index + 1} without Office fonts`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000);
            await mkdir(reference.outputDir, { recursive: true });
            await writeFile(join(reference.outputDir, `aared-absent-sheet-${index + 1}.png`), screenshot);
          }
          const parts: AaredPart[] = [];
          for (const part of plan.parts[index] ?? []) {
            // Fresh render per part: bandPart mutates drawing boxes
            // (shift + clip) and hides off-band rows, so reusing one DOM
            // across parts would double-shift and carry hidden boxes into
            // the next part (the chart part once rendered empty for this
            // exact reason). Parts are viewports by curation, each from a
            // clean full render.
            // The band sits at the top of the fixed A4 viewport like a
            // native print-area page: same framing both sides.
            await within(`rendering ${part.name}`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${JSON.stringify(media)})`), 30_000);
            const band = await within(`isolating ${part.name}`, view.evaluate(`bandPart(${JSON.stringify(model)}, ${JSON.stringify(part.range)})`), 30_000) as {
              scale: number;
            };
            expect(band.scale).toBeGreaterThan(0);
            expect(band.scale).toBeLessThanOrEqual(1);
            if (RETAIN_ARTIFACTS) {
              const screenshot = await within(`capturing ${part.name}`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000);
              expect(pngSize(screenshot).width / pngSize(screenshot).height).toBeCloseTo(Math.SQRT1_2, 2);
              const file = `aared-sheet-${index + 1}-part-${partSlug(part.name)}.png`;
              await writeFile(join(reference.outputDir, file), screenshot);
              parts.push({ name: part.name, range: part.range, file, scale: band.scale });
              if (ABSENT_FONTS) {
                await within(`rendering ${part.name} without Office fonts`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${JSON.stringify(media)}, true)`), 60_000);
                await within(`isolating ${part.name} without Office fonts`, view.evaluate(`bandPart(${JSON.stringify(model)}, ${JSON.stringify(part.range)})`), 30_000);
                await writeFile(join(reference.outputDir, `aared-absent-${file.slice("aared-".length)}`), await view.screenshot({ encoding: "buffer", format: "png" }));
              }
            }
          }
          sheets.push({ name: names[index] ?? `sheet-${index + 1}`, direction: full.direction, file: fullFile, scale: full.scale, parts });
        }
        if (RETAIN_ARTIFACTS) {
          await writeReview(reference, plan.file, fixtureHash, sheets, ABSENT_FONTS);
        }
      } finally {
        if (!RETAIN_ARTIFACTS) await rm(reference.outputDir, { recursive: true, force: true });
      }
    }
  } finally {
    server.stop();
  }
}, ABSENT_FONTS ? 600_000 : 300_000);
