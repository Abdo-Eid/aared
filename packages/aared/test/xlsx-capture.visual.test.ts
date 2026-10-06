import { expect, test } from "bun:test";
import { pngSize, within } from "./harness.js";
import { createHash } from "node:crypto";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { open } from "../src/xlsx/index.js";
import { ABSENT_FONTS, absentFontsModule } from "./absent-fonts.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const RETAIN_ARTIFACTS = process.env.AARED_VISUAL_ARTIFACTS === "1";
const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;
// A4 portrait at 96 CSS px/in with Excel "Normal" margins: the page box the
// native tool prints. The full sheet only ever shrinks to fit, exactly like
// Excel's fit-to-page.
const PAGE = { width: 794, height: 1123, top: 72, bottom: 72, left: 67, right: 67 };

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
      // Deep slice forces real scrolling + virtualized re-render (the exact
      // path that once shipped stale pixels past a green gate).
      2: [{ name: "deep-rows", range: "A60:I70" }],
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

interface ExcelPartRender {
  name: string;
  range: string;
  file: string;
  widthPx: number;
  heightPx: number;
}

interface ExcelSheetRender {
  index: number;
  name: string;
  direction: "rtl" | "ltr";
  usedRange: string;
  empty: boolean;
  file: string | null;
  widthPx: number;
  heightPx: number;
  parts: ExcelPartRender[];
}

interface ExcelRender {
  outputDir: string;
  excelVersion: string;
  sheetCount: number;
  sheets: ExcelSheetRender[];
  rasterizer: { tool: string; version: string; dpi: number };
  pages: string[];
}

interface ExcelPartSpec {
  sheet: number;
  name: string;
  range: string;
}

interface AaredPart {
  name: string;
  range: string;
  file: string;
  scale: number;
}

interface AaredSheet {
  name: string;
  direction: string | null;
  file: string;
  scale: number;
  parts: AaredPart[];
}

async function renderExcel(fixturePath: string, parts: ExcelPartSpec[]): Promise<ExcelRender> {
  const child = Bun.spawn(["bun", "tools/office-export.ts", "--format", "xlsx", fixturePath,
    ...parts.flatMap((part) => ["--part", `${part.sheet}:${part.name}=${part.range}`])], {
    cwd: `${import.meta.dir}/../../..`,
    stdout: "pipe",
    stderr: "pipe",
  });
  const [exitCode, stdout, stderr] = await Promise.all([
    child.exited,
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
  ]);
  if (exitCode !== 0) {
    throw new Error(`aared visual review: Excel export failed with exit code ${exitCode}: ${(stderr || stdout).trim()}`);
  }
  return JSON.parse(stdout) as ExcelRender;
}

async function browserRenderer(): Promise<string> {
  const result = await Bun.build({
    entrypoints: [`${import.meta.dir}/../src/xlsx/index.ts`],
    target: "browser",
    format: "esm",
  });
  if (!result.success) throw new Error(result.logs.map((log) => log.message).join("\n"));
  return result.outputs[0]!.text();
}


const HOST_PAGE = `<!doctype html><style>
  html,body{margin:0;background:#fff}
  #viewport{width:${PAGE.width}px;height:${PAGE.height}px;overflow:hidden;background:#fff}
  #pad{padding:${PAGE.top}px ${PAGE.right}px ${PAGE.bottom}px ${PAGE.left}px;width:max-content}
  #host{transform-origin:top left}
  [data-aared-element="sheet-notices"]{display:none !important}
  [data-aared-frozen]{position:static !important}
  /* Capture chrome, not renderer behavior: the scrolling grid keeps live
  scrollbars for users, but the A4 review pairs print pages (no
  scrollbars) so the harness hides them here. */
  [data-aared-element="sheet-scroll"]{scrollbar-width:none !important}
  [data-aared-element="sheet-scroll"]::-webkit-scrollbar{display:none !important}
</style><div id="viewport"><div id="pad"><main id="host"></main></div></div><script type="module">
  import { render } from "/xlsx.mjs";
  ${ABSENT_FONTS ? 'import "/absent-fonts.mjs";' : ""}
  window.renderedSheet = null;
  const FIT_W = ${PAGE.width - PAGE.left - PAGE.right};
  const FIT_H = ${PAGE.height - PAGE.top - PAGE.bottom};
  const raf = () => new Promise(requestAnimationFrame);
  const raf3 = () => raf().then(raf).then(raf);
  // Unhandled page errors (notably exceptions thrown inside
  // requestAnimationFrame callbacks, which never reach the test runner)
  // are collected here so failures name their cause instead of going
  // silent behind a green-looking gate.
  window.__errors = [];
  window.onerror = (message) => { window.__errors.push(String(message)); };
  // Force the virtualized grid to re-render at live metrics. Resizes never
  // fire scroll events, so without this the DOM keeps the first paint's
  // window (rows beyond it never materialize). A synthetic scroll event is
  // synchronous: the render pass it schedules always runs before the frames
  // awaited after it — strictly ordered, unlike real scroll timing.
  const resync = async (scroller) => {
    scroller.dispatchEvent(new Event("scroll"));
    await raf3();
  };
  const colNameToIndex = (name) => {
    let index = 0;
    for (const ch of name) index = index * 26 + (ch.charCodeAt(0) - 64);
    return index - 1;
  };
  const parseRange = (range) => {
    const match = range.match(/^([A-Z]+)(\\d+):([A-Z]+)(\\d+)$/);
    if (match === null) throw new Error("aared visual review: malformed range " + range);
    return {
      c0: colNameToIndex(match[1]), r0: Number(match[2]) - 1,
      c1: colNameToIndex(match[3]), r1: Number(match[4]) - 1,
    };
  };
  // Render one sheet and report its natural size. Frozen stickiness is
  // neutralized by stylesheet (native print never repeats frozen rows),
  // while the sticky heading row/column repeat on scroll like print
  // headings. The viewport fill is off: it pads a short sheet with inert
  // empty rows out to the clipping ancestor (here the whole A4 box), which
  // a printed page never has and which would inflate the natural size the
  // fit-to-page scale is measured from.
  // Detached models carry no session, so the paired review feeds the
  // fixture's own bytes (same honesty rule as the PPTX capture): without
  // this every floating picture paints a placeholder and the Excel side of
  // the pair has nothing honest to compare against.
  const decode = (base64) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)).buffer;
  const resolverFor = (media) => media === undefined ? undefined : async (mediaId) => {
    const base64 = media[mediaId];
    if (base64 === undefined) throw new Error("missing media " + mediaId);
    return decode(base64);
  };
  const waitForDrawings = () => new Promise((resolve, reject) => {
    const deadline = Date.now() + 8000;
    const poll = () => {
      const images = [...document.querySelectorAll('[data-aared-element="picture"] img')];
      if (images.every((img) => img.complete && img.naturalWidth > 0)) resolve(undefined);
      else if (Date.now() > deadline) reject(new Error("timed out waiting for sheet images"));
      else setTimeout(poll, 50);
    };
    poll();
  });
  window.renderFull = async (model, media, absent) => {
    window.renderedSheet?.destroy();
    if (absent) model = await window.simulateAbsentFonts(model);
    const host = document.querySelector("#host");
    host.style.transform = "";
    host.style.width = "auto";
    window.renderedSheet = await render({
      sheet: model,
      target: host,
      appearance: { spreadsheet: { headers: true, fill: false } },
      mediaResolver: resolverFor(media),
    });
    // A sheet loads a delivered face only once its text lays out: let that
    // happen before waiting for the fonts in flight.
    await raf3();
    await document.fonts.ready;
    await waitForDrawings();
    await raf3();
    const grid = document.querySelector('[data-aared-element="grid"]');
    const scroller = document.querySelector('[data-aared-element="sheet-scroll"]');
    await resync(scroller);
    const canvas = document.querySelector('[data-aared-element="sheet-canvas"]');
    // Fit sheet on one page: shrink only, never stretch.
    const scale = Math.min(1, FIT_W / canvas.scrollWidth, FIT_H / canvas.scrollHeight);
    if (scale < 1) {
      host.style.width = canvas.scrollWidth + "px";
      host.style.transform = "scale(" + scale + ")";
      await raf3();
    }
    return {
      direction: grid?.dataset.aaredSheetDirection ?? null,
      canvasCount: document.querySelectorAll("canvas").length,
      scale,
      totalHeight: canvas.scrollHeight,
    };
  };
  // Isolate a curated row band for a part capture. Bands span the full
  // width starting at column A (a band past column A fails loud). Steps:
  // drop any fit transform (parts render at full resolution), expand the
  // scroller so virtualization renders every row, prove the band's rows
  // exist, then hide every other body row so the band sits at the top like
  // a print area. No scrolling is involved at any point: scroll positions
  // proved unassertable here (positions read back zero with no error), so
  // the design refuses to depend on them. Merges crossing a band edge
  // would paint across the cut: every curated band is chosen merge-clean.
  window.bandPart = async (model, range) => {
    const { c0, r0, c1, r1 } = parseRange(range);
    if (c0 !== 0) throw new Error("aared visual review: parts must start at column A");
    const host = document.querySelector("#host");
    host.style.transform = "";
    host.style.width = "auto";
    const grid = document.querySelector('[data-aared-element="grid"]');
    const scroller = document.querySelector('[data-aared-element="sheet-scroll"]');
    scroller.style.minWidth = "0";
    scroller.style.minHeight = "0";
    scroller.style.width = "100000px";
    scroller.style.height = "100000px";
    await resync(scroller);
    const missing = [];
    for (let row = r0; row <= r1; row++) {
      if (grid.querySelector('tr[data-aared-row="' + row + '"]') === null) missing.push(row);
    }
    if (missing.length > 0 || window.__errors.length > 0) {
      throw new Error("aared visual review: part " + range + " missing rows [" + missing.join(",") + "]"
        + " pageErrors=[" + window.__errors.join(" | ") + "]");
    }
    for (const line of grid.querySelectorAll("tbody tr")) {
      const row = line.dataset.aaredRow === undefined ? -1 : Number(line.dataset.aaredRow);
      if (row < r0 || row > r1) line.style.display = "none";
    }
    // Floating drawings are positioned absolute from the sheet top, so rows
    // hidden above would detach them: shift every drawing box up by exactly
    // the collapsed height (the band's own top in full-grid coordinates),
    // then hide boxes that do not intersect the band at all — native print
    // areas clip those away, and a leftover chart floating in whitespace
    // below the band is the visible symptom. Partial overlaps keep their
    // full box (no curated band splits a drawing).
    const hiddenR = new Set(model.hiddenRows);
    const heights = new Map(model.rows.map((row) => [row.index, row.height ?? model.defaultRowHeight]));
    let shift = 0;
    for (let row = 0; row < r0; row++) {
      if (!hiddenR.has(row)) shift += (heights.get(row) ?? model.defaultRowHeight) * 96;
    }
    if (shift > 0) {
      for (const box of grid.parentElement.querySelectorAll('[data-aared-element="chart"],[data-aared-element="picture"]')) {
        const element = box;
        element.style.top = Math.max(0, Number.parseFloat(element.style.top || "0") - shift) + "px";
      }
    }
    await raf3();
    const gridTop = grid.getBoundingClientRect().top;
    const bandTop = grid.querySelector('tr[data-aared-row="' + r0 + '"]').getBoundingClientRect().top - gridTop;
    const bandBottom = grid.querySelector('tr[data-aared-row="' + r1 + '"]').getBoundingClientRect().bottom - gridTop;
    for (const box of grid.parentElement.querySelectorAll('[data-aared-element="chart"],[data-aared-element="picture"]')) {
      const element = box;
      const boxTop = Number.parseFloat(element.style.top || "0");
      const boxBottom = boxTop + element.offsetHeight;
      if (boxBottom <= bandTop || boxTop >= bandBottom) element.style.display = "none";
    }
    await raf3();
    if (window.__errors.length > 0) {
      throw new Error("aared visual review: errors while isolating " + range + " [" + window.__errors.join(" | ") + "]");
    }
    // Wide bands shrink to the printable width (same rule as full sheets);
    // tall bands fail loud — parts are viewports by curation, not pages.
    // No scrolling is involved, so transformed coordinates need no mapping.
    const canvas = document.querySelector('[data-aared-element="sheet-canvas"]');
    if (canvas.scrollHeight > FIT_H + 1) {
      throw new Error("aared visual review: part " + range + " exceeds printable height");
    }
    let scale = 1;
    if (canvas.scrollWidth > FIT_W + 1) {
      scale = FIT_W / canvas.scrollWidth;
      host.style.width = canvas.scrollWidth + "px";
      host.style.transform = "scale(" + scale + ")";
      await raf3();
    }
    if (window.__errors.length > 0) {
      throw new Error("aared visual review: errors while shrinking " + range + " [" + window.__errors.join(" | ") + "]");
    }
    return { scale };
  };
</script>`;

async function writeReview(reference: ExcelRender, fixtureFile: string, fixtureHash: string, sheets: AaredSheet[]): Promise<void> {
  const cards = sheets.map((sheet, index) => {
    const native = reference.sheets[index];
    const partPairs = sheet.parts.map((part) => {
      const nativePart = native?.parts.find((entry) => entry.name === part.name) ?? null;
      return `
      <h3>${part.name} (${part.range})</h3>
      <div class="pair">
        <figure><figcaption>Excel ${reference.excelVersion} print (${reference.rasterizer.tool} ${reference.rasterizer.version}, ${reference.rasterizer.dpi} dpi)</figcaption>${nativePart === null ? "<p>No native counterpart part.</p>" : `<img src="${nativePart.file}">`}</figure>
        <figure><figcaption>aared DOM grid (scale ${part.scale})</figcaption><img src="${part.file}"></figure>
      </div>`;
    }).join("\n");
    return `
    <section>
      <h2>Sheet ${index + 1} — ${sheet.name} (${sheet.direction ?? "unknown direction"})</h2>
      <div class="pair">
        <figure><figcaption>Excel ${reference.excelVersion} print</figcaption>${native?.file === null || native === undefined ? "<p>Empty sheet.</p>" : `<img src="${native.file}">`}</figure>
        <figure><figcaption>aared DOM grid on A4 (scale ${sheet.scale})</figcaption><img src="${sheet.file}"></figure>${ABSENT_FONTS ? `
        <figure><figcaption>aared, no Office fonts (aared-fonts substitutes)</figcaption><img src="aared-absent-${sheet.file.slice("aared-".length)}"></figure>` : ""}
      </div>${partPairs}
    </section>`;
  }).join("\n");
  await Promise.all([
    writeFile(join(reference.outputDir, "review.html"), `<!doctype html>
<meta charset="utf-8"><title>aared XLSX visual review</title>
<style>body{margin:24px;background:#f4f4f4;color:#111;font:16px system-ui,sans-serif}h1{margin:0 0 8px}.pair{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(0,1fr);gap:16px}figure{margin:0 0 16px;background:#fff;padding:12px}figcaption{font-weight:700;margin-bottom:8px}img{display:block;width:100%;height:auto}@media(max-width:900px){.pair{grid-template-columns:1fr}}</style>
<h1>aared XLSX visual review — ${fixtureFile}</h1><p>Native A4 print vs aared A4 fit, scaled to equal display width. Human review only, never pixel diffs. Parts are curated row bands of the live grid (the other rows hidden, same content as the full sheet, at full resolution); row ranges label each part. The formula notice is viewer chrome hidden here for print parity (covered by DOM tests).</p>${cards}`),
    writeFile(join(reference.outputDir, "review.json"), JSON.stringify({
      fixture: fixtureFile,
      fixtureSha256: fixtureHash,
      excelVersion: reference.excelVersion,
      rasterizer: reference.rasterizer,
      page: { ...PAGE, unit: "px at 96 per inch" },
      sheetCount: reference.sheetCount,
      sheets: sheets.map((sheet, index) => ({
        index,
        name: sheet.name,
        direction: sheet.direction,
        scale: sheet.scale,
        excelFile: reference.sheets[index]?.file ?? null,
        aaredFile: sheet.file,
        parts: sheet.parts,
      })),
    }, null, 2)),
  ]);
}

visual("visual review: XLSX pairs Excel print with aared full sheets and parts", async () => {
  const xlsx = await browserRenderer();
  const absentFonts = ABSENT_FONTS ? await absentFontsModule() : "";
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/xlsx.mjs") {
        return new Response(xlsx, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (new URL(request.url).pathname === "/absent-fonts.mjs") {
        return new Response(absentFonts, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(HOST_PAGE, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
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
            models.push(await handle.read());
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
          // and its summary always fit at full size.
          expect(full.scale).toBeGreaterThan(0);
          expect(full.scale).toBeLessThanOrEqual(1);
          if (plan.path === COVERAGE_PATH && index <= 1) {
            expect(full.scale).toBe(1);
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
              const file = `aared-sheet-${index + 1}-part-${part.name}.png`;
              await writeFile(join(reference.outputDir, file), screenshot);
              parts.push({ name: part.name, range: part.range, file, scale: band.scale });
            }
          }
          sheets.push({ name: names[index] ?? `sheet-${index + 1}`, direction: full.direction, file: fullFile, scale: full.scale, parts });
        }
        if (RETAIN_ARTIFACTS) {
          await writeReview(reference, plan.file, fixtureHash, sheets);
        }
      } finally {
        if (!RETAIN_ARTIFACTS) await rm(reference.outputDir, { recursive: true, force: true });
      }
    }
  } finally {
    server.stop();
  }
}, ABSENT_FONTS ? 600_000 : 300_000);
