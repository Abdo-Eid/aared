import { join } from "node:path";
import type { ExcelRender } from "../../../tools/review/office-export.js";
import { reviewPage, escapeHtml } from "../../../tools/review/page.js";
import { browserModule } from "./harness.js";

/** What the review page reads from a native Excel export; both callers' render types satisfy it. */
export type ReviewReference = ExcelRender;

/**
 * Shared XLSX review pieces: the A4 host page the aared side is laid out in,
 * the bundled renderer it loads, and the review page that pairs each native
 * Excel print with its aared capture. Used by the committed-fixture visual
 * suite (`packages/aared/test/xlsx-capture.visual.test.ts`) and by
 * `tools/review/xlsx.ts`, which pairs any workbook.
 */

/** A4 portrait at 96 CSS px/in with Excel "Normal" margins: the page box the native tool prints. */
export const PAGE = { width: 794, height: 1123, top: 72, bottom: 72, left: 67, right: 67 };

export interface AaredPart {
  name: string;
  range: string;
  file: string;
  scale: number;
}

export interface AaredSheet {
  name: string;
  direction: string | null;
  file: string;
  scale: number;
  parts: AaredPart[];
}

/** The bundled XLSX renderer, served to the host page as `/xlsx.mjs`. */
export async function buildXlsxRenderer(): Promise<string> {
  return browserModule(join(import.meta.dir, "../src/xlsx/renderer.ts"));
}

/**
 * The page that lays out one sheet at a time. `absentFonts` loads the
 * `/absent-fonts.mjs` module the caller must serve.
 */
export function hostPage(absentFonts: boolean): string {
  return `<!doctype html><style>
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
  ${absentFonts ? 'import "/absent-fonts.mjs";' : ""}
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
    const canvas = document.querySelector('[data-aared-element="sheet-canvas"]');
    scroller.style.minWidth = "0";
    scroller.style.minHeight = "0";
    // Expand to the whole grid, not an arbitrary huge width: RTL deliberately
    // hugs the scroller's right edge, which otherwise puts every cell offscreen.
    scroller.style.width = canvas.style.width;
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
    const painted = grid.getBoundingClientRect();
    const viewport = document.querySelector("#viewport").getBoundingClientRect();
    if (painted.right <= viewport.left || painted.left >= viewport.right || painted.bottom <= viewport.top || painted.top >= viewport.bottom) {
      throw new Error("aared visual review: part " + range + " is outside the capture viewport");
    }
    return { scale };
  };
</script>`;
}

/**
 * The review page that pairs each native Excel print with its aared capture,
 * written beside the native pages in the export's output directory.
 */
export async function writeReview(
  reference: ReviewReference,
  fixtureFile: string,
  fixtureHash: string,
  sheets: AaredSheet[],
  absentFonts: boolean,
): Promise<void> {
  const cards = sheets.map((sheet, index) => {
    const native = reference.sheets[index];
    const partPairs = sheet.parts.map((part) => {
      const nativePart = native?.parts.find((entry) => entry.name === part.name) ?? null;
      return `
      <h3>${escapeHtml(part.name)} (${escapeHtml(part.range)})</h3>
      <div class="pair">
        <figure><figcaption>Excel ${reference.excelVersion} print (${reference.rasterizer.tool} ${reference.rasterizer.version}, ${reference.rasterizer.dpi} dpi)</figcaption>${nativePart === null ? "<p>No native counterpart part.</p>" : `<img src="${nativePart.file}">`}</figure>
        <figure><figcaption>aared DOM grid (scale ${part.scale})</figcaption><img src="${part.file}"></figure>${absentFonts ? `
        <figure><figcaption>aared, Office fonts absent</figcaption><img src="aared-absent-${part.file.slice("aared-".length)}"></figure>` : ""}
      </div>`;
    }).join("\n");
    return `
    <section>
      <h2>Sheet ${index + 1} — ${escapeHtml(sheet.name)} (${sheet.direction ?? "unknown direction"})</h2>
      <div class="pair">
        <figure><figcaption>Excel ${reference.excelVersion} print</figcaption>${native?.file === null || native === undefined ? "<p>Empty sheet.</p>" : `<img src="${native.file}">`}</figure>
        <figure><figcaption>aared DOM grid on A4 (scale ${sheet.scale})</figcaption><img src="${sheet.file}"></figure>${absentFonts ? `
        <figure><figcaption>aared, no Office fonts (aared-fonts substitutes)</figcaption><img src="aared-absent-${sheet.file.slice("aared-".length)}"></figure>` : ""}
      </div>${partPairs}
    </section>`;
  }).join("\n");
  await Promise.all([
    Bun.write(join(reference.outputDir, "review.html"), reviewPage(`aared XLSX review - ${fixtureFile}`,
      `<p>Native A4 print vs aared A4 fit, scaled to equal display width. Human review only, never pixel diffs. Parts are curated row bands of the live grid (the other rows hidden, same content as the full sheet, at full resolution); row ranges label each part. The formula notice is viewer chrome hidden here for print parity (covered by DOM tests).</p>${cards}`)),
    Bun.write(join(reference.outputDir, "review.json"), JSON.stringify({
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
