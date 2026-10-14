import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { findCoverageSlide } from "./fixture.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { findCoverageSheet } from "./xlsx-fixture.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

// Only the ancestor is hostile: this capability prevents inherited host
// traits, not arbitrary declarations directly targeting document descendants.
const HOSTILE_CSS = `.hostile {
  font-weight:700;font-style:italic;font-variant:small-caps;font-stretch:50%;
  letter-spacing:2px;word-spacing:5px;text-transform:uppercase;text-indent:2em;
}`;

function preflightCss(): Promise<string> {
  // Use the bench's actual installed reset; a missing dependency must fail,
  // not silently turn the preflight case into another clean-host case.
  return Bun.file(`${import.meta.dir}/../../../bench/node_modules/tailwindcss/preflight.css`).text();
}

interface TextSnapshot {
  text: string;
  fw: string;
  fs: string;
  fv: string;
  fst: string;
  ls: string;
  ws: string;
  tt: string;
  ti: string;
  family: string;
  size: string;
  pitch: string;
  top: number;
  left: number;
  width: number;
  height: number;
}

const PPTX_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

const SNAPSHOT_SCRIPT = `
window.setHostile = (on) => document.getElementById("ancestor").classList.toggle("hostile", on);
window.snapshot = () => {
  const out = [];
  for (const el of document.querySelectorAll("#host *")) {
    if (!el.textContent || !el.textContent.trim() || el.children.length !== 0) continue;
    const s = getComputedStyle(el);
    const r = el.getBoundingClientRect();
    out.push({
      text: el.textContent.trim(),
      fw: s.fontWeight,
      fs: s.fontStyle,
      fv: s.fontVariant,
      fst: s.fontStretch,
      ls: s.letterSpacing,
      ws: s.wordSpacing,
      tt: s.textTransform,
      ti: s.textIndent,
      family: s.fontFamily,
      size: s.fontSize,
      pitch: s.lineHeight,
      top: Math.round(r.top * 10) / 10,
      left: Math.round(r.left * 10) / 10,
      width: Math.round(r.width * 10) / 10,
      height: Math.round(r.height * 10) / 10,
    });
  }
  out.sort((a, b) => (a.text < b.text ? -1 : a.text > b.text ? 1 : 0));
  return out;
};
window.negativeControl = () => {
  const s = getComputedStyle(document.querySelector("#unprotected"));
  return { fw: s.fontWeight, fs: s.fontStyle, fv: s.fontVariant, fst: s.fontStretch, ls: s.letterSpacing, ws: s.wordSpacing, tt: s.textTransform, ti: s.textIndent };
};
window.frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 40))));
`;

visual("visual DOM output: PPTX host typography does not leak into inherited text", async () => {
  const pkg = await OpcPackage.open(await Bun.file(PPTX_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const { slide } = await findCoverageSlide(slides, "rtl-table");
  const module = await browserModule(`${import.meta.dir}/../src/pptx/renderer.ts`);
  const preflight = await preflightCss();
  for (const usePreflight of [false, true]) {
    const page = `<!doctype html><style>
html,body{margin:0}
${usePreflight ? preflight : ""}
${HOSTILE_CSS}
#ancestor{width:1280px}
#host{width:1280px}
</style>
<div id="ancestor"><main id="host"></main><p id="unprotected">unprotected inheritance probe</p></div>
<script type="module">
import { render } from "/pptx.mjs";
window.rendered = null;
window.renderFixture = async (model) => { window.rendered?.destroy(); window.rendered = await render({ slide: model, target: document.querySelector("#host") }); };
${SNAPSHOT_SCRIPT}
</script>`;
    await withVisualPage({ modules: { "/pptx.mjs": module }, page, width: 1280, height: 900 }, async (view) => {
      await view.evaluate(`setHostile(false)`);
      await view.evaluate(`renderFixture(${JSON.stringify(slide)})`);
      await view.evaluate(`frame()`);
      const clean = await view.evaluate(`snapshot()`) as TextSnapshot[];
      expect(clean.some((text) => text.fw === "700")).toBe(true);
      expect(clean.some((text) => text.fw === "400")).toBe(true);
      await view.evaluate(`setHostile(true)`);
      await view.evaluate(`renderFixture(${JSON.stringify(slide)})`);
      await view.evaluate(`frame()`);
      const hostile = await view.evaluate(`snapshot()`) as TextSnapshot[];
      expect(hostile.length).toBeGreaterThan(0);
      expect(hostile).toEqual(clean);
      const nc = await view.evaluate(`negativeControl()`) as Record<string, string>;
      expect(nc.fw).toBe("700");
      expect(nc.fs).toBe("italic");
      expect(nc.tt).toBe("uppercase");
    });
  }
}, 60_000);

visual("visual DOM output: XLSX host typography does not leak into inherited text", async () => {
  const pkg = await OpcPackage.open(await Bun.file(XLSX_PATH).arrayBuffer());
  const workbook = await readWorkbook(pkg);
  const sheets = await Promise.all(workbook.sheets.map((_, index) => parseXlsxSheet(pkg, workbook, index)));
  const target = (await findCoverageSheet(sheets, "rtl-grid")).sheet;
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/renderer.ts`);
  const preflight = await preflightCss();
  for (const usePreflight of [false, true]) {
    const page = `<!doctype html><style>
html,body{margin:0}
${usePreflight ? preflight : ""}
${HOSTILE_CSS}
#ancestor{width:1280px}
#host{width:1280px}
</style>
<div id="ancestor"><main id="host"></main><p id="unprotected">unprotected inheritance probe</p></div>
<script type="module">
import { render } from "/xlsx.mjs";
window.rendered = null;
window.renderFixture = async (model) => { window.rendered?.destroy(); window.rendered = await render({ sheet: model, target: document.querySelector("#host") }); };
${SNAPSHOT_SCRIPT}
</script>`;
    await withVisualPage({ modules: { "/xlsx.mjs": module }, page, width: 1280, height: 900 }, async (view) => {
      await view.evaluate(`setHostile(false)`);
      await view.evaluate(`renderFixture(${JSON.stringify(target)})`);
      await view.evaluate(`frame()`);
      const clean = await view.evaluate(`snapshot()`) as TextSnapshot[];
      expect(clean.some((text) => text.fw === "700")).toBe(true);
      expect(clean.some((text) => text.fw === "400")).toBe(true);
      await view.evaluate(`setHostile(true)`);
      await view.evaluate(`renderFixture(${JSON.stringify(target)})`);
      await view.evaluate(`frame()`);
      const hostile = await view.evaluate(`snapshot()`) as TextSnapshot[];
      expect(hostile.length).toBeGreaterThan(0);
      expect(hostile).toEqual(clean);
      const nc = await view.evaluate(`negativeControl()`) as Record<string, string>;
      expect(nc.fw).toBe("700");
      expect(nc.fs).toBe("italic");
      expect(nc.tt).toBe("uppercase");
    });
  }
}, 60_000);

visual("visual DOM output: DOCX host typography does not leak into flow or layout text", async () => {
  const pkg = await OpcPackage.open(await Bun.file(DOCX_PATH).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const module = await browserModule(`${import.meta.dir}/../src/docx/renderer.ts`);
  const preflight = await preflightCss();
  const payload = JSON.stringify(sections);
  for (const usePreflight of [false, true]) {
    const page = `<!doctype html><style>
html,body{margin:0}
${usePreflight ? preflight : ""}
${HOSTILE_CSS}
#ancestor{width:1280px}
#host{width:1280px}
</style>
<div id="ancestor"><main id="host"></main><p id="unprotected">unprotected inheritance probe</p></div>
<script type="module">
import { render } from "/docx.mjs";
window.rendered = null;
window.renderFlow = async (model) => { window.rendered?.destroy(); window.rendered = await render({ document: model, target: document.querySelector("#host") }); };
window.renderLayout = async (model) => { window.rendered?.destroy(); window.rendered = await render({ document: model, target: document.querySelector("#host"), pagination: { mode: "layout" }, zoom: 1 }); };
${SNAPSHOT_SCRIPT}
</script>`;
    await withVisualPage({ modules: { "/docx.mjs": module }, page, width: 1280, height: 1000 }, async (view) => {
      for (const mode of ["flow", "layout"] as const) {
        const fn = mode === "flow" ? "renderFlow" : "renderLayout";
        await view.evaluate(`setHostile(false)`);
        await view.evaluate(`${fn}(${payload})`);
        await view.evaluate(`frame()`);
        const clean = await view.evaluate(`snapshot()`) as TextSnapshot[];
        expect(clean.some((text) => text.fw === "700")).toBe(true);
        expect(clean.some((text) => text.fw === "400")).toBe(true);
        await view.evaluate(`setHostile(true)`);
        await view.evaluate(`${fn}(${payload})`);
        await view.evaluate(`frame()`);
        const hostile = await view.evaluate(`snapshot()`) as TextSnapshot[];
        expect(hostile.length).toBeGreaterThan(0);
        expect(hostile).toEqual(clean);
      }
      const nc = await view.evaluate(`negativeControl()`) as Record<string, string>;
      expect(nc.fw).toBe("700");
      expect(nc.fs).toBe("italic");
      expect(nc.tt).toBe("uppercase");
    });
  }
}, 120_000);
