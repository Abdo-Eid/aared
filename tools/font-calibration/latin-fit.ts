// Line-break acceptance for an Aptos substitute. For each reference (the real
// Aptos, Aptos Display, Aptos Narrow from Office's cloud-font cache) and each
// shortlisted candidate: fit size-adjust on the total width of every Latin run
// in the fixtures (a single face, so width is linear in it), register the
// candidate the way the resolver does (every delivered subset under one alias,
// size-adjust + ascent/descent overrides), then lay out real paragraphs in the
// DOM across a width sweep and count line starts that land where the real font
// puts them. "Arimo as-is" is the baseline: on Windows today, an Aptos document
// paints in the generic sans-serif, which is Arial there (Arimo's metrics).
//
//   bun tools/font-calibration/latin-fit.ts [--refs a,b] [--cands x,y] [--shipped]
// --shipped also checks the numbers in FONT_CALIBRATION against a fresh fit.
import { readdirSync } from "node:fs";
import { OpcPackage } from "../../../packages/aared/src/core/opc.ts";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.ts";
import { indexPptx, parsePptxSlide } from "../../../packages/aared/src/core/pptx.ts";
import { readWorkbook } from "../../../packages/aared/src/core/xlsx/workbook.ts";
import { parseXlsxSheet } from "../../../packages/aared/src/core/xlsx/sheet.ts";
import { FONT_CALIBRATION, FONT_SUBSTITUTIONS, fontCssUrl, parseFontCss } from "../../../packages/aared-fonts/src/index.ts";
import { cloudFontFile } from "./cloud-fonts.ts";

const arg = (name: string): string[] | undefined => {
  const at = process.argv.indexOf(name);
  return at === -1 ? undefined : process.argv[at + 1]!.split(",");
};
const FIT_SIZE = 40;
const WIDTHS = [160, 240, 320, 400, 480, 560, 640, 720];
// Body text at 11pt, a spreadsheet cell at 11pt, slide headings larger.
const SIZES: Record<string, number> = { Aptos: 14.67, "Aptos Narrow": 14.67, "Aptos Display": 24 };
const REFS = arg("--refs") ?? ["Aptos", "Aptos Display", "Aptos Narrow"];
const CANDS = arg("--cands") ?? ["Arimo"];
const SHIPPED = process.argv.includes("--shipped");

const refFiles: Record<string, string> = {};
for (const family of REFS) {
  refFiles[`${family}:400`] = await cloudFontFile(family, family);
  refFiles[`${family}:700`] = await cloudFontFile(family, `${family} Bold`);
}

// --- samples ----------------------------------------------------------------------
const ARABIC = /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/;
const runs = new Set<string>();
const paragraphs = new Set<string>();
const collect = (model: unknown): void => {
  const seen = new Set<object>();
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object" || seen.has(value)) return;
    seen.add(value);
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    for (const key of ["text", "formatted"]) {
      const raw = record[key];
      if (typeof raw === "string" && !ARABIC.test(raw) && /[A-Za-z]/.test(raw) && raw.trim().length >= 8) runs.add(raw.trim());
    }
    if (Array.isArray(record.runs) && record.runs.every((run) => typeof (run as { text?: unknown })?.text === "string")) {
      const text = (record.runs as Array<{ text: string }>).map((run) => run.text).join("").replace(/\s+/g, " ").trim();
      if (text.length >= 30 && !ARABIC.test(text)) paragraphs.add(text);
    }
    if (typeof record.formatted === "string" && record.formatted.length >= 30 && !ARABIC.test(record.formatted)) paragraphs.add(record.formatted.trim());
    for (const name in record) visit(record[name]);
  };
  visit(model);
};
for (const name of readdirSync("fixtures")) {
  const open = async () => OpcPackage.open(await Bun.file(`fixtures/${name}`).arrayBuffer());
  if (name.endsWith(".docx")) {
    const doc = await open();
    collect(await parseDocxSections(doc, await indexDocx(doc)));
  } else if (name.endsWith(".pptx")) {
    const deck = await open();
    const index = await indexPptx(deck);
    for (let i = 0; i < index.slideParts.length; i++) collect(await parsePptxSlide(deck, index, i));
  } else if (name.endsWith(".xlsx")) {
    const book = await open();
    const workbook = await readWorkbook(book);
    for (let i = 0; i < workbook.sheets.length; i++) collect(await parseXlsxSheet(book, workbook, i));
  }
}
// The fixtures are mostly Arabic, so their English paragraphs are few; these add
// ordinary report prose (written for this probe) so line breaks are counted on
// enough text: capitals, digits, punctuation and long and short words.
const PROSE = [
  "The quarterly review covers production spend, delivery dates and the open risks raised by each department since the last meeting in March.",
  "Revenue grew 14.2% year on year to $3,481,900, driven mainly by licensing; operating costs rose more slowly, so the margin widened by 2.6 points.",
  "Please confirm the final schedule with the post-production team before Friday, including colour grading, sound mix and the two remaining VFX shots.",
  "Figure 3 shows weekly throughput for all four studios. Output dipped in week 31 when the render farm was offline for maintenance, then recovered.",
  "Key assumptions: headcount stays flat, exchange rates hold at today's levels, and no new market launches before the second half of next year.",
  "An independent auditor reviewed the accounts and found no material misstatements; three minor recommendations are listed in Appendix B.",
  "Quality, Wellbeing & Growth: we introduced flexible hours, a mentoring programme and quarterly workshops open to every member of staff.",
  "If the vendor cannot meet the revised deadline of 30 September, we will move the remaining work in-house and reallocate the budget accordingly.",
];
for (const text of PROSE) paragraphs.add(text);
const RUNS = [...runs];
const PARAGRAPHS = [...paragraphs];
console.log(`${RUNS.length} Latin runs to fit on; ${PARAGRAPHS.length} paragraphs (${PROSE.length} of them written prose) to check line breaks on\n`);

// --- the resolver's exact faces ----------------------------------------------------
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36";
const families = new Set([...CANDS, ...(SHIPPED ? REFS.map((ref) => FONT_SUBSTITUTIONS[ref]).filter((x): x is string => x !== undefined) : [])]);
const fontFaces: Record<string, ReturnType<typeof parseFontCss>> = {};
for (const family of families) {
  const response = await fetch(fontCssUrl([family]), { headers: { "User-Agent": UA } });
  if (!response.ok) throw new Error(`${family}: CSS ${response.status}`);
  fontFaces[family] = parseFontCss(await response.text()).filter((face) => face.family === family && face.style === "normal");
}
const shipped = SHIPPED ? Object.fromEntries(REFS.flatMap((ref) => {
  const substitute = FONT_SUBSTITUTIONS[ref];
  return substitute === undefined ? [] : [[ref, { substitute, calibration: FONT_CALIBRATION[ref] }]];
})) : {};

const config = { fitSize: FIT_SIZE, sizes: SIZES, widths: WIDTHS, runs: RUNS, paragraphs: PARAGRAPHS, refs: REFS, cands: CANDS, fontFaces, shipped };
const server = Bun.serve({
  port: 0,
  fetch(request) {
    const path = decodeURIComponent(new URL(request.url).pathname);
    if (path.startsWith("/ref/")) return new Response(Bun.file(refFiles[path.slice(5)]!), { headers: { "content-type": "font/ttf" } });
    return new Response(`<!doctype html><meta charset="utf-8"><body style="margin:0"><script>window.CFG = ${JSON.stringify(config)};</script></body>`, { headers: { "content-type": "text/html; charset=utf-8" } });
  },
});

try {
  await using view = new Bun.WebView({ width: 1400, height: 900 });
  await view.navigate(server.url.href);
  const raw = (await view.evaluate(`(async () => {
    const { fitSize, sizes, widths, runs, paragraphs, refs, cands, fontFaces, shipped } = window.CFG;
    const allText = runs.concat(paragraphs).join(" ");
    const registered = [];
    const add = (alias, url, weight, range, o) => {
      const d = { weight: String(weight) };
      if (range) d.unicodeRange = range;
      if (o) { d.sizeAdjust = o.sizeAdjustPercent + "%"; d.ascentOverride = o.ascentOverridePercent + "%"; d.descentOverride = o.descentOverridePercent + "%"; }
      const face = new FontFace(alias, "url(" + url + ")", d);
      document.fonts.add(face); registered.push({ alias, face });
    };
    const load = (alias, weight) => document.fonts.load(weight + " " + fitSize + "px '" + alias + "'", allText);
    const ctx = document.createElement("canvas").getContext("2d");
    const widthsOf = (alias, weight, texts) => { ctx.font = weight + " " + fitSize + "px '" + alias + "'"; return texts.map((t) => ctx.measureText(t).width); };
    const sum = (xs) => xs.reduce((a, b) => a + b, 0);
    const round1 = (x) => Math.round(x * 10) / 10;
    const candidate = (name, family, weight, overrides) => {
      for (const f of fontFaces[family].filter((f) => f.weight === String(weight))) add(name, f.url, weight, f.unicodeRange, overrides);
    };
    const box = document.createElement("div");
    box.style.cssText = "position:absolute;left:0;top:0;white-space:normal;line-height:1.2";
    document.body.append(box);
    const layout = (family, weight, size, text, width) => {
      box.style.fontFamily = "'" + family + "'"; box.style.fontWeight = weight; box.style.fontSize = size + "px";
      box.style.width = width + "px"; box.textContent = text;
      const node = box.firstChild, range = document.createRange(), starts = [];
      let lastTop = -1e9, word = 0;
      for (let i = 0; i < text.length; i++) {
        if (text[i] === " " || (i > 0 && text[i - 1] !== " ")) continue;
        range.setStart(node, i); range.setEnd(node, i + 1);
        const rect = range.getClientRects()[0];
        if (rect && rect.top > lastTop + size * 0.5) { starts.push(word); lastTop = rect.top; }
        word++;
      }
      return starts;
    };
    const out = { refs: {}, missing: [] };
    for (const ref of refs) {
      const result = out.refs[ref] = {};
      for (const weight of [400, 700]) {
        const refName = "ref:" + ref + ":" + weight;
        add(refName, "/ref/" + encodeURIComponent(ref + ":" + weight), weight);
        await load(refName, weight);
        ctx.font = weight + " 10000px '" + refName + "'";
        const m = ctx.measureText("Hx");
        const ascent = m.fontBoundingBoxAscent / 10000, descent = m.fontBoundingBoxDescent / 10000;
        const refWidths = widthsOf(refName, weight, runs);
        const size = sizes[ref];
        const refLayouts = paragraphs.map((p) => widths.map((w) => layout(refName, weight, size, p, w)));
        const variants = [];
        for (const cand of cands) {
          const probe = "probe:" + ref + ":" + cand + ":" + weight;
          candidate(probe, cand, weight);
          await load(probe, weight);
          const scale = sum(refWidths) / sum(widthsOf(probe, weight, runs));
          const fitted = { sizeAdjustPercent: round1(scale * 100), ascentOverridePercent: round1(ascent / scale * 100), descentOverridePercent: round1(descent / scale * 100) };
          variants.push({ label: cand + " fitted", family: cand, overrides: fitted });
          if (cand === "Arimo") variants.push({ label: "Arimo as-is", family: cand, overrides: undefined });
        }
        const ship = shipped[ref];
        if (ship) {
          const o = weight >= 600 && ship.calibration && ship.calibration.bold ? ship.calibration.bold : ship.calibration;
          variants.push({ label: ship.substitute + " shipped", family: ship.substitute, overrides: o });
        }
        for (const v of variants) {
          const name = "v:" + ref + ":" + v.label + ":" + weight;
          candidate(name, v.family, weight, v.overrides);
          await load(name, weight);
          const ws = widthsOf(name, weight, runs);
          const ratios = ws.map((w, i) => w / refWidths[i]);
          const mean = sum(ws) / sum(refWidths);
          const sd = Math.sqrt(sum(ratios.map((x) => (x - mean) ** 2)) / ratios.length);
          let same = 0, total = 0, identical = 0, layouts = 0;
          paragraphs.forEach((p, pi) => widths.forEach((w, wi) => {
            const want = refLayouts[pi][wi];
            const got = layout(name, weight, size, p, w);
            const wantSet = new Set(want);
            same += got.filter((x) => wantSet.has(x)).length; total += Math.max(want.length, got.length);
            layouts++; if (got.length === want.length && got.every((x, i) => x === want[i])) identical++;
          }));
          // Vertical: the line box at line-height normal, against the real font's.
          box.style.lineHeight = "normal"; box.style.width = "auto"; box.style.whiteSpace = "nowrap";
          box.style.fontSize = "1000px"; box.style.fontWeight = weight; box.textContent = "Hxgy";
          box.style.fontFamily = "'" + refName + "'"; const refBox = box.getBoundingClientRect().height;
          box.style.fontFamily = "'" + name + "'"; const box2 = box.getBoundingClientRect().height;
          box.style.lineHeight = "1.2"; box.style.whiteSpace = "normal";
          (result[weight] ??= []).push({ label: v.label, overrides: v.overrides ?? null, mean, sd, lineStarts: same / total, identical: identical / layouts, lineBox: box2 / refBox });
        }
      }
    }
    for (const r of registered) if (r.face.status !== "loaded" && r.face.status !== "unloaded") out.missing.push(r.alias + " " + r.face.status);
    for (const alias of new Set(registered.map((r) => r.alias))) if (!registered.some((r) => r.alias === alias && r.face.status === "loaded")) out.missing.push(alias + " (nothing loaded)");
    return JSON.stringify(out);
  })()`)) as string;
  const data = JSON.parse(raw);
  await Bun.write("tools/font-calibration/latin-fit.last.json", JSON.stringify({ runs: RUNS, paragraphs: PARAGRAPHS, data }, null, 1));
  if (data.missing.length > 0) console.log("!! faces that did not load (their numbers are not trustworthy):", data.missing);
  const pct = (x: number): string => `${(x * 100).toFixed(1)}%`;
  for (const [ref, result] of Object.entries(data.refs) as Array<[string, any]>) {
    console.log(`== ${ref} (${SIZES[ref]}px, widths ${WIDTHS.join("/")})`);
    for (const weight of Object.keys(result)) {
      const rows = [...result[weight]].sort((a: any, b: any) => b.identical - a.identical);
      for (const row of rows) {
        const o = row.overrides;
        console.log(`  ${weight} ${row.label.padEnd(34)} ${o ? `${o.sizeAdjustPercent}% / ${o.ascentOverridePercent}% / ${o.descentOverridePercent}%` : "no overrides"}`.padEnd(80)
          + ` width ${pct(row.mean)} sd ${row.sd.toFixed(4)}  paragraphs broken identically ${pct(row.identical)}  line starts ${pct(row.lineStarts)}  line box ${pct(row.lineBox)}`);
      }
    }
  }
} finally {
  server.stop();
}
