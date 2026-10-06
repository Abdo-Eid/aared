// Calibrates each ARABIC_SCRIPT_FALLBACKS face the way it is actually painted: under
// the requested family's alias, next to its Latin substitute. The Arabic face only
// covers ARABIC_UNICODE_RANGE, so the spaces, Western digits and punctuation inside
// Arabic text come from the Latin face (Arimo, Tinos, Caladea, Carlito) at its own
// size. Measuring the Arabic substitute on its own (arabic-survey.ts --calib) counts
// its own spaces instead, and for Arial that left Arabic ~4% too wide.
//
// Fit: the combined alias's width is linear in the Arabic face's size-adjust, so two
// probes (100% and 80%) solve for the size-adjust that matches the reference's total
// width over every Arabic-only run in the fixtures. Vertical overrides are the
// reference's own ascent/descent (read at 10000px, not rounded) divided by it.
// Verify: every real paragraph laid out in the DOM at 11pt across a width sweep, in
// the real Office font and in the shipped vs the refitted alias, line starts compared.
//
//   bun tools/font-calibration/arabic-fallback-fit.ts
import { OpcPackage } from "../../../packages/aared/src/core/opc.ts";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.ts";
import { indexPptx, parsePptxSlide } from "../../../packages/aared/src/core/pptx.ts";
import { ARABIC_SCRIPT_FALLBACKS, ARABIC_UNICODE_RANGE, FONT_CALIBRATION, FONT_SUBSTITUTIONS, parseFontCss } from "../../../packages/aared-fonts/src/index.ts";

const ROOT = "D:/work/aared";
const FIT_SIZE = 40;
const SIZE = 14.67; // 11pt, for the line-break check
const WIDTHS = [240, 320, 400, 480, 560, 640, 720];
// The real Office file whose Arabic each fallback must match. Cambria has no Arabic
// glyphs and Word draws that text in Times New Roman (word-arabic-fallback.ps1);
// Aptos has none either, and Excel and Word draw it in Arial (aptos-arabic-fallback.ps1).
const REFERENCES: Record<string, { 400: string; 700?: string }> = {
  Arial: { 400: "arial.ttf", 700: "arialbd.ttf" },
  Aptos: { 400: "arial.ttf", 700: "arialbd.ttf" },
  "Aptos Display": { 400: "arial.ttf", 700: "arialbd.ttf" },
  "Aptos Narrow": { 400: "arial.ttf", 700: "arialbd.ttf" },
  "Times New Roman": { 400: "times.ttf", 700: "timesbd.ttf" },
  Cambria: { 400: "times.ttf", 700: "timesbd.ttf" },
  Calibri: { 400: "calibri.ttf", 700: "calibrib.ttf" },
  "Calibri Light": { 400: "calibril.ttf" },
};

// --- samples ----------------------------------------------------------------------
const ARABIC = /[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]/;
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
    // The fit set: every Arabic-only run (arabic-survey.ts's definition).
    if (typeof record.text === "string" && ARABIC.test(record.text) && !/[A-Za-z]/.test(record.text)) {
      const text = record.text.trim();
      if (text.length >= 12) runs.add(text);
    }
    // The check set: whole paragraphs containing Arabic (Arabic-only and mixed).
    if (Array.isArray(record.runs) && record.runs.every((run) => typeof (run as { text?: unknown })?.text === "string")) {
      const text = (record.runs as Array<{ text: string }>).map((run) => run.text).join("").replace(/\s+/g, " ").trim();
      if (text.length >= 15 && ARABIC.test(text)) paragraphs.add(text);
    }
    for (const name in record) visit(record[name]);
  };
  visit(model);
};
const docx = await OpcPackage.open(await Bun.file(`${ROOT}/fixtures/docx-coverage.docx`).arrayBuffer());
collect(await parseDocxSections(docx, await indexDocx(docx)));
for (const name of ["pptx-coverage.pptx", "post-production-review.pptx"]) {
  const pkg = await OpcPackage.open(await Bun.file(`${ROOT}/fixtures/${name}`).arrayBuffer());
  const deck = await indexPptx(pkg);
  for (let i = 0; i < deck.slideParts.length; i++) collect(await parsePptxSlide(pkg, deck, i));
}
const RUNS = [...runs];
const PARAGRAPHS = [...paragraphs].map((text) => ({ text, kind: /[A-Za-z]/.test(text) ? "mixed" : "arabic" }));
console.log(`${RUNS.length} Arabic-only runs to fit on; ${PARAGRAPHS.length} paragraphs with Arabic to check line breaks on\n`);

// --- the exact Google Fonts files the resolver delivers ------------------------------
async function faces(family: string) {
  const url = `https://fonts.googleapis.com/css2?family=${encodeURIComponent(family).replace(/%20/g, "+")}:wght@400;700&display=swap`;
  const response = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36" } });
  if (!response.ok) throw new Error(`${family}: CSS ${response.status}`);
  return parseFontCss(await response.text()).filter((face) => face.style === "normal");
}
const families = Object.keys(ARABIC_SCRIPT_FALLBACKS);
const fontFaces: Record<string, Awaited<ReturnType<typeof faces>>> = {};
for (const family of families) {
  for (const name of [FONT_SUBSTITUTIONS[family]!, ARABIC_SCRIPT_FALLBACKS[family]!.substitute]) fontFaces[name] ??= await faces(name);
}
const config = {
  fitSize: FIT_SIZE,
  size: SIZE,
  widths: WIDTHS,
  runs: RUNS,
  paragraphs: PARAGRAPHS,
  arabicRange: ARABIC_UNICODE_RANGE,
  fontFaces,
  setups: families.map((family) => ({
    family,
    reference: REFERENCES[family]!,
    latin: FONT_SUBSTITUTIONS[family]!,
    arabic: ARABIC_SCRIPT_FALLBACKS[family]!.substitute,
    shipped: ARABIC_SCRIPT_FALLBACKS[family]!.calibration,
    // A calibrated Latin substitute (Aptos's) is registered with its own numbers,
    // so its spaces paint at that size: fit through the alias as it really is.
    latinCalibration: FONT_CALIBRATION[family] ?? null,
  })),
};

const server = Bun.serve({
  port: 0,
  async fetch(request) {
    const path = new URL(request.url).pathname;
    if (path.startsWith("/ref/")) return new Response(Bun.file(`C:/Windows/Fonts/${path.slice(5)}`), { headers: { "content-type": "font/ttf" } });
    return new Response(`<!doctype html><meta charset="utf-8"><body style="margin:0"><script>window.CFG = ${JSON.stringify(config)};</script></body>`, { headers: { "content-type": "text/html; charset=utf-8" } });
  },
});

try {
  await using view = new Bun.WebView({ width: 1400, height: 900 });
  await view.navigate(server.url.href);
  const raw = (await view.evaluate(`(async () => {
    const { fitSize, size, widths, runs, paragraphs, arabicRange, fontFaces, setups } = window.CFG;
    const allText = runs.concat(paragraphs.map((p) => p.text)).join(" ");
    const isArabicSubset = (f) => /U\\+0600-06FF/.test(f.unicodeRange);
    const registered = [];
    const add = (alias, url, weight, range, o) => {
      const d = { weight: String(weight) };
      if (range) d.unicodeRange = range;
      if (o) { d.sizeAdjust = o.sizeAdjustPercent + "%"; d.ascentOverride = o.ascentOverridePercent + "%"; d.descentOverride = o.descentOverridePercent + "%"; }
      const face = new FontFace(alias, "url(" + url + ")", d);
      document.fonts.add(face); registered.push({ alias, face, arabic: range === arabicRange });
    };
    const load = async (alias, weight) => document.fonts.load(weight + " " + fitSize + "px '" + alias + "'", allText);
    const ctx = document.createElement("canvas").getContext("2d");
    const widthsOf = (alias, weight, texts) => { ctx.font = weight + " " + fitSize + "px '" + alias + "'"; return texts.map((t) => ctx.measureText(t).width); };
    const sum = (xs) => xs.reduce((a, b) => a + b, 0);
    const round1 = (x) => Math.round(x * 10) / 10;
    // One alias = the Latin substitute's faces as delivered (with its own calibration, if any)
    // + the Arabic substitute's Arabic face.
    const pick = (c, w) => (c && w >= 600 && c.bold ? c.bold : c);
    const alias = (name, setup, weight, overrides) => {
      for (const f of fontFaces[setup.latin].filter((f) => f.weight === String(weight))) add(name, f.url, weight, f.unicodeRange, pick(setup.latinCalibration, weight) ?? undefined);
      for (const f of fontFaces[setup.arabic].filter((f) => f.weight === String(weight) && isArabicSubset(f))) add(name, f.url, weight, arabicRange, overrides);
    };

    const out = { setups: {}, missing: [] };
    for (const setup of setups) {
      const result = out.setups[setup.family] = { fit: {}, check: {} };
      for (const weight of [400, 700]) {
        const file = setup.reference[weight];
        if (!file) continue;
        const ref = "ref:" + setup.family + ":" + weight;
        add(ref, "/ref/" + file, weight);
        await load(ref, weight);
        // Reference vertical metrics at 10000px: canvas rounds them to whole pixels.
        ctx.font = weight + " 10000px '" + ref + "'";
        const m = ctx.measureText("Hx");
        const ascent = m.fontBoundingBoxAscent / 10000, descent = m.fontBoundingBoxDescent / 10000;
        // Two probes through the alias; width is linear in the Arabic face's size-adjust.
        const probe = {};
        for (const s of [100, 80]) {
          const name = "probe:" + setup.family + ":" + weight + ":" + s;
          alias(name, setup, weight, { sizeAdjustPercent: s, ascentOverridePercent: 100, descentOverridePercent: 30 });
          await load(name, weight);
          probe[s] = sum(widthsOf(name, weight, runs));
        }
        const target = sum(widthsOf(ref, weight, runs));
        const scale = 1 + (target - probe[100]) / ((probe[100] - probe[80]) / 0.2);
        const fitted = {
          sizeAdjustPercent: round1(scale * 100),
          ascentOverridePercent: round1(ascent / scale * 100),
          descentOverridePercent: round1(descent / scale * 100),
        };
        result.fit[weight] = fitted;

        // Shipped vs refitted: per-run spread through the alias, and line breaks.
        const shippedOverrides = weight === 700 && setup.shipped.bold ? setup.shipped.bold : setup.shipped;
        const variants = { shipped: shippedOverrides, fitted };
        for (const [label, overrides] of Object.entries(variants)) {
          const name = label + ":" + setup.family + ":" + weight;
          alias(name, setup, weight, overrides);
          await load(name, weight);
        }
        const refWidths = widthsOf(ref, weight, runs);
        const box = document.createElement("div");
        box.style.cssText = "position:absolute;left:0;top:0;white-space:normal;line-height:1.15;font-size:" + size + "px;font-weight:" + weight;
        document.body.append(box);
        const layout = (family, text, width) => {
          box.style.fontFamily = "'" + family + "'";
          box.style.width = width + "px";
          box.dir = "rtl";
          box.textContent = text;
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
        const refLayouts = paragraphs.map((p) => widths.map((w) => layout(ref, p.text, w)));
        for (const [label] of Object.entries(variants)) {
          const name = label + ":" + setup.family + ":" + weight;
          const ratios = widthsOf(name, weight, runs).map((w, i) => w / refWidths[i]);
          const mean = sum(widthsOf(name, weight, runs)) / sum(refWidths);
          const sd = Math.sqrt(sum(ratios.map((x) => (x - mean) ** 2)) / ratios.length);
          const lines = { arabic: [0, 0], mixed: [0, 0] };
          paragraphs.forEach((p, pi) => widths.forEach((w, wi) => {
            const want = new Set(refLayouts[pi][wi]);
            const got = layout(name, p.text, w);
            lines[p.kind][0] += got.filter((x) => want.has(x)).length;
            lines[p.kind][1] += want.size;
          }));
          (result.check[weight] ??= {})[label] = { mean, sd, arabicLines: lines.arabic[0] / lines.arabic[1], mixedLines: lines.mixed[0] / lines.mixed[1] };
        }
        box.remove();
      }
    }
    // Every alias needs its Latin face and its Arabic face actually loaded.
    const byAlias = {};
    for (const r of registered) if (r.face.status === "loaded") byAlias[r.alias] = (byAlias[r.alias] ?? 0) | (r.arabic ? 2 : 1);
    for (const r of registered) if (!r.alias.startsWith("ref:") && byAlias[r.alias] !== 3 && !out.missing.includes(r.alias)) out.missing.push(r.alias);
    for (const r of registered) if (r.alias.startsWith("ref:") && r.face.status !== "loaded") out.missing.push(r.alias);
    return JSON.stringify(out);
  })()`)) as string;
  const data = JSON.parse(raw);
  await Bun.write(`${ROOT}/tools/font-calibration/arabic-fallback-fit.last.json`, JSON.stringify({ runs: RUNS, paragraphs: PARAGRAPHS, data }, null, 1));
  if (data.missing.length > 0) console.log("!! aliases missing a loaded face (their numbers are not trustworthy):", data.missing);

  const pct = (x: number): string => `${(x * 100).toFixed(1)}%`;
  for (const [family, result] of Object.entries(data.setups) as Array<[string, any]>) {
    console.log(`== ${family}`);
    for (const weight of Object.keys(result.fit)) {
      const f = result.fit[weight];
      console.log(`  ${weight} fitted ${f.sizeAdjustPercent}% / ${f.ascentOverridePercent}% / ${f.descentOverridePercent}%`);
      for (const [label, c] of Object.entries(result.check[weight]) as Array<[string, any]>) {
        console.log(`      ${label.padEnd(8)} width vs original ${pct(c.mean)}  sd ${c.sd.toFixed(4)}  lines like the original: Arabic ${pct(c.arabicLines)}, mixed ${pct(c.mixedLines)}`);
      }
    }
  }
} finally {
  server.stop();
}
