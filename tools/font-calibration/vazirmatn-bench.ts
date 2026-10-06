// Vazirmatn vs the shipped Arial substitute (Arimo + calibrated Lateef), measured
// the way a reader sees it: every real paragraph in the fixtures is laid out in
// the DOM at 11pt across a sweep of widths, in the real Arial font file and in
// each substitute setup, and the line breaks are compared word for word.
//
// Setups (each one a private alias, so no installed font can stand in):
//   current   Arimo (no overrides) + Lateef Arabic face with LATEEF_FOR_ARIAL, as shipped
//   current*  the same, Lateef's size refitted through the combined alias
//   vazir     Vazirmatn for both scripts, the way aared would ship it: the Latin and
//             Arabic faces split by unicode-range, each with its own size-adjust and
//             ascent/descent overrides fitted to Arial
//   vazir0    Vazirmatn as-is, one family, same size, no overrides
//
//   bun tools/font-calibration/vazirmatn-bench.ts
import { OpcPackage } from "../../../packages/aared/src/core/opc.ts";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.ts";
import { indexPptx, parsePptxSlide } from "../../../packages/aared/src/core/pptx.ts";
import { ARABIC_UNICODE_RANGE, ARABIC_SCRIPT_FALLBACKS, parseFontCss } from "../../../packages/aared-fonts/src/index.ts";

const ROOT = "D:/work/aared";
const SIZE = 14.67; // 11pt
const WIDTHS = [240, 320, 400, 480, 560, 640, 720];
const WEIGHTS = [400, 700];
const LATEEF = ARABIC_SCRIPT_FALLBACKS.Arial!.calibration;

// --- samples: whole paragraphs of real text -------------------------------------
const ARABIC = /[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]/;
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
    if (Array.isArray(record.runs) && record.runs.every((run) => typeof (run as { text?: unknown })?.text === "string")) {
      const text = (record.runs as Array<{ text: string }>).map((run) => run.text).join("").replace(/\s+/g, " ").trim();
      if (text.length >= 15) paragraphs.add(text);
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
const kindOf = (text: string): "english" | "arabic" | "mixed" => {
  const ar = ARABIC.test(text), la = /[A-Za-z]/.test(text);
  return ar && la ? "mixed" : ar ? "arabic" : "english";
};
const SAMPLES = [...paragraphs].map((text) => ({ text, kind: kindOf(text) }));
const count = (kind: string): number => SAMPLES.filter((s) => s.kind === kind).length;
console.log(`${SAMPLES.length} paragraphs: ${count("english")} English, ${count("arabic")} Arabic, ${count("mixed")} mixed\n`);

// --- the exact Google Fonts files the resolver would deliver --------------------
async function faces(family: string) {
  const url = `https://fonts.googleapis.com/css2?family=${encodeURIComponent(family).replace(/%20/g, "+")}:wght@400;700&display=swap`;
  const response = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36" } });
  if (!response.ok) throw new Error(`${family}: CSS ${response.status}`);
  return parseFontCss(await response.text()).filter((face) => face.style === "normal");
}
const config = {
  size: SIZE,
  widths: WIDTHS,
  weights: WEIGHTS,
  samples: SAMPLES,
  arabicRange: ARABIC_UNICODE_RANGE,
  lateef: LATEEF,
  fonts: { arimo: await faces("Arimo"), lateef: await faces("Lateef"), vazir: await faces("Vazirmatn") },
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
    const { size, widths, weights, samples, arabicRange, lateef, fonts } = window.CFG;
    const all = samples.map((s) => s.text).join(" ");
    const english = samples.filter((s) => s.kind === "english").map((s) => s.text);
    const arabic = samples.filter((s) => s.kind === "arabic").map((s) => s.text);
    const isArabicSubset = (f) => /U\\+0600-06FF/.test(f.unicodeRange);
    const registered = [];
    const add = (alias, url, weight, range, adjust) => {
      const d = { weight: String(weight) };
      if (range) d.unicodeRange = range;
      if (adjust) { d.sizeAdjust = adjust.s + "%"; d.ascentOverride = adjust.a + "%"; d.descentOverride = adjust.d + "%"; }
      const face = new FontFace(alias, "url(" + url + ")", d);
      document.fonts.add(face); registered.push({ alias, face, url });
    };
    const loadAlias = async (alias) => { for (const w of weights) await document.fonts.load(w + " " + size + "px '" + alias + "'", all); };
    const ctx = document.createElement("canvas").getContext("2d");
    const widthOf = (alias, weight, texts) => { ctx.font = weight + " " + size + "px '" + alias + "'"; return texts.map((t) => ctx.measureText(t).width); };
    const sum = (xs) => xs.reduce((a, b) => a + b, 0);
    const byWeight = (list, w) => list.filter((f) => f.weight === String(w));

    // Reference: the real Arial files.
    add("ref", "/ref/arial.ttf", 400); add("ref", "/ref/arialbd.ttf", 700);
    await loadAlias("ref");
    ctx.font = "400 " + size + "px 'ref'";
    const refMetric = ctx.measureText("Hx");
    const refAscent = refMetric.fontBoundingBoxAscent / size, refDescent = refMetric.fontBoundingBoxDescent / size;
    const over = (s) => ({ s: +(s * 100).toFixed(2), a: +(refAscent / s * 100).toFixed(2), d: +(refDescent / s * 100).toFixed(2) });

    // Fit Vazirmatn's Latin scale on English paragraphs, per weight.
    const fit = { vazirLatin: {}, vazirArabic: {}, lateefArabic: {} };
    for (const w of weights) {
      for (const f of byWeight(fonts.vazir, w)) if (!isArabicSubset(f)) add("fitVL", f.url, w, f.unicodeRange);
    }
    await loadAlias("fitVL");
    for (const w of weights) fit.vazirLatin[w] = sum(widthOf("ref", w, english)) / sum(widthOf("fitVL", w, english));

    // Fit each Arabic face through the combined alias (spaces and digits come from the Latin
    // face): width is linear in the Arabic face's size-adjust, so two probes solve it.
    const latinFaces = { arimo: (w) => byWeight(fonts.arimo, w), vazir: (w) => byWeight(fonts.vazir, w).filter((f) => !isArabicSubset(f)) };
    const arabicFace = { lateef: (w) => byWeight(fonts.lateef, w).filter(isArabicSubset), vazir: (w) => byWeight(fonts.vazir, w).filter(isArabicSubset) };
    const fitArabic = async (latin, latinAdjust, ar, key) => {
      for (const probe of [1, 0.8]) {
        const alias = "fit-" + key + "-" + probe;
        for (const w of weights) {
          for (const f of latinFaces[latin](w)) add(alias, f.url, w, f.unicodeRange, latinAdjust ? over(latinAdjust[w]) : null);
          for (const f of arabicFace[ar](w)) add(alias, f.url, w, arabicRange, { s: probe * 100, a: 100, d: 30 });
        }
        await loadAlias(alias);
      }
      for (const w of weights) {
        const target = sum(widthOf("ref", w, arabic));
        const w1 = sum(widthOf("fit-" + key + "-1", w, arabic)), w8 = sum(widthOf("fit-" + key + "-0.8", w, arabic));
        const slope = (w1 - w8) / 0.2;
        fit[key][w] = 1 + (target - w1) / slope;
      }
    };
    await fitArabic("vazir", fit.vazirLatin, "vazir", "vazirArabic");
    await fitArabic("arimo", null, "lateef", "lateefArabic");

    // The setups.
    for (const w of weights) {
      for (const f of latinFaces.arimo(w)) { add("current", f.url, w, f.unicodeRange); add("current*", f.url, w, f.unicodeRange); }
      for (const f of arabicFace.lateef(w)) {
        add("current", f.url, w, arabicRange, { s: lateef.sizeAdjustPercent, a: lateef.ascentOverridePercent, d: lateef.descentOverridePercent });
        add("current*", f.url, w, arabicRange, over(fit.lateefArabic[w]));
      }
      for (const f of latinFaces.vazir(w)) add("vazir", f.url, w, f.unicodeRange, over(fit.vazirLatin[w]));
      for (const f of arabicFace.vazir(w)) add("vazir", f.url, w, arabicRange, over(fit.vazirArabic[w]));
      for (const f of byWeight(fonts.vazir, w)) add("vazir0", f.url, w, f.unicodeRange);
    }
    const SETUPS = ["current", "current*", "vazir", "vazir0"];
    for (const s of SETUPS) await loadAlias(s);
    // Which faces each alias actually loaded: weight, script (Arabic range or a Latin subset) and
    // the first code point of the range, so a missing Latin or Arabic face shows at a glance.
    const faceStatus = {};
    for (const r of registered) {
      if (r.face.status !== "loaded") continue;
      const range = r.face.unicodeRange;
      (faceStatus[r.alias] ??= []).push(r.face.weight + ":" + (range === arabicRange ? "arabic" : range.split(",")[0].trim()));
    }

    // Line starts, as word indices, for one paragraph laid out in one font at one width.
    const box = document.createElement("div");
    box.style.cssText = "position:absolute;left:0;top:0;white-space:normal;overflow-wrap:normal;word-break:normal;line-height:1.15";
    document.body.append(box);
    const layout = (alias, weight, text, width) => {
      box.style.fontFamily = "'" + alias + "'";
      box.style.fontWeight = weight;
      box.style.fontSize = size + "px";
      box.style.width = width + "px";
      box.dir = ARABIC_RE.test(text) ? "rtl" : "ltr";
      box.textContent = text;
      const node = box.firstChild;
      const range = document.createRange();
      const starts = [];
      let lastTop = -1e9, word = 0;
      for (let i = 0; i < text.length; i++) {
        if (i > 0 && text[i - 1] !== " ") continue;
        if (text[i] === " ") continue;
        range.setStart(node, i); range.setEnd(node, i + 1);
        const rect = range.getClientRects()[0];
        if (rect && rect.top > lastTop + size * 0.5) { starts.push(word); lastTop = rect.top; }
        word++;
      }
      return { starts, height: box.getBoundingClientRect().height };
    };
    const ARABIC_RE = /[\\u0600-\\u06FF]/;

    const results = {};
    for (const s of SETUPS) {
      for (const w of weights) {
        const acc = results[s + "@" + w] = {};
        for (const kind of ["english", "arabic", "mixed"]) acc[kind] = { cases: 0, sameParagraph: 0, lines: 0, sameLines: 0, lineCountOff: 0, heightRatio: [] };
        for (const sample of samples) {
          for (const width of widths) {
            const ref = layout("ref", w, sample.text, width);
            const got = layout(s, w, sample.text, width);
            const a = acc[sample.kind];
            a.cases++;
            const refSet = new Set(ref.starts);
            const same = got.starts.filter((x) => refSet.has(x)).length;
            a.lines += ref.starts.length;
            a.sameLines += same;
            if (same === ref.starts.length && got.starts.length === ref.starts.length) a.sameParagraph++;
            if (got.starts.length !== ref.starts.length) a.lineCountOff++;
            if (got.starts.length === ref.starts.length) a.heightRatio.push(got.height / ref.height);
          }
        }
      }
    }

    // Letter heights after calibration, and the line box at line-height: normal.
    const letters = {};
    for (const alias of ["ref", ...SETUPS]) {
      // Glyph bounds come back whole-pixel, so read letter heights at 100px (as % of the em).
      ctx.font = "400 100px '" + alias + "'";
      const m = (t) => ctx.measureText(t).actualBoundingBoxAscent;
      box.style.lineHeight = "normal"; box.style.width = "2000px";
      box.style.fontFamily = "'" + alias + "'"; box.style.fontWeight = 400;
      box.dir = "ltr"; box.textContent = "Hx";
      const normalEn = box.getBoundingClientRect().height;
      box.dir = "rtl"; box.textContent = "\\u0627\\u0644\\u0633\\u0644\\u0627\\u0645";
      const normalAr = box.getBoundingClientRect().height;
      box.textContent = "Arabic \\u0627\\u0644\\u0633\\u0644\\u0627\\u0645 text";
      const normalMixed = box.getBoundingClientRect().height;
      box.style.lineHeight = "1.15";
      letters[alias] = { cap: m("H"), x: m("x"), alef: m("\\u0627\\u0627\\u0627"), seen: m("\\u0633\\u0633\\u0633\\u0633"), normalEn, normalAr, normalMixed };
    }
    return JSON.stringify({ fit, faceStatus, results, letters, refAscent, refDescent });
  })()`)) as string;
  const data = JSON.parse(raw);
  await Bun.write(`${ROOT}/tools/font-calibration/vazirmatn-bench.last.json`, JSON.stringify(data, null, 1));

  console.log("loaded faces per alias (each setup needs a Latin and an Arabic face at 400 and 700):");
  for (const [alias, loaded] of Object.entries(data.faceStatus)) console.log(`  ${alias.padEnd(20)} ${(loaded as string[]).join("  ")}`);
  console.log(`\nfitted size-adjust: Vazirmatn Latin ${JSON.stringify(data.fit.vazirLatin)}, Vazirmatn Arabic ${JSON.stringify(data.fit.vazirArabic)}, Lateef Arabic ${JSON.stringify(data.fit.lateefArabic)} (shipped Lateef: ${LATEEF.sizeAdjustPercent / 100})`);

  const pct = (a: number, b: number): string => `${((a / b) * 100).toFixed(1)}%`;
  for (const w of WEIGHTS) {
    console.log(`\n== weight ${w}: lines starting at the same word as real Arial (whole paragraphs identical) [line count differs]`);
    for (const setup of ["current", "current*", "vazir", "vazir0"]) {
      const acc = data.results[`${setup}@${w}`];
      const cells = ["english", "arabic", "mixed"].map((kind) => {
        const a = acc[kind];
        const height = a.heightRatio.reduce((x: number, y: number) => x + y, 0) / a.heightRatio.length;
        return `${kind} ${pct(a.sameLines, a.lines)} (${pct(a.sameParagraph, a.cases)}) [${pct(a.lineCountOff, a.cases)}] h${height.toFixed(3)}`;
      });
      console.log(`  ${setup.padEnd(9)} ${cells.join("   ")}`);
    }
  }
  console.log("\nh = paragraph height at line-height 1.15 vs Arial, where the line count agrees");
  console.log("\nletter heights (% of the em) and line box at 11pt, line-height: normal (px):");
  for (const [alias, l] of Object.entries(data.letters) as Array<[string, Record<string, number>]>) {
    console.log(`  ${alias.padEnd(9)} cap ${l.cap!.toFixed(1)} x ${l.x!.toFixed(1)} alef ${l.alef!.toFixed(1)} seen ${l.seen!.toFixed(1)} | normal EN ${l.normalEn!.toFixed(2)} AR ${l.normalAr!.toFixed(2)} mixed ${l.normalMixed!.toFixed(2)}`);
  }
} finally {
  server.stop();
}
