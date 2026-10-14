// Line-box check for every calibrated face: at line-height: normal and 2000px, the
// real Office font's line box against the substitute registered with the shipped
// numbers (and, for the record, the numbers before Notes §46). A main-face
// substitute (FONT_CALIBRATION) is all its subsets under one alias; an Arabic
// fallback (ARABIC_SCRIPT_FALLBACKS) is the Latin substitute plus the Arabic face,
// as the resolver registers them. 2000px: line boxes round to whole pixels, so a
// smaller size hides a 0.5% error in rounding.
//
//   bun tools/font-calibration/vertical-check.ts
import { ARABIC_SCRIPT_FALLBACKS, ARABIC_UNICODE_RANGE, FONT_CALIBRATION, FONT_SUBSTITUTIONS, parseFontCss } from "../../packages/aared-fonts/src/index.ts";
import type { FontCalibration } from "../../packages/aared-fonts/src/index.ts";

const SIZE = 2000;
const FILES: Record<string, { 400: string; 700?: string }> = {
  "Simplified Arabic": { 400: "simpo.ttf", 700: "simpbdo.ttf" },
  "Traditional Arabic": { 400: "trado.ttf", 700: "tradbdo.ttf" },
  "Sakkal Majalla": { 400: "majalla.ttf", 700: "majallab.ttf" },
  "Arabic Typesetting": { 400: "arabtype.ttf" },
  Arial: { 400: "arial.ttf", 700: "arialbd.ttf" },
  "Times New Roman": { 400: "times.ttf", 700: "timesbd.ttf" },
  Cambria: { 400: "times.ttf", 700: "timesbd.ttf" },
  Calibri: { 400: "calibri.ttf", 700: "calibrib.ttf" },
  "Calibri Light": { 400: "calibril.ttf" },
};
const BEFORE: Record<string, FontCalibration> = {
  "Simplified Arabic": { sizeAdjustPercent: 130.3, ascentOverridePercent: 90.2, descentOverridePercent: 36.5 },
  "Traditional Arabic": { sizeAdjustPercent: 80.0, ascentOverridePercent: 124.9, descentOverridePercent: 62.5, bold: { sizeAdjustPercent: 84.5, ascentOverridePercent: 121.4, descentOverridePercent: 59.2 } },
  "Sakkal Majalla": { sizeAdjustPercent: 91.0, ascentOverridePercent: 96.1, descentOverridePercent: 57.7 },
  "Arabic Typesetting": { sizeAdjustPercent: 67.1, ascentOverridePercent: 104.3, descentOverridePercent: 67.1 },
  Arial: { sizeAdjustPercent: 131.8, ascentOverridePercent: 68.3, descentOverridePercent: 15.2 },
  "Times New Roman": { sizeAdjustPercent: 130.2, ascentOverridePercent: 69.2, descentOverridePercent: 17.3 },
  Cambria: { sizeAdjustPercent: 130.2, ascentOverridePercent: 69.2, descentOverridePercent: 17.3 },
  Calibri: { sizeAdjustPercent: 107.2, ascentOverridePercent: 88.7, descentOverridePercent: 25.7, bold: { sizeAdjustPercent: 105.8, ascentOverridePercent: 89.8, descentOverridePercent: 26.0 } },
  "Calibri Light": { sizeAdjustPercent: 105.0, ascentOverridePercent: 90.5, descentOverridePercent: 26.2 },
};

async function faces(family: string) {
  const url = `https://fonts.googleapis.com/css2?family=${encodeURIComponent(family).replace(/%20/g, "+")}:wght@400;700&display=swap`;
  const response = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36" } });
  if (!response.ok) throw new Error(`${family}: CSS ${response.status}`);
  return parseFontCss(await response.text()).filter((face) => face.style === "normal");
}
const entries = Object.keys(FILES).map((family) => {
  const fallback = ARABIC_SCRIPT_FALLBACKS[family];
  return fallback === undefined
    ? { family, files: FILES[family]!, main: FONT_SUBSTITUTIONS[family]!, now: FONT_CALIBRATION[family]!, before: BEFORE[family]! }
    : { family, files: FILES[family]!, latin: FONT_SUBSTITUTIONS[family]!, arabic: fallback.substitute, now: fallback.calibration, before: BEFORE[family]! };
});
const fontFaces: Record<string, Awaited<ReturnType<typeof faces>>> = {};
for (const e of entries) for (const name of [e.main, e.latin, e.arabic]) if (name !== undefined) fontFaces[name] ??= await faces(name);

const server = Bun.serve({
  port: 0,
  fetch(request) {
    const path = new URL(request.url).pathname;
    if (path.startsWith("/ref/")) return new Response(Bun.file(`C:/Windows/Fonts/${path.slice(5)}`), { headers: { "content-type": "font/ttf" } });
    return new Response(`<!doctype html><meta charset="utf-8"><body style="margin:0"><script>window.CFG = ${JSON.stringify({ SIZE, entries, fontFaces, arabicRange: ARABIC_UNICODE_RANGE })};</script></body>`, { headers: { "content-type": "text/html; charset=utf-8" } });
  },
});
try {
  await using view = new Bun.WebView({ width: 1600, height: 900 });
  await view.navigate(server.url.href);
  const rows = JSON.parse((await view.evaluate(`(async () => {
    const { SIZE, entries, fontFaces, arabicRange } = window.CFG;
    const TEXTS = { arabic: "\\u0627\\u0644\\u0633\\u0644\\u0627\\u0645 \\u0639\\u0644\\u064a\\u0643\\u0645", latin: "Hxgy", mixed: "Hxgy \\u0627\\u0644\\u0633\\u0644\\u0627\\u0645" };
    const isArabicSubset = (f) => /U\\+0600-06FF/.test(f.unicodeRange);
    const pick = (c, w) => (w >= 600 && c.bold ? c.bold : c);
    const desc = (o) => ({ sizeAdjust: o.sizeAdjustPercent + "%", ascentOverride: o.ascentOverridePercent + "%", descentOverride: o.descentOverridePercent + "%" });
    const box = document.createElement("div");
    box.style.cssText = "position:absolute;white-space:nowrap;line-height:normal;font-size:" + SIZE + "px";
    document.body.append(box);
    const height = async (family, weight, text) => {
      await document.fonts.load(weight + " " + SIZE + "px '" + family + "'", text);
      box.style.fontFamily = "'" + family + "'"; box.style.fontWeight = weight; box.textContent = text;
      return box.getBoundingClientRect().height;
    };
    const out = [];
    for (const e of entries) {
      for (const [w, file] of Object.entries(e.files)) {
        const weight = Number(w);
        const ref = "ref:" + e.family + w;
        document.fonts.add(new FontFace(ref, "url(/ref/" + file + ")", { weight: w }));
        for (const label of ["before", "now"]) {
          const alias = label + ":" + e.family + w, o = pick(e[label], weight);
          if (e.main) {
            for (const f of fontFaces[e.main].filter((f) => f.weight === w)) document.fonts.add(new FontFace(alias, "url(" + f.url + ")", { weight: w, unicodeRange: f.unicodeRange, ...desc(o) }));
          } else {
            for (const f of fontFaces[e.latin].filter((f) => f.weight === w)) document.fonts.add(new FontFace(alias, "url(" + f.url + ")", { weight: w, unicodeRange: f.unicodeRange }));
            for (const f of fontFaces[e.arabic].filter((f) => f.weight === w && isArabicSubset(f))) document.fonts.add(new FontFace(alias, "url(" + f.url + ")", { weight: w, unicodeRange: arabicRange, ...desc(o) }));
          }
        }
        const texts = e.main ? ["arabic"] : ["arabic", "mixed"];
        for (const kind of texts) {
          const r = await height(ref, weight, TEXTS[kind]);
          out.push({ family: e.family, weight, kind, ref: r, before: await height("before:" + e.family + w, weight, TEXTS[kind]), now: await height("now:" + e.family + w, weight, TEXTS[kind]) });
        }
      }
    }
    return JSON.stringify(out);
  })()`)) as string) as Array<{ family: string; weight: number; kind: string; ref: number; before: number; now: number }>;
  console.log(`line box at line-height: normal, ${SIZE}px (px; % off the real font)`);
  const off = (x: number, ref: number): string => `${x.toFixed(1)} (${(((x / ref) - 1) * 100).toFixed(1)}%)`;
  for (const r of rows) console.log(`  ${`${r.family} ${r.weight} ${r.kind}`.padEnd(30)} real ${r.ref.toFixed(1)}   before ${off(r.before, r.ref).padEnd(16)} now ${off(r.now, r.ref)}`);
} finally {
  server.stop();
}
