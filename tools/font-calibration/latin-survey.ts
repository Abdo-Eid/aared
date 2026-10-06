// Aptos substitute survey. Measures every Latin run in the five fixtures in the
// REAL Aptos, Aptos Display and Aptos Narrow (Office's cloud-font cache: Office
// downloads them on demand and never installs them as system fonts) against
// open CANDIDATES delivered exactly as the resolver delivers them
// (`fontCssUrl` + `parseFontCss`, every subset, private family names so a local
// install can never stand in). Every load is verified: a reference or candidate
// that silently fell back would measure the generic font and look plausible.
// Reports per pair: mean width ratio (what size-adjust removes), sd of
// per-sample ratios (what line breaks follow), worst deviation, and x-height /
// cap-height after adjusting. Digit strings (spreadsheet numbers) are reported
// on their own.
//
//   bun tools/font-calibration/latin-survey.ts [--bold] [--italic] [--refs a,b] [--cands x,y] [--calib "Ref=Cand,..."]
import { readdirSync } from "node:fs";
import { OpcPackage } from "../../../packages/aared/src/core/opc.ts";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.ts";
import { indexPptx, parsePptxSlide } from "../../../packages/aared/src/core/pptx.ts";
import { readWorkbook } from "../../../packages/aared/src/core/xlsx/workbook.ts";
import { parseXlsxSheet } from "../../../packages/aared/src/core/xlsx/sheet.ts";
import { fontCssUrl, parseFontCss } from "../../../packages/aared-fonts/src/index.ts";
import { cloudFontFile } from "./cloud-fonts.ts";

const BOLD = process.argv.includes("--bold");
const ITALIC = process.argv.includes("--italic");
const arg = (name: string): string[] | undefined => {
  const at = process.argv.indexOf(name);
  return at === -1 ? undefined : process.argv[at + 1]!.split(",");
};
const WEIGHT = BOLD ? 700 : 400;
const STYLE = ITALIC ? "italic" : "normal";
const SIZE = 40;

const styleName = (family: string): string => {
  const suffix = [BOLD ? "Bold" : "", ITALIC ? "Italic" : ""].filter(Boolean).join(" ");
  return suffix === "" ? family : `${family} ${suffix}`;
};
const REFS = arg("--refs") ?? ["Aptos", "Aptos Display", "Aptos Narrow"];
const refFiles: Record<string, string> = {};
for (const family of REFS) refFiles[family] = await cloudFontFile(family, styleName(family));

const CANDIDATES = arg("--cands") ?? [
  "Arimo", "Open Sans", "Noto Sans", "Roboto", "Inter", "Source Sans 3", "IBM Plex Sans", "Work Sans", "Public Sans", "Figtree",
  "Instrument Sans", "Albert Sans", "Hanken Grotesk", "Schibsted Grotesk", "Plus Jakarta Sans", "DM Sans", "Lato", "Carlito",
  "Mona Sans", "Hubot Sans", "Red Hat Text", "Red Hat Display", "Libre Franklin", "Archivo", "Rubik", "Karla", "Barlow",
  "Nunito Sans", "PT Sans", "Ubuntu", "Overpass", "Be Vietnam Pro", "Atkinson Hyperlegible Next", "Wix Madefor Text",
  "Radio Canada", "Reddit Sans", "Gantari", "Epilogue", "Funnel Sans", "Host Grotesk", "Noto Sans Display", "Fira Sans",
  "Mulish", "Rethink Sans", "Afacad", "Asap", "Cabin", "Chivo", "Sofia Sans", "SUSE", "Ysabeau Office", "Inclusive Sans",
  "Roboto Condensed", "Archivo Narrow", "IBM Plex Sans Condensed", "Fira Sans Condensed", "Barlow Semi Condensed",
  "Sofia Sans Semi Condensed", "Sofia Sans Condensed", "Asap Condensed", "Pathway Extreme",
];

// --- samples: every Latin run (no Arabic letters) of real text ----------------
const ARABIC = /[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]/;
const LATIN = /[A-Za-z]/;
const DIGITS = /^[\d\s.,%$\u20ac\u00a3:/()+\-\u2212]+$/;
const texts = new Set<string>();
const numbers = new Set<string>();
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
      if (typeof raw !== "string" || ARABIC.test(raw)) continue;
      const text = raw.trim();
      if (LATIN.test(text) && text.length >= 8) texts.add(text);
      else if (DIGITS.test(text) && /\d/.test(text) && text.length >= 3) numbers.add(text);
    }
    for (const name in record) visit(record[name]);
  };
  visit(model);
};
const FIXTURES = "fixtures";
for (const name of readdirSync(FIXTURES)) {
  const pkg = async () => OpcPackage.open(await Bun.file(`${FIXTURES}/${name}`).arrayBuffer());
  if (name.endsWith(".docx")) {
    const doc = await pkg();
    collect(await parseDocxSections(doc, await indexDocx(doc)));
  } else if (name.endsWith(".pptx")) {
    const deck = await pkg();
    const index = await indexPptx(deck);
    for (let i = 0; i < index.slideParts.length; i++) collect(await parsePptxSlide(deck, index, i));
  } else if (name.endsWith(".xlsx")) {
    const book = await pkg();
    const workbook = await readWorkbook(book);
    for (let i = 0; i < workbook.sheets.length; i++) collect(await parseXlsxSheet(book, workbook, i));
  }
}
const TEXTS = [...texts];
const NUMBERS = [...numbers];
console.log(`${TEXTS.length} Latin samples (${TEXTS.reduce((n, s) => n + s.length, 0)} chars), ${NUMBERS.length} digit strings; weight ${WEIGHT} ${STYLE}\n`);

// --- candidate CSS: the resolver's exact request ------------------------------
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36";
const candidates: Record<string, Array<{ url: string; unicodeRange: string; style: string; weight: string }>> = {};
for (const family of CANDIDATES) {
  try {
    const response = await fetch(fontCssUrl([family]), { headers: { "User-Agent": UA } });
    const faces = response.ok ? parseFontCss(await response.text()).filter((face) => face.family === family && face.weight === String(WEIGHT)) : [];
    // A family with no italic gets the upright face, slanted by the browser, as it would in production.
    const styled = faces.filter((face) => face.style === STYLE);
    const picked = styled.length > 0 ? styled : faces.filter((face) => face.style === "normal");
    if (picked.length === 0) console.log(`(skip ${family}: nothing at ${WEIGHT} ${STYLE})`);
    else candidates[family] = picked;
    if (styled.length === 0 && picked.length > 0) console.log(`(${family}: no ${STYLE} face, synthesized)`);
  } catch (error) {
    console.log(`(skip ${family}: ${error})`);
  }
}

const config = { size: SIZE, weight: WEIGHT, style: STYLE, texts: TEXTS, numbers: NUMBERS, refs: REFS, candidates };
const server = Bun.serve({
  port: 0,
  async fetch(request) {
    const path = decodeURIComponent(new URL(request.url).pathname);
    if (path.startsWith("/ref/")) return new Response(Bun.file(refFiles[path.slice(5)]!), { headers: { "content-type": "font/ttf" } });
    return new Response(`<!doctype html><meta charset="utf-8"><body><script>window.CFG = ${JSON.stringify(config)};</script></body>`, { headers: { "content-type": "text/html; charset=utf-8" } });
  },
});

interface Measure { texts: number[]; numbers: number[]; ascent: number; descent: number; xHeight: number; capHeight: number; status?: string; loaded?: number }
try {
  await using view = new Bun.WebView({ width: 1200, height: 800 });
  await view.navigate(server.url.href);
  const raw = (await view.evaluate(`(async () => {
    const { size, weight, style, texts, numbers, refs, candidates } = window.CFG;
    const all = texts.join(" ") + " " + numbers.join(" ");
    const out = { controls: {}, ref: {}, cand: {}, loadErrors: [] };
    const refFaces = {}, candFaces = {};
    for (const name of refs) {
      const face = new FontFace("ref:" + name, "url(/ref/" + encodeURIComponent(name) + ")", { weight: String(weight), style });
      document.fonts.add(face); refFaces[name] = face;
    }
    for (const [name, faces] of Object.entries(candidates)) {
      candFaces[name] = faces.map((f) => {
        const face = new FontFace("cand:" + name, "url(" + f.url + ")", { weight: String(weight), style, unicodeRange: f.unicodeRange });
        document.fonts.add(face); return face;
      });
    }
    const ctx = document.createElement("canvas").getContext("2d");
    const font = (px, family) => style + " " + weight + " " + px + "px " + family;
    const measure = (family) => {
      // Vertical metrics come back in whole pixels: read them at 10000px.
      ctx.font = font(10000, family);
      const big = ctx.measureText("x");
      const x = ctx.measureText("x").actualBoundingBoxAscent / 10000 * size;
      const cap = ctx.measureText("H").actualBoundingBoxAscent / 10000 * size;
      ctx.font = font(size, family);
      return {
        texts: texts.map((s) => ctx.measureText(s).width),
        numbers: numbers.map((s) => ctx.measureText(s).width),
        ascent: big.fontBoundingBoxAscent / 10000 * size,
        descent: big.fontBoundingBoxDescent / 10000 * size,
        xHeight: x,
        capHeight: cap,
      };
    };
    out.controls.sans = measure("sans-serif");
    for (const [name, face] of Object.entries(refFaces)) { try { await face.load(); } catch (e) { out.loadErrors.push(name + ": " + e); } }
    for (const name of Object.keys(candFaces)) await document.fonts.load(font(size, "'cand:" + name + "'"), all).catch((e) => out.loadErrors.push(name + ": " + e));
    for (const name of refs) out.ref[name] = { ...measure("'ref:" + name + "'"), status: refFaces[name].status };
    for (const [name, faces] of Object.entries(candFaces)) out.cand[name] = { ...measure("'cand:" + name + "'"), loaded: faces.filter((f) => f.status === "loaded").length };
    return JSON.stringify(out);
  })()`)) as string;
  const data = JSON.parse(raw) as { controls: Record<string, Measure>; ref: Record<string, Measure>; cand: Record<string, Measure>; loadErrors: string[] };
  if (data.loadErrors.length > 0) console.log("LOAD ERRORS:", data.loadErrors);
  for (const [name, m] of Object.entries(data.ref)) if (m.status !== "loaded") console.log(`!! reference ${name} did not load (status ${m.status})`);
  for (const [name, m] of Object.entries(data.cand)) {
    const same = m.texts.every((w, i) => Math.abs(w - data.controls.sans!.texts[i]!) < 0.01);
    if (same || m.loaded === 0) console.log(`!! candidate ${name} suspect: fallbackLike=${same} facesLoaded=${m.loaded}`);
  }

  const sum = (values: number[]): number => values.reduce((a, b) => a + b, 0);
  const stats = (cand: number[], ref: number[]) => {
    const ratios = cand.map((w, i) => w / ref[i]!);
    const mean = sum(cand) / sum(ref);
    const sd = Math.sqrt(sum(ratios.map((x) => (x - mean) ** 2)) / ratios.length);
    const worst = Math.max(...ratios.map((x) => Math.abs(x / mean - 1)));
    return { mean, sd, worst };
  };
  console.log(`vertical metrics (em): ascent / descent / x-height / cap-height`);
  for (const [name, m] of Object.entries(data.ref)) console.log(`  ref  ${name.padEnd(26)} ${(m.ascent / SIZE).toFixed(4)} ${(m.descent / SIZE).toFixed(4)} ${(m.xHeight / SIZE).toFixed(3)} ${(m.capHeight / SIZE).toFixed(3)}`);
  console.log(`\nwidth ratio candidate/reference; sd = spread of per-sample ratios (line breaks follow it), mean = what size-adjust removes;`);
  console.log(`x/cap = x-height and cap-height after size-adjust against the reference's (1.00 = same proportions); digits = the number strings' own mean and sd.\n`);
  const results: Record<string, Array<{ name: string; mean: number; sd: number; worst: number; x: number; cap: number; dmean: number; dsd: number }>> = {};
  for (const ref of REFS) {
    const r = data.ref[ref];
    if (r === undefined) continue;
    const rows = Object.entries(data.cand).map(([name, c]) => {
      const t = stats(c.texts, r.texts);
      const d = stats(c.numbers, r.numbers);
      return { name, ...t, x: c.xHeight / r.xHeight / t.mean, cap: c.capHeight / r.capHeight / t.mean, dmean: d.mean / t.mean, dsd: d.sd };
    }).sort((a, b) => a.sd - b.sd);
    results[ref] = rows;
    console.log(`== ${ref}`);
    for (const row of rows.slice(0, 18)) {
      console.log(`   ${row.name.padEnd(26)} sd=${row.sd.toFixed(4)} worst=${(row.worst * 100).toFixed(1).padStart(4)}% mean=${row.mean.toFixed(3)} sizeAdjust=${(100 / row.mean).toFixed(1)}% x=${row.x.toFixed(2)} cap=${row.cap.toFixed(2)} digits(after adjust)=${row.dmean.toFixed(3)} dsd=${row.dsd.toFixed(4)}`);
    }
    const arimo = rows.find((row) => row.name === "Arimo");
    if (arimo !== undefined) console.log(`   (Arimo, roughly what paints today: sd=${arimo.sd.toFixed(4)} mean=${arimo.mean.toFixed(3)})`);
  }

  // --calib "Ref=Cand,...": the exact FontFace descriptors. A browser multiplies an
  // override percentage by size-adjust's factor (size-adjust-interaction.ts), so
  // override = reference em metric / size-adjust.
  const calib = arg("--calib");
  if (calib !== undefined) {
    console.log(`\ncalibration (size-adjust / ascent-override / descent-override), weight ${WEIGHT} ${STYLE}:`);
    for (const pair of calib) {
      const [refName, candName] = pair.split("=") as [string, string];
      const r = data.ref[refName], c = data.cand[candName];
      if (r === undefined || c === undefined) { console.log(`   ${pair}: missing`); continue; }
      const s = sum(r.texts) / sum(c.texts);
      console.log(`   ${refName} <- ${candName}: ${(s * 100).toFixed(1)}% / ${((r.ascent / SIZE / s) * 100).toFixed(1)}% / ${((r.descent / SIZE / s) * 100).toFixed(1)}%  (unrounded size-adjust ${(s * 100).toFixed(3)})`);
    }
  }
  await Bun.write(`tools/font-calibration/latin-survey${BOLD ? "-bold" : ""}${ITALIC ? "-italic" : ""}.last.json`, JSON.stringify({ weight: WEIGHT, style: STYLE, size: SIZE, texts: TEXTS, numbers: NUMBERS, refFiles, results, data }, null, 1));
} finally {
  server.stop();
}
