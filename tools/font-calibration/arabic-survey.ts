// Arabic substitute survey. Measures real shaped Arabic text (every Arabic-only
// run in the real fixtures) in each REFERENCE font (the exact Windows font files,
// served locally) against each CANDIDATE open font (the exact Google Fonts files
// the resolver delivers, loaded through FontFace under a private name so a local
// install can never stand in), and reports width ratio, spread and vertical
// metrics. Every load is verified: a candidate that silently fell back would
// measure a generic font and look plausible.
//
//   bun tools/font-calibration/arabic-survey.ts [--bold] [--refs a,b] [--cands x,y]
import { OpcPackage } from "../../../packages/aared/src/core/opc.ts";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.ts";
import { indexPptx, parsePptxSlide } from "../../../packages/aared/src/core/pptx.ts";

const BOLD = process.argv.includes("--bold");
const arg = (name: string): string[] | undefined => {
  const at = process.argv.indexOf(name);
  return at === -1 ? undefined : process.argv[at + 1]!.split(",");
};
const WEIGHT = BOLD ? 700 : 400;
const SIZE = 40;

const WINDOWS = "C:/Windows/Fonts";
const REFERENCES: Record<string, { regular: string; bold: string }> = {
  Arial: { regular: "arial.ttf", bold: "arialbd.ttf" },
  "Times New Roman": { regular: "times.ttf", bold: "timesbd.ttf" },
  Tahoma: { regular: "tahoma.ttf", bold: "tahomabd.ttf" },
  Calibri: { regular: "calibri.ttf", bold: "calibrib.ttf" },
  "Calibri Light": { regular: "calibril.ttf", bold: "calibril.ttf" },
  "Courier New": { regular: "cour.ttf", bold: "courbd.ttf" },
  "Simplified Arabic": { regular: "simpo.ttf", bold: "simpbdo.ttf" },
  "Traditional Arabic": { regular: "trado.ttf", bold: "tradbdo.ttf" },
  "Sakkal Majalla": { regular: "majalla.ttf", bold: "majallab.ttf" },
  "Arabic Typesetting": { regular: "arabtype.ttf", bold: "arabtype.ttf" },
};
const CANDIDATES = arg("--cands") ?? [
  "Amiri",
  "Noto Naskh Arabic",
  "Noto Sans Arabic",
  "Noto Kufi Arabic",
  "Scheherazade New",
  "Lateef",
  "Markazi Text",
  "IBM Plex Sans Arabic",
  "Vazirmatn",
  "Cairo",
  "Tajawal",
];
const REFS = arg("--refs") ?? Object.keys(REFERENCES);

// --- samples: every Arabic-only run (no Latin letters) of real text -----------
const ARABIC = /[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]/;
const samples = new Set<string>();
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
    if (typeof record.text === "string" && ARABIC.test(record.text) && !/[A-Za-z]/.test(record.text)) {
      const text = record.text.trim();
      if (text.length >= 12) samples.add(text);
    }
    for (const name in record) visit(record[name]);
  };
  visit(model);
};
const docx = await OpcPackage.open(await Bun.file("fixtures/docx-coverage.docx").arrayBuffer());
collect(await parseDocxSections(docx, await indexDocx(docx)));
for (const name of ["pptx-coverage.pptx", "post-production-review.pptx"]) {
  const pkg = await OpcPackage.open(await Bun.file(`fixtures/${name}`).arrayBuffer());
  const deck = await indexPptx(pkg);
  for (let i = 0; i < deck.slideParts.length; i++) collect(await parsePptxSlide(pkg, deck, i));
}
const SAMPLES = [...samples];
console.log(`${SAMPLES.length} Arabic-only samples, ${SAMPLES.reduce((n, s) => n + s.length, 0)} chars\n`);

// --- candidate CSS (the same faces the resolver registers) ---------------------
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36";
interface CssFace { subset: string; url: string; unicodeRange: string; weight: string }
async function candidateFaces(family: string): Promise<CssFace[] | undefined> {
  const url = `https://fonts.googleapis.com/css2?family=${encodeURIComponent(family).replace(/%20/g, "+")}:wght@${WEIGHT}&display=swap`;
  const response = await fetch(url, { headers: { "User-Agent": UA } });
  if (!response.ok) return undefined;
  const css = await response.text();
  const faces: CssFace[] = [];
  for (const block of css.matchAll(/\/\*\s*([\w-]+)\s*\*\/\s*@font-face\s*{([^}]*)}/g)) {
    const body = block[2]!;
    const src = /url\((https:[^)]+)\)/.exec(body)?.[1];
    if (src === undefined) continue;
    faces.push({
      subset: block[1]!,
      url: src,
      unicodeRange: /unicode-range\s*:\s*([^;]+);/.exec(body)?.[1]?.trim() ?? "",
      weight: /font-weight\s*:\s*([^;]+);/.exec(body)?.[1]?.trim() ?? String(WEIGHT),
    });
  }
  return faces;
}
const candidates: Record<string, CssFace[]> = {};
for (const family of CANDIDATES) {
  const faces = await candidateFaces(family);
  if (faces === undefined || faces.length === 0) console.log(`(skip ${family}: no CSS at weight ${WEIGHT})`);
  else candidates[family] = faces;
}

// --- serve the page + the exact Windows font files ------------------------------
const config = {
  size: SIZE,
  weight: WEIGHT,
  samples: SAMPLES,
  refs: REFS.map((name) => ({ name, file: BOLD ? REFERENCES[name]!.bold : REFERENCES[name]!.regular })),
  candidates,
};
const server = Bun.serve({
  port: 0,
  async fetch(request) {
    const path = new URL(request.url).pathname;
    if (path.startsWith("/ref/")) {
      const file = Bun.file(`${WINDOWS}/${path.slice(5)}`);
      return (await file.exists()) ? new Response(file, { headers: { "content-type": "font/ttf" } }) : new Response("nope", { status: 404 });
    }
    return new Response(
      `<!doctype html><meta charset="utf-8"><body><script>window.CFG = ${JSON.stringify(config)};</script></body>`,
      { headers: { "content-type": "text/html; charset=utf-8" } },
    );
  },
});

interface Measure { widths: number[]; ascent: number; descent: number; alefHeight: number; seenHeight: number; status?: string }
try {
  await using view = new Bun.WebView({ width: 1200, height: 800 });
  await view.navigate(server.url.href);
  const raw = (await view.evaluate(`(async () => {
    const { size, weight, samples, refs, candidates } = window.CFG;
    const all = samples.join(" ");
    const added = [];
    for (const ref of refs) {
      const face = new FontFace("ref:" + ref.name, "url(/ref/" + ref.file + ")", { weight: String(weight) });
      document.fonts.add(face); added.push({ key: ref.name, face });
    }
    for (const [name, faces] of Object.entries(candidates)) {
      for (const f of faces) {
        const face = new FontFace("cand:" + name, "url(" + f.url + ")", { weight: f.weight, unicodeRange: f.unicodeRange });
        document.fonts.add(face); added.push({ key: name, face, subset: f.subset });
      }
    }
    const ctx = document.createElement("canvas").getContext("2d");
    const measure = (family) => {
      // Vertical metrics come back in whole pixels: read them at 10000px, or a 40px
      // reading is off by up to half a pixel (2.5% of the em in the shipped overrides).
      ctx.font = weight + " 10000px " + family;
      const big = ctx.measureText("x");
      ctx.font = weight + " " + size + "px " + family;
      const rows = samples.map((s) => ctx.measureText(s));
      return {
        widths: rows.map((m) => m.width),
        ascent: big.fontBoundingBoxAscent / 10000 * size,
        descent: big.fontBoundingBoxDescent / 10000 * size,
        // Letter height, from plain letters with no marks or tails: alef (the tallest stroke) and a
        // row of seen (the body height). Whole-sample ink extents are dominated by diacritics.
        alefHeight: ctx.measureText("\\u0627\\u0627\\u0627").actualBoundingBoxAscent,
        seenHeight: ctx.measureText("\\u0633\\u0633\\u0633\\u0633").actualBoundingBoxAscent,
      };
    };
    const out = { controls: {}, ref: {}, cand: {}, loadErrors: [] };
    out.controls.sans = measure("sans-serif");
    out.controls.serif = measure("serif");
    for (const { key, face } of added) {
      try { await face.load(); } catch (e) { out.loadErrors.push(key + ": " + e); }
    }
    // unicode-range faces only load on demand: force the ones these samples need
    for (const [name] of Object.entries(candidates)) await document.fonts.load(weight + " " + size + "px 'cand:" + name + "'", all);
    for (const ref of refs) {
      out.ref[ref.name] = measure("'ref:" + ref.name + "'");
      out.ref[ref.name].status = added.find((a) => a.key === ref.name).face.status;
    }
    for (const [name] of Object.entries(candidates)) {
      out.cand[name] = measure("'cand:" + name + "'");
      out.cand[name].loaded = added.filter((a) => a.key === name && a.face.status === "loaded").map((a) => a.subset);
    }
    return JSON.stringify(out);
  })()`)) as string;
  const data = JSON.parse(raw) as {
    controls: Record<string, Measure>;
    ref: Record<string, Measure>;
    cand: Record<string, Measure & { loaded: string[] }>;
    loadErrors: string[];
  };
  if (data.loadErrors.length > 0) console.log("LOAD ERRORS:", data.loadErrors);

  const sum = (values: number[]): number => values.reduce((a, b) => a + b, 0);
  // Guard: a face that did not reach "loaded" measured a fallback. (Comparing against the generic
  // fallback instead would false-positive here: on Windows generic sans-serif/serif ARE Arial/TNR.)
  for (const [name, m] of Object.entries(data.ref)) {
    if (m.status !== "loaded") console.log(`!! reference ${name} did not load (status ${m.status})`);
  }
  for (const [name, m] of Object.entries(data.cand)) {
    const same = m.widths.every((w, i) => Math.abs(w - data.controls.sans!.widths[i]!) < 0.01 || Math.abs(w - data.controls.serif!.widths[i]!) < 0.01);
    const arabicLoaded = m.loaded.includes("arabic");
    if (same || !arabicLoaded) console.log(`!! candidate ${name} suspect: fallbackLike=${same} arabicSubsetLoaded=${arabicLoaded} loaded=[${m.loaded}]`);
  }

  console.log(`Vertical metrics at ${SIZE}px (fontBoundingBox ascent/descent, em):`);
  for (const [name, m] of Object.entries(data.ref)) console.log("  ref ", name.padEnd(20), (m.ascent / SIZE).toFixed(3), (m.descent / SIZE).toFixed(3));
  for (const [name, m] of Object.entries(data.cand)) console.log("  cand", name.padEnd(20), (m.ascent / SIZE).toFixed(3), (m.descent / SIZE).toFixed(3));

  console.log(`\nwidth ratio candidate/reference (weight ${WEIGHT}); mean = sum/sum, sd = spread of per-sample ratios, worst = max |ratio/mean - 1|`);
  console.log("lower sd = calibration keeps line breaks; the mean is what size-adjust removes.\n");
  for (const ref of REFS) {
    const r = data.ref[ref];
    if (r === undefined) continue;
    console.log(`== ${ref}`);
    const rows = Object.entries(data.cand).map(([name, c]) => {
      const ratios = c.widths.map((w, i) => w / r.widths[i]!);
      const mean = sum(c.widths) / sum(r.widths);
      const sd = Math.sqrt(sum(ratios.map((x) => (x - mean) ** 2)) / ratios.length);
      const worst = Math.max(...ratios.map((x) => Math.abs(x / mean - 1)));
      // After size-adjust (1/mean) matches widths, how tall do alef and the seen body end up vs the
      // reference? 1.00 = same proportions; <1 = letters end up shorter than the original's.
      const alef = c.alefHeight / r.alefHeight / mean;
      const seen = c.seenHeight / r.seenHeight / mean;
      return { name, mean, sd, worst, alef, seen };
    }).sort((a, b) => a.sd - b.sd);
    for (const row of rows) {
      console.log(`   ${row.name.padEnd(22)} mean=${row.mean.toFixed(4)} sd=${row.sd.toFixed(4)} worst=${(row.worst * 100).toFixed(1)}% sizeAdjust=${(100 / row.mean).toFixed(1)}% alef=${row.alef.toFixed(2)} seen=${row.seen.toFixed(2)}`);
    }
  }

  // --calib "Ref=Candidate,Ref=Candidate": the exact FontFace descriptors for chosen pairs.
  // A browser multiplies an override percentage by size-adjust's factor (size-adjust-interaction.ts),
  // so override = (reference em metric) / sizeAdjust.
  // Only for a main-face substitute (FONT_CALIBRATION), whose own spaces paint. An Arabic
  // fallback face (ARABIC_SCRIPT_FALLBACKS) paints next to a Latin face that owns the spaces:
  // fit it through its alias with arabic-fallback-fit.ts, never with these numbers.
  const calib = arg("--calib");
  if (calib !== undefined) {
    console.log("\ncalibration (regular metrics; size-adjust / ascent-override / descent-override):");
    for (const pair of calib) {
      const [refName, candName] = pair.split("=") as [string, string];
      const r = data.ref[refName], c = data.cand[candName];
      if (r === undefined || c === undefined) { console.log(`   ${pair}: missing`); continue; }
      const mean = sum(c.widths) / sum(r.widths);
      const s = 1 / mean;
      console.log(`   ${refName} <- ${candName}: ${(s * 100).toFixed(1)}% / ${((r.ascent / SIZE / s) * 100).toFixed(1)}% / ${((r.descent / SIZE / s) * 100).toFixed(1)}%`);
    }
  }

  console.log("\nreference vs reference (does Arial's Arabic match Simplified Arabic?):");
  for (const a of ["Arial", "Times New Roman", "Tahoma"]) {
    for (const b of ["Simplified Arabic", "Traditional Arabic", "Sakkal Majalla", "Arabic Typesetting", "Arial", "Times New Roman"]) {
      const ra = data.ref[a], rb = data.ref[b];
      if (ra === undefined || rb === undefined || a === b) continue;
      const ratios = ra.widths.map((w, i) => w / rb.widths[i]!);
      const mean = sum(ra.widths) / sum(rb.widths);
      const sd = Math.sqrt(sum(ratios.map((x) => (x - mean) ** 2)) / ratios.length);
      console.log(`   ${a} / ${b}: mean=${mean.toFixed(4)} sd=${sd.toFixed(4)}`);
    }
  }
  await Bun.write(Bun.env.SURVEY_OUT ?? "tools/font-calibration/arabic-survey.last.json", JSON.stringify({ weight: WEIGHT, size: SIZE, samples: SAMPLES, data }, null, 1));
} finally {
  server.stop();
}
