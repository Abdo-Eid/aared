// Real-shaped-text width and vertical-metric comparison: each real Office
// Arabic font against open Naskh candidates, at a fixed size, using actual
// Arabic paragraph text (ground truth over glyph tables — shaping matters).
const SAMPLES = [
  "الحفاظ على انتباه المشاهد خلال افتتاحية مدتها ست دقائق بلا تعليق صوتي. ألا تتجاوز مواد الأرشيف عشرين بالمئة من زمن الحلقة. تسليم نسخة نظيفة بلا نصوص مع كل نسخة مترجمة.",
  "مواعيد ما قبل التثبيت النهائي للصورة قابلة للتعديل، أما موعد التثبيت نفسه فغير قابل للتأجيل. توجَّه استفسارات ما بعد الإنتاج إلى قسم المونتاج أولًا، ثم إلى منسق الإنتاج عند الحاجة.",
  "الموعد المجدول للتشغيل التجاري 31 ديسمبر 2028. يُسمح بالتشغيل الجزئي لكتل شمسية لا تقل عن 250 ميجاواط تيار متردد، وتُشترى الطاقة المبكرة بنسبة 75% من تعرفة الطاقة.",
];
const OFFICE_FONTS = ["Simplified Arabic", "Traditional Arabic", "Sakkal Majalla", "Arabic Typesetting"];
const CANDIDATES = ["Noto Naskh Arabic", "Amiri", "Noto Kufi Arabic"];
const SIZE = 40;
const BOLD = process.argv.includes("--bold");

await using view = new Bun.WebView({ width: 1400, height: 900 });
await view.navigate(`data:text/html,<!doctype html><html><body></body></html>`);
const result = await view.evaluate(`(() => {
  const samples = ${JSON.stringify(SAMPLES)};
  const fonts = ${JSON.stringify([...OFFICE_FONTS, ...CANDIDATES])};
  const size = ${SIZE};
  const bold = ${JSON.stringify(BOLD)};
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  const out = {};
  for (const font of fonts) {
    ctx.font = (bold ? "bold " : "") + size + "px '" + font + "'";
    out[font] = samples.map((s) => {
      const m = ctx.measureText(s);
      return {
        width: m.width,
        ascent: m.fontBoundingBoxAscent,
        descent: m.fontBoundingBoxDescent,
        actualAscent: m.actualBoundingBoxAscent,
        actualDescent: m.actualBoundingBoxDescent,
      };
    });
  }
  return JSON.stringify(out);
})()`);
const data = JSON.parse(result as string) as Record<string, Array<{ width: number; ascent: number; descent: number; actualAscent: number; actualDescent: number }>>;

console.log("Per-font per-sample (width, fontBoundingBox ascent/descent):");
for (const [font, rows] of Object.entries(data)) {
  console.log(font.padEnd(20), rows.map((r) => `w=${r.width.toFixed(1)} a=${r.ascent.toFixed(1)} d=${r.descent.toFixed(1)}`).join(" | "));
}

console.log("\nWidth ratio candidate/office (per sample, then mean):");
for (const office of OFFICE_FONTS) {
  for (const candidate of CANDIDATES) {
    const ratios = data[office]!.map((row, i) => data[candidate]![i]!.width / row.width);
    const mean = ratios.reduce((a, b) => a + b, 0) / ratios.length;
    const spread = Math.max(...ratios) - Math.min(...ratios);
    console.log(`${office} <- ${candidate}: ratios=${ratios.map((r) => r.toFixed(4)).join(",")} mean=${mean.toFixed(4)} spread=${spread.toFixed(4)}`);
  }
}

console.log("\nVertical metrics (fontBoundingBox ascent+descent at size " + SIZE + "):");
for (const font of [...OFFICE_FONTS, ...CANDIDATES]) {
  const r = data[font]![0]!;
  console.log(font.padEnd(20), "ascent", r.ascent.toFixed(2), "descent", r.descent.toFixed(2), "total", (r.ascent + r.descent).toFixed(2));
}
