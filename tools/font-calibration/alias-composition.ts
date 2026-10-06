// Does one alias carrying a Latin face (Arimo) AND a unicode-range-limited Arabic face
// (calibrated Amiri) compose the way the resolver needs? Checks, in a real browser:
//   1. Latin text under the alias measures exactly like Arimo alone (Arabic face never touches it)
//   2. Arabic text under the alias measures exactly like calibrated Amiri alone, and is NOT the
//      generic fallback (control: Arimo alone falls through to a deliberately distinct font)
//   3. Punctuation that both Google subsets claim (U+2010, U+200C-200E, U+2019) stays Arimo
//   4. Mixed text works, and add order (Arabic first vs last) makes no difference
//   5. size-adjust/ascent/descent descriptors apply to the Arabic face only
//
//   bun tools/font-calibration/alias-composition.ts
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36";
async function faces(family: string): Promise<Array<{ url: string; unicodeRange: string; subset: string }>> {
  const css = await (await fetch(`https://fonts.googleapis.com/css2?family=${family.replace(/ /g, "+")}:wght@400&display=swap`, { headers: { "User-Agent": UA } })).text();
  return [...css.matchAll(/\/\*\s*([\w-]+)\s*\*\/\s*@font-face\s*{([^}]*)}/g)].map((m) => ({
    subset: m[1]!,
    url: /url\((https:[^)]+)\)/.exec(m[2]!)![1]!,
    unicodeRange: /unicode-range\s*:\s*([^;]+);/.exec(m[2]!)![1]!.trim(),
  }));
}
const arimo = await faces("Arimo");
const amiri = await faces("Amiri");
console.log("Arimo subsets:", arimo.map((f) => f.subset).join(","), "| Amiri subsets:", amiri.map((f) => f.subset).join(","));

const CAL = { sizeAdjust: "99.9%", ascentOverride: "90.1%", descentOverride: "20%" };
const ARABIC_RANGE = "U+0600-06FF, U+0750-077F, U+0870-08FF, U+FB50-FDFF, U+FE70-FEFF";
const samples = (JSON.parse(await Bun.file("tools/font-calibration/arabic-survey.last.json").text()) as { samples: string[] }).samples.slice(0, 6);

const server = Bun.serve({
  port: 0,
  fetch: () => new Response(`<!doctype html><meta charset="utf-8"><script>window.CFG=${JSON.stringify({ arimo, amiri, CAL, ARABIC_RANGE, samples })};</script>`, { headers: { "content-type": "text/html; charset=utf-8" } }),
});
try {
  await using view = new Bun.WebView({ width: 900, height: 600 });
  await view.navigate(server.url.href);
  const out = JSON.parse((await view.evaluate(`(async () => {
    const { arimo, amiri, CAL, ARABIC_RANGE, samples } = window.CFG;
    const add = (family, list, extra = {}, range) => {
      for (const f of list) {
        const face = new FontFace(family, "url(" + f.url + ")", { unicodeRange: range ?? f.unicodeRange, ...extra });
        document.fonts.add(face);
      }
    };
    const arabicAmiri = amiri.filter((f) => /0600-06FF/.test(f.unicodeRange));
    // Alias A: Arabic added LAST.   Alias B: Arabic added FIRST.
    add("t:latinFirst", arimo);   add("t:latinFirst", arabicAmiri, CAL, ARABIC_RANGE);
    add("t:arabicFirst", arabicAmiri, CAL, ARABIC_RANGE);  add("t:arabicFirst", arimo);
    // Controls
    add("c:arimoOnly", arimo);
    add("c:amiriCal", amiri, CAL);               // all Amiri subsets, calibrated, same descriptors
    add("c:amiriPlain", amiri);
    const latin = "The quick brown fox 0123456789 (Word) - test";
    const punct = "\\u2010\\u200C\\u200D\\u200E\\u2019\\u2013 ok";
    const mixed = samples[0] + " Word 2024 " + samples[1];
    await document.fonts.load("40px 't:latinFirst'", latin + samples.join(" "));
    await document.fonts.load("40px 't:arabicFirst'", latin + samples.join(" "));
    await document.fonts.load("40px 'c:arimoOnly'", latin);
    await document.fonts.load("40px 'c:amiriCal'", latin + samples.join(" "));
    await document.fonts.load("40px 'c:amiriPlain'", latin + samples.join(" "));
    const ctx = document.createElement("canvas").getContext("2d");
    const w = (stack, text) => { ctx.font = "40px " + stack; const m = ctx.measureText(text); return { width: m.width, ascent: m.fontBoundingBoxAscent, descent: m.fontBoundingBoxDescent }; };
    const res = {};
    for (const [name, text] of [["latin", latin], ["punct", punct], ["arabic", samples.join(" ").replace(/[^؀-ۿ]/g, "")], ["arabicWithSpaces", samples.join(" ")], ["mixed", mixed]]) {
      res[name] = {
        alias: w("'t:latinFirst', 'Tahoma'", text),
        aliasArabicFirst: w("'t:arabicFirst', 'Tahoma'", text),
        arimoOnly: w("'c:arimoOnly', 'Tahoma'", text),
        amiriCal: w("'c:amiriCal', 'Tahoma'", text),
        amiriPlain: w("'c:amiriPlain', 'Tahoma'", text),
      };
    }
    res.status = [...document.fonts].map((f) => f.family + ":" + f.status);
    return JSON.stringify(res);
  })()`)) as string);

  const eq = (a: number, b: number): string => (Math.abs(a - b) < 0.01 ? "EQUAL" : `DIFFER (${a.toFixed(2)} vs ${b.toFixed(2)})`);
  console.log("\n1. Latin under alias vs Arimo alone:        ", eq(out.latin.alias.width, out.latin.arimoOnly.width));
  console.log("   punctuation both subsets claim vs Arimo: ", eq(out.punct.alias.width, out.punct.arimoOnly.width));
  console.log("2. Arabic under alias vs calibrated Amiri:  ", eq(out.arabic.alias.width, out.arabic.amiriCal.width));
  console.log("   Arabic under alias vs Arimo-only (falls to Tahoma): ", eq(out.arabic.alias.width, out.arabic.arimoOnly.width), "<- must DIFFER");
  console.log("   calibration applied? alias / uncalibrated Amiri width =", (out.arabic.alias.width / out.arabic.amiriPlain.width).toFixed(4), "(expect 0.999)");
  console.log("   Arabic + spaces/digits (spaces come from Arimo under the alias, by design):", eq(out.arabicWithSpaces.alias.width, out.arabicWithSpaces.amiriCal.width), "-> diff", (out.arabicWithSpaces.alias.width - out.arabicWithSpaces.amiriCal.width).toFixed(2), "px");
  console.log("   calibrated Amiri vertical metrics", out.arabic.amiriCal.ascent, out.arabic.amiriCal.descent, "vs uncalibrated", out.arabic.amiriPlain.ascent, out.arabic.amiriPlain.descent, "(override took effect if these differ)");
  console.log("3. mixed Arabic+Latin, alias vs calibrated Amiri:", eq(out.mixed.alias.width, out.mixed.amiriCal.width), "(Latin part differs by design: Arimo vs Amiri Latin)");
  console.log("4. add order: latin-first vs arabic-first     :", ["latin", "punct", "arabic", "arabicWithSpaces", "mixed"].map((k) => eq(out[k].alias.width, out[k].aliasArabicFirst.width)).join(" | "));
  console.log("5. vertical metrics at 40px (ascent/descent): alias latin", out.latin.alias.ascent, out.latin.alias.descent, "| alias arabic", out.arabic.alias.ascent, out.arabic.alias.descent, "| Arimo", out.latin.arimoOnly.ascent, out.latin.arimoOnly.descent, "| calibrated Amiri", out.arabic.amiriCal.ascent, out.arabic.amiriCal.descent);
  console.log("   faces:", (out.status as string[]).filter((s) => s.startsWith("t:")).join(" "));
} finally {
  server.stop();
}
