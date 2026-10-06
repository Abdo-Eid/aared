// Empirical test: does `ascent-override`/`descent-override` apply before or
// after `size-adjust` scales the font? Load the same local Amiri file under
// several FontFace descriptor combinations and measure the resulting
// fontBoundingBoxAscent/Descent via canvas at a fixed font-size.
await using view = new Bun.WebView({ width: 900, height: 700 });
await view.navigate("data:text/html,<!doctype html><html><body></body></html>");
const amiriPath = "C:/Windows/Fonts/Amiri-Regular.ttf";
const amiriBytes = Buffer.from(await Bun.file(amiriPath).arrayBuffer());
const amiriBase64 = amiriBytes.toString("base64");

const result = await view.evaluate(`(async () => {
  const base64 = ${JSON.stringify(amiriBase64)};
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const size = 100;
  const cases = [
    { name: "baseline", desc: {} },
    { name: "size-adjust 50%", desc: { sizeAdjust: "50%" } },
    { name: "ascent-override 60% descent-override 20%", desc: { ascentOverride: "60%", descentOverride: "20%" } },
    { name: "size-adjust 50% + ascent-override 60% descent-override 20%", desc: { sizeAdjust: "50%", ascentOverride: "60%", descentOverride: "20%" } },
  ];
  const out = {};
  let i = 0;
  for (const c of cases) {
    const family = "Probe" + (i++);
    const face = new FontFace(family, bytes.buffer.slice(0), c.desc);
    await face.load();
    document.fonts.add(face);
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    ctx.font = size + "px '" + family + "'";
    const m = ctx.measureText("محمد عربي نص");
    out[c.name] = { width: m.width, ascent: m.fontBoundingBoxAscent, descent: m.fontBoundingBoxDescent };
  }
  return JSON.stringify(out);
})()`);
console.log(JSON.stringify(JSON.parse(result as string), null, 1));
