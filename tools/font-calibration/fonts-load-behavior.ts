// Does document.fonts.load() resolve immediately for a family with no
// registered FontFace (a real installed system font, or a name nobody
// registered)? And does it correctly wait for a real, slow-registered face?
await using view = new Bun.WebView({ width: 400, height: 300 });
await view.navigate("data:text/html,<!doctype html><html><body></body></html>");
const result = await view.evaluate(`(async () => {
  const out = {};
  const t0 = performance.now();
  const a = await document.fonts.load("16px Calibri", "hello");
  out.unregisteredInstalled = { ms: performance.now() - t0, result: a.length };
  const t1 = performance.now();
  const b = await document.fonts.load("16px 'Nobody Registered This'", "hello");
  out.unregisteredMissing = { ms: performance.now() - t1, result: b.length };
  // A real registered face, with an artificial delay before .load() settles.
  const bytes = Uint8Array.from(atob(${JSON.stringify(Buffer.from(await Bun.file("C:/Windows/Fonts/Amiri-Regular.ttf").arrayBuffer()).toString("base64"))}), (c) => c.charCodeAt(0));
  const face = new FontFace("SlowFace", bytes.buffer.slice(0));
  document.fonts.add(face);
  const loadPromise = face.load();
  const t2 = performance.now();
  const c = await document.fonts.load("16px SlowFace", "hello");
  out.registeredSlow = { ms: performance.now() - t2, result: c.length, faceStatusAfter: face.status };
  await loadPromise;
  return JSON.stringify(out);
})()`);
console.log(result);
