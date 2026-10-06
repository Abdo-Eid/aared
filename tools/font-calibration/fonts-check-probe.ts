await using view = new Bun.WebView({ width: 400, height: 300 });
await view.navigate("data:text/html,<!doctype html><html><body></body></html>");
console.log(await view.evaluate(`(() => JSON.stringify({
  arial: document.fonts.check('16px "Arial"'),
  timesNewRoman: document.fonts.check('16px "Times New Roman"'),
  nonsense: document.fonts.check('16px "Definitely Not A Real Font Zzqx"'),
  calibri: document.fonts.check('16px "Calibri"'),
  simplifiedArabic: document.fonts.check('16px "Simplified Arabic"'),
}))()`));
