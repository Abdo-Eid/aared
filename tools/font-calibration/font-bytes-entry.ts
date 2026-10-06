import { render as renderPptx } from "../../../packages/aared/src/pptx/index.ts";
import { render as renderDocx } from "../../../packages/aared/src/docx/index.ts";
import { render as renderXlsx } from "../../../packages/aared/src/xlsx/index.ts";
import { collectStrings } from "../../../packages/aared/src/core/collect.ts";
import { ARABIC_UNICODE_RANGE, createResolver } from "../../../packages/aared-fonts/src/index.ts";
import { simulateAbsentFonts } from "../../../packages/aared-fonts/test/absent-fonts.browser.ts";

// Every delivered face, with the file it would fetch: recorded by a FontFace subclass installed
// as the global before anything constructs one.
const faces: Array<{ face: FontFace; url: string; arabicOnly: boolean }> = [];
const Native = globalThis.FontFace;
class Recording extends Native {
  constructor(family: string, source: string | BufferSource, descriptors?: FontFaceDescriptors) {
    super(family, source, descriptors);
    if (typeof source === "string" && source.startsWith("url(")) faces.push({ face: this, url: source.slice(4, -1), arabicOnly: descriptors?.unicodeRange === ARABIC_UNICODE_RANGE });
  }
}
globalThis.FontFace = Recording as typeof FontFace;

const raf = () => new Promise(requestAnimationFrame);

(window as unknown as Record<string, unknown>).run = async (format: string, input: unknown, absent: boolean, eager: boolean) => {
  let model = input;
  if (absent) model = await simulateAbsentFonts(model);
  else {
    // What the session does for a real open(): the resolver, once per family the model names.
    const families = new Set<string>();
    collectStrings(model, "fontFamily", families);
    collectStrings(model, "typeface", families);
    const resolver = createResolver();
    await Promise.all([...families].map((family) => resolver(family)));
  }
  // The behaviour before: every non-Arabic face loaded as soon as it was registered.
  if (eager) for (const entry of faces) if (!entry.arabicOnly) entry.face.load().catch(() => undefined);
  const target = document.querySelector("#host") as HTMLElement;
  if (format === "pptx") await renderPptx({ slide: model as never, target });
  else if (format === "docx") await renderDocx({ document: model as never, target, pagination: { mode: "layout" } });
  else await renderXlsx({ sheet: model as never, target });
  await raf(); await raf();
  await document.fonts.ready;
  await new Promise((resolve) => setTimeout(resolve, 500));
  await document.fonts.ready;
  return JSON.stringify([...new Set(faces.filter((entry) => entry.face.status !== "unloaded").map((entry) => entry.url))]);
};
