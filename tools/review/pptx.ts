import { createHash } from "node:crypto";
import { writeFile } from "node:fs/promises";
import { basename, join } from "node:path";
import { open } from "../../packages/aared/src/pptx/index.js";
import { collectStrings } from "../../packages/aared/src/core/collect.js";
import { absentFontsModule } from "../../packages/aared/test/absent-fonts.js";
import { pngSize, within } from "../../packages/aared/test/harness.js";
import type { PowerPointRender } from "./office-export.js";
import { reviewPage, escapeHtml } from "./page.js";

export interface SlideCapture {
  width: number;
  height: number;
  canvasCount: number;
  /** Painted scene size in CSS px, independent of the forced screenshot viewport. */
  slide: { width: number; height: number };
  /** Public document size in inches for this slide. */
  expectedSize: { width: number; height: number };
  png: { width: number; height: number };
}

function assertSlideProportions(slide: SlideCapture["slide"], expected: SlideCapture["expectedSize"], index: number): void {
  const ratio = slide.width / slide.height;
  if (!(slide.width > 0 && slide.height > 0) || !Number.isFinite(ratio)
    || Math.abs(ratio - expected.width / expected.height) >= 0.005) {
    throw new Error(`aared PPTX review: slide ${index + 1} does not match its public size proportions`);
  }
}

/** The capture formerly embedded in the PPTX visual test; CLI and suite use one path. */
export async function capturePptx(
  fixture: string,
  reference: PowerPointRender,
  options: { absentFonts?: boolean; retainArtifacts?: boolean } = {},
): Promise<SlideCapture[]> {
  const absent = options.absentFonts ?? false;
  const retain = options.retainArtifacts ?? true;
  const bytes = await Bun.file(fixture).arrayBuffer();
  const deck = await open(bytes);
  const sizes = deck.info.sizes;
  const slides = [];
  const media: Record<string, string> = {};
  try {
    if (deck.info.count !== reference.slideCount) throw new Error("aared PPTX review: native and parsed slide counts differ");
    if (sizes === undefined || sizes.length !== deck.info.count) throw new Error("aared PPTX review: public slide sizes are missing");
    for (const handle of deck.slides.list()) slides.push(await handle.read());
    const ids = new Set<string>();
    for (const key of ["mediaId", "pictureMediaId", "fallbackMediaId", "pictureFallbackMediaId"]) collectStrings(slides, key, ids);
    for (const id of ids) media[id] = Buffer.from(await deck.media.get(id)).toString("base64");
  } finally {
    deck.close();
  }

  const build = await Bun.build({
    entrypoints: [join(import.meta.dir, "../../packages/aared/src/pptx/index.ts")],
    target: "browser", format: "esm",
  });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  const module = await build.outputs[0]!.text();
  const absentModule = absent ? await absentFontsModule() : "";
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const path = new URL(request.url).pathname;
      if (path === "/pptx.mjs") return new Response(module, { headers: { "content-type": "text/javascript" } });
      if (path === "/absent-fonts.mjs") return new Response(absentModule, { headers: { "content-type": "text/javascript" } });
      return new Response(`<!doctype html><style>html,body,#host{margin:0;width:100%;height:100%;overflow:hidden}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        ${absent ? 'import "/absent-fonts.mjs";' : ""}
        let handle;
        const decode = (base64) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)).buffer;
        window.renderSlide = async (model, media, absent) => {
          handle?.destroy();
          if (absent) model = await window.simulateAbsentFonts(model);
          handle = await render({
            slide: model, target: document.querySelector("#host"),
            mediaResolver: async (id) => {
              if (media[id] === undefined) throw new Error("missing media " + id);
              return decode(media[id]);
            },
          });
          await document.fonts.ready;
          // Detached models still use the file's own media. Wait for decoded
          // pictures/fills rather than screenshotting unresolved placeholders.
          await new Promise((resolve, reject) => {
            const deadline = Date.now() + 8000;
            const poll = () => {
              const images = [...document.querySelectorAll('[data-aared-element="picture-image"]')];
              const loading = document.querySelector("[data-aared-fill-loading]");
              if (!loading && images.every((img) => img.src.startsWith("blob:") && img.complete && img.naturalWidth > 0)) resolve();
              else if (Date.now() > deadline) reject(new Error("timed out waiting for slide images"));
              else setTimeout(poll, 50);
            };
            poll();
          });
          const stage = document.querySelector('[data-aared-element="slide"]');
          const scene = stage.querySelector('[data-aared-element="slide-scene"]').getBoundingClientRect();
          return { width: stage.clientWidth, height: stage.clientHeight,
            slide: { width: scene.width, height: scene.height }, canvasCount: document.querySelectorAll("canvas").length };
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  const captures: SlideCapture[] = [];
  try {
    await using view = new Bun.WebView({ width: reference.width, height: reference.height });
    await within("loading the PPTX review host", view.navigate(server.url.href), 10_000);
    await within("sizing the PPTX review viewport", view.resize(reference.width, reference.height), 10_000);
    for (const [index, slide] of slides.entries()) {
      const output = await within(`rendering slide ${index + 1}`, view.evaluate(`renderSlide(${JSON.stringify(slide)}, ${JSON.stringify(media)})`), 15_000) as Omit<SlideCapture, "png" | "expectedSize">;
      const expectedSize = sizes![index]!;
      assertSlideProportions(output.slide, expectedSize, index);
      const screenshot = await within(`capturing slide ${index + 1}`, view.screenshot({ encoding: "buffer", format: "png" }), 10_000);
      const png = pngSize(screenshot);
      if (output.canvasCount !== 0 || output.width !== reference.width || output.height !== reference.height
        || png.width !== reference.width || png.height !== reference.height) {
        throw new Error(`aared PPTX review: slide ${index + 1} is not a same-size DOM/SVG capture`);
      }
      captures.push({ ...output, expectedSize: { ...expectedSize }, png });
      if (retain) await writeFile(join(reference.outputDir, `aared-slide-${index + 1}.png`), screenshot);
      if (absent) {
        const absentOutput = await within(`rendering slide ${index + 1} without Office fonts`, view.evaluate(`renderSlide(${JSON.stringify(slide)}, ${JSON.stringify(media)}, true)`), 30_000) as Omit<SlideCapture, "png" | "expectedSize">;
        assertSlideProportions(absentOutput.slide, expectedSize, index);
        if (retain) await writeFile(join(reference.outputDir, `aared-absent-slide-${index + 1}.png`), await view.screenshot({ encoding: "buffer", format: "png" }));
      }
    }
  } finally {
    server.stop();
  }
  if (retain) {
    const cards = captures.map((_, index) => `<section><h2>Slide ${index + 1}</h2><div class="pair">
      <figure><figcaption>PowerPoint ${escapeHtml(reference.powerpointVersion)}</figcaption><img src="powerpoint-slide-${index + 1}.png"></figure>
      <figure><figcaption>aared DOM/SVG</figcaption><img src="aared-slide-${index + 1}.png"></figure>${absent ? `
      <figure><figcaption>aared, Office fonts absent</figcaption><img src="aared-absent-slide-${index + 1}.png"></figure>` : ""}
    </div></section>`).join("\n");
    await Promise.all([
      writeFile(join(reference.outputDir, "review.html"), reviewPage(`aared PPTX review - ${basename(fixture)}`,
        `<p>PowerPoint export and aared capture are both ${reference.width} x ${reference.height}px. Human review, never pixel diffs.</p>${cards}`)),
      writeFile(join(reference.outputDir, "review.json"), JSON.stringify({
        fixture: basename(fixture), fixtureSha256: createHash("sha256").update(new Uint8Array(bytes)).digest("hex"),
        powerpointVersion: reference.powerpointVersion, slideCount: reference.slideCount,
        width: reference.width, height: reference.height,
      }, null, 2)),
    ]);
  }
  return captures;
}
