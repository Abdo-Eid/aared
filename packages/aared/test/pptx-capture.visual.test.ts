import { expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { open } from "../src/pptx/index.js";
import { ABSENT_FONTS, absentFontsModule } from "./absent-fonts.js";
import { pngSize, within } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const RETAIN_ARTIFACTS = process.env.AARED_VISUAL_ARTIFACTS === "1";
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

interface PowerPointRender {
  outputDir: string;
  powerpointVersion: string;
  slideCount: number;
  width: number;
  height: number;
  pages: string[];
}

async function renderPowerPoint(): Promise<PowerPointRender> {
  const child = Bun.spawn(["bun", "tools/powerpoint-render/index.ts", FIXTURE_PATH, "--width", "1280"], {
    cwd: `${import.meta.dir}/../../..`,
    stdout: "pipe",
    stderr: "pipe",
  });
  const [exitCode, stdout, stderr] = await Promise.all([
    child.exited,
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
  ]);
  if (exitCode !== 0) {
    throw new Error(`aared visual review: PowerPoint export failed with exit code ${exitCode}: ${(stderr || stdout).trim()}`);
  }
  return JSON.parse(stdout) as PowerPointRender;
}

async function browserRenderer(): Promise<string> {
  const result = await Bun.build({
    entrypoints: [`${import.meta.dir}/../src/pptx/index.ts`],
    target: "browser",
    format: "esm",
  });
  if (!result.success) throw new Error(result.logs.map((log) => log.message).join("\n"));
  return result.outputs[0]!.text();
}

async function writeReview(output: PowerPointRender, fixtureHash: string): Promise<void> {
  const cards = output.pages.map((_, index) => `
    <section>
      <h2>Slide ${index + 1}</h2>
      <div class="pair">
        <figure><figcaption>PowerPoint ${output.powerpointVersion}</figcaption><img src="powerpoint-slide-${index + 1}.png"></figure>
        <figure><figcaption>aared DOM/SVG</figcaption><img src="aared-slide-${index + 1}.png"></figure>${ABSENT_FONTS ? `
        <figure><figcaption>aared, no Office fonts (aared-fonts substitutes)</figcaption><img src="aared-absent-slide-${index + 1}.png"></figure>` : ""}
      </div>
    </section>`).join("\n");
  await Promise.all([
    writeFile(join(output.outputDir, "review.html"), `<!doctype html>
<meta charset="utf-8"><title>aared visual review</title>
<style>body{margin:24px;background:#f4f4f4;color:#111;font:16px system-ui,sans-serif}h1{margin:0 0 8px}.pair{display:grid;grid-template-columns:repeat(${ABSENT_FONTS ? 3 : 2},minmax(0,1fr));gap:16px}figure{margin:0;background:#fff;padding:12px}figcaption{font-weight:700;margin-bottom:8px}img{display:block;width:100%;height:auto}@media(max-width:900px){.pair{grid-template-columns:1fr}}</style>
<h1>aared visual review</h1><p>PowerPoint export and aared capture are both ${output.width} × ${output.height}px.</p>${cards}`),
    writeFile(join(output.outputDir, "review.json"), JSON.stringify({
      fixture: "pptx-coverage.pptx",
      fixtureSha256: fixtureHash,
      powerpointVersion: output.powerpointVersion,
      slideCount: output.slideCount,
      width: output.width,
      height: output.height,
    }, null, 2)),
  ]);
}

visual("visual review: PPTX produces same-size PowerPoint and aared slide captures", async () => {
  const reference = await renderPowerPoint();
  try {
    const deck = await open(await Bun.file(FIXTURE_PATH).arrayBuffer());
    const firstSlideSize = deck.info.sizes?.[0];
    if (firstSlideSize === undefined) throw new Error("aared visual review: PPTX Worker did not report slide dimensions");
    const slides: unknown[] = [];
    let media: Record<string, string>;
    try {
      for (let index = 0; index < deck.info.count; index++) {
        slides.push(await deck.slides.get(index).read());
      }
      expect(reference.slideCount).toBe(deck.info.count);
      // Every media part the models name (pictures, backgrounds, picture
      // bullets), fetched through the session before it closes.
      const ids = new Set<string>();
      for (const match of JSON.stringify(slides).matchAll(/"(?:mediaId|pictureMediaId)":"([^"]+)"/g)) ids.add(match[1]!);
      media = {};
      for (const id of ids) media[id] = Buffer.from(await deck.media.get(id)).toString("base64");
    } finally {
      deck.close();
    }

    const pptx = await browserRenderer();
    const absentFonts = ABSENT_FONTS ? await absentFontsModule() : "";
    const server = Bun.serve({
      port: 0,
      fetch(request) {
        if (new URL(request.url).pathname === "/pptx.mjs") {
          return new Response(pptx, { headers: { "content-type": "text/javascript; charset=utf-8" } });
        }
        if (new URL(request.url).pathname === "/absent-fonts.mjs") {
          return new Response(absentFonts, { headers: { "content-type": "text/javascript; charset=utf-8" } });
        }
        return new Response(`<!doctype html><style>html,body,#host{margin:0;width:100%;height:100%;overflow:hidden}</style><main id="host"></main><script type="module">
          import { render } from "/pptx.mjs";
          ${ABSENT_FONTS ? 'import "/absent-fonts.mjs";' : ""}
          const decode = (base64) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)).buffer;
          window.renderSlide = async (model, media, absent) => {
            if (absent) model = await window.simulateAbsentFonts(model);
            await render({
              slide: model,
              target: document.querySelector("#host"),
              // Detached models carry no session, so the paired review feeds
              // the fixture's own bytes: without this every picture (ink
              // fallbacks, photos) paints a placeholder and the PowerPoint
              // side of the pair has nothing honest to compare against.
              mediaResolver: media === undefined ? undefined : async (mediaId) => {
                const base64 = media[mediaId];
                if (base64 === undefined) throw new Error("missing media " + mediaId);
                return decode(base64);
              },
            });
            await document.fonts.ready;
            // Pictures resolve asynchronously; the screenshot must wait for
            // blob URLs and decoded images, not a fixed interval.
            await new Promise((resolve, reject) => {
              const deadline = Date.now() + 8000;
              const poll = () => {
                const images = [...document.querySelectorAll('[data-aared-element="picture-image"]')];
                // Picture fills (shapes, cells, backgrounds) paint once their
                // image decodes; the fill painter marks them until then.
                const fillsLoading = document.querySelector("[data-aared-fill-loading]") !== null;
                if (!fillsLoading && images.every((img) => img.src.startsWith("blob:") && img.complete && img.naturalWidth > 0)) resolve(undefined);
                else if (Date.now() > deadline) reject(new Error("timed out waiting for slide images"));
                else setTimeout(poll, 50);
              };
              poll();
            });
            const stage = document.querySelector('[data-aared-element="slide"]');
            return { width: stage.clientWidth, height: stage.clientHeight, canvasCount: document.querySelectorAll("canvas").length };
          };
        </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
      },
    });
    try {
      await using view = new Bun.WebView({ width: reference.width, height: reference.height });
      await within("loading the WebView fixture host", view.navigate(server.url.href), 10_000);
      await within("sizing the WebView viewport", view.resize(reference.width, reference.height), 10_000);
      for (const [index, slide] of slides.entries()) {
        const output = await within(`rendering slide ${index + 1}`, view.evaluate(`renderSlide(${JSON.stringify(slide)}, ${JSON.stringify(media)})`), 15_000) as {
          width: number;
          height: number;
          canvasCount: number;
        };
        expect(output.canvasCount).toBe(0);
        expect(output.width).toBe(reference.width);
        expect(output.height).toBe(reference.height);
        expect(output.width / output.height).toBeCloseTo(firstSlideSize.width / firstSlideSize.height, 2);
        const screenshot = await within(`capturing slide ${index + 1}`, view.screenshot({ encoding: "buffer", format: "png" }), 10_000);
        expect(pngSize(screenshot)).toEqual({ width: reference.width, height: reference.height });
        if (RETAIN_ARTIFACTS) {
          await writeFile(join(reference.outputDir, `aared-slide-${index + 1}.png`), screenshot);
        }
        if (ABSENT_FONTS) {
          await within(`rendering slide ${index + 1} without Office fonts`, view.evaluate(`renderSlide(${JSON.stringify(slide)}, ${JSON.stringify(media)}, true)`), 30_000);
          const absent = await within(`capturing slide ${index + 1} without Office fonts`, view.screenshot({ encoding: "buffer", format: "png" }), 10_000);
          await writeFile(join(reference.outputDir, `aared-absent-slide-${index + 1}.png`), absent);
        }
      }
      if (RETAIN_ARTIFACTS) {
        const fixtureHash = createHash("sha256").update(new Uint8Array(await Bun.file(FIXTURE_PATH).arrayBuffer())).digest("hex");
        await writeReview(reference, fixtureHash);
      }
    } finally {
      server.stop();
    }
  } finally {
    if (!RETAIN_ARTIFACTS) await rm(reference.outputDir, { recursive: true, force: true });
  }
}, ABSENT_FONTS ? 300_000 : 45_000);
