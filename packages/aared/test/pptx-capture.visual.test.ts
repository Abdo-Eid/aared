import { expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { open } from "../src/pptx/index.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const RETAIN_ARTIFACTS = process.env.AARED_VISUAL_ARTIFACTS === "1";
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

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

async function within<T>(label: string, task: Promise<T>, milliseconds = 10_000): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      task,
      new Promise<T>((_, reject) => {
        timeout = setTimeout(() => reject(new Error(`aared visual review: timed out while ${label}`)), milliseconds);
      }),
    ]);
  } finally {
    if (timeout !== undefined) clearTimeout(timeout);
  }
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

function pngSize(bytes: Uint8Array): { width: number; height: number } {
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (!signature.every((value, index) => bytes[index] === value)) {
    throw new Error("aared visual review: WebView capture was not a PNG");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { width: view.getUint32(16), height: view.getUint32(20) };
}

async function writeReview(output: PowerPointRender, fixtureHash: string): Promise<void> {
  const cards = output.pages.map((_, index) => `
    <section>
      <h2>Slide ${index + 1}</h2>
      <div class="pair">
        <figure><figcaption>PowerPoint ${output.powerpointVersion}</figcaption><img src="powerpoint-slide-${index + 1}.png"></figure>
        <figure><figcaption>aared DOM/SVG</figcaption><img src="aared-slide-${index + 1}.png"></figure>
      </div>
    </section>`).join("\n");
  await Promise.all([
    writeFile(join(output.outputDir, "review.html"), `<!doctype html>
<meta charset="utf-8"><title>aared visual review</title>
<style>body{margin:24px;background:#f4f4f4;color:#111;font:16px system-ui,sans-serif}h1{margin:0 0 8px}.pair{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}figure{margin:0;background:#fff;padding:12px}figcaption{font-weight:700;margin-bottom:8px}img{display:block;width:100%;height:auto}@media(max-width:900px){.pair{grid-template-columns:1fr}}</style>
<h1>aared visual review</h1><p>PowerPoint export and aared capture are both ${output.width} × ${output.height}px.</p>${cards}`),
    writeFile(join(output.outputDir, "review.json"), JSON.stringify({
      fixture: "arabic-pptx-rtl-coverage.pptx",
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
    try {
      for (let index = 0; index < deck.info.count; index++) {
        slides.push(await deck.slides.get(index).read());
      }
      expect(reference.slideCount).toBe(deck.info.count);
    } finally {
      deck.close();
    }

    const pptx = await browserRenderer();
    const server = Bun.serve({
      port: 0,
      fetch(request) {
        if (new URL(request.url).pathname === "/pptx.mjs") {
          return new Response(pptx, { headers: { "content-type": "text/javascript; charset=utf-8" } });
        }
        return new Response(`<!doctype html><style>html,body,#host{margin:0;width:100%;height:100%;overflow:hidden}</style><main id="host"></main><script type="module">
          import { render } from "/pptx.mjs";
          window.renderSlide = async (model) => {
            await render({ slide: model, target: document.querySelector("#host") });
            await document.fonts.ready;
            await new Promise(requestAnimationFrame);
            const stage = document.querySelector('[data-aared-element="slide"]');
            return { width: stage.clientWidth, height: stage.clientHeight, canvasCount: document.querySelectorAll("canvas").length };
          };
        </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
      },
    });
    try {
      await using view = new Bun.WebView({ width: reference.width, height: reference.height });
      await within("loading the WebView fixture host", view.navigate(server.url.href));
      await within("sizing the WebView viewport", view.resize(reference.width, reference.height));
      for (const [index, slide] of slides.entries()) {
        const output = await within(`rendering slide ${index + 1}`, view.evaluate(`renderSlide(${JSON.stringify(slide)})`)) as {
          width: number;
          height: number;
          canvasCount: number;
        };
        expect(output.canvasCount).toBe(0);
        expect(output.width).toBe(reference.width);
        expect(output.height).toBe(reference.height);
        expect(output.width / output.height).toBeCloseTo(firstSlideSize.width / firstSlideSize.height, 2);
        const screenshot = await within(`capturing slide ${index + 1}`, view.screenshot({ encoding: "buffer", format: "png" }));
        expect(pngSize(screenshot)).toEqual({ width: reference.width, height: reference.height });
        if (RETAIN_ARTIFACTS) {
          await writeFile(join(reference.outputDir, `aared-slide-${index + 1}.png`), screenshot);
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
}, 45_000);
