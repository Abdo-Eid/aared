import { readdir, rm } from "node:fs/promises";
import { basename, extname, join, resolve } from "node:path";

const OUTPUT_ROOT = join(import.meta.dir, "out");

export interface PowerPointRender {
  outputDir: string;
  powerpointVersion: string;
  slideCount: number;
  width: number;
  height: number;
  pages: string[];
}

function outputDirectory(fixturePath: string): string {
  const stem = basename(fixturePath, extname(fixturePath)).replace(/[^A-Za-z0-9._-]+/g, "-");
  if (stem === "") throw new Error("aared PowerPoint export: fixture path has no usable file name");
  return join(OUTPUT_ROOT, stem);
}

async function command(command: string, args: string[]): Promise<string> {
  const child = Bun.spawn([command, ...args], { stdout: "pipe", stderr: "pipe" });
  const [exitCode, stdout, stderr] = await Promise.all([
    child.exited,
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
  ]);
  if (exitCode === 0) return stdout;
  const detail = [stderr.trim(), stdout.trim()].filter(Boolean).join("\n").slice(0, 4_000);
  throw new Error(`aared PowerPoint export: native export failed with exit code ${exitCode}${detail === "" ? "" : `:\n${detail}`}`);
}

/** Export a real PPTX with the installed PowerPoint renderer. Windows only. */
export async function renderPowerPoint(fixturePath: string, width = 1280): Promise<PowerPointRender> {
  if (process.platform !== "win32") {
    throw new Error("aared PowerPoint export: native PowerPoint export requires Windows with Microsoft PowerPoint installed");
  }
  if (!Number.isInteger(width) || width <= 0) {
    throw new Error("aared PowerPoint export: width must be a positive integer");
  }
  const pwsh = Bun.which("pwsh");
  if (pwsh === null) throw new Error("aared PowerPoint export: pwsh is required to automate Microsoft PowerPoint");

  const fixture = resolve(fixturePath);
  const outputDir = outputDirectory(fixture);
  await rm(outputDir, { recursive: true, force: true });
  await command(pwsh, ["-NoProfile", "-NonInteractive", "-STA", "-File", join(import.meta.dir, "export.ps1"), "-InputPath", fixture, "-OutputDir", outputDir, "-Width", String(width)]);

  const metadata = await Bun.file(join(outputDir, "powerpoint-export.json")).json() as Omit<PowerPointRender, "outputDir" | "pages">;
  const pages = (await readdir(outputDir))
    .filter((name) => /^powerpoint-slide-\d+\.png$/.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
    .map((name) => join(outputDir, name));
  if (pages.length !== metadata.slideCount) {
    throw new Error(`aared PowerPoint export: expected ${metadata.slideCount} PNGs but found ${pages.length}`);
  }
  return { outputDir, ...metadata, pages };
}

if (import.meta.main) {
  const [fixture, ...args] = process.argv.slice(2);
  if (fixture === undefined || args.some((arg) => arg !== "--width" && !/^\d+$/.test(arg))) {
    throw new Error("Usage: bun tools/powerpoint-render/index.ts <fixture.pptx> [--width <pixels>]");
  }
  const widthIndex = args.indexOf("--width");
  const width = widthIndex === -1 ? 1280 : Number(args[widthIndex + 1]);
  if (widthIndex !== -1 && args[widthIndex + 1] === undefined) {
    throw new Error("aared PowerPoint export: --width requires a pixel value");
  }
  console.log(JSON.stringify(await renderPowerPoint(fixture, width)));
}
