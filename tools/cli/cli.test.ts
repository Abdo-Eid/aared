import { describe, expect, test } from "bun:test";
import { join, resolve } from "node:path";
import { mkdir, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { parseToolCommand } from "./options.js";
import { parseParts } from "../review/office-export.js";
import { partSlug, outputDirectory } from "../shared/office.js";
import { reviewPage } from "../review/page.js";
import { benchmarkExitCode, validateBenchmarkArgs, sameBenchmarkMachine } from "../benchmarks/measure.js";
import { REPO_ROOT, TOOLS_ROOT, OUTPUT_ROOT, reviewOutputRoot } from "../shared/paths.js";

const CLI = join(TOOLS_ROOT, "cli.ts");

async function invoke(args: string[], cwd = import.meta.dir): Promise<{ code: number; out: string; err: string }> {
  const child = Bun.spawn([process.execPath, CLI, ...args], { cwd, stdout: "pipe", stderr: "pipe" });
  const [code, out, err] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()]);
  return { code, out, err };
}

describe("tools command contracts", () => {
  test("tool implementations stay grouped and generated output has one ignored home", async () => {
    const root = await readdir(TOOLS_ROOT, { withFileTypes: true });
    expect(root.filter((entry) => entry.isFile()).map((entry) => entry.name).sort()).toEqual(["README.md", "cli.ts", "tsconfig.json"]);
    for (const dir of ["cli", "benchmarks", "review", "checks", "inspect", "shared", "chart-oracle", "font-calibration", "probe"]) {
      expect(root.some((entry) => entry.isDirectory() && entry.name === dir)).toBe(true);
    }
    expect(OUTPUT_ROOT).toBe(join(TOOLS_ROOT, "out"));
    for (const format of ["pptx", "xlsx", "docx"] as const) expect(reviewOutputRoot(format)).toBe(join(OUTPUT_ROOT, format));
    const ignore = await Bun.file(join(REPO_ROOT, ".gitignore")).text();
    expect(ignore).toContain("tools/out/");
    expect(ignore).not.toContain("tools/bench-results/");
  });

  test("help works from outside the repo without launching a backend", async () => {
    const result = await invoke(["review", "--help"], tmpdir());
    expect(result.code).toBe(0);
    expect(result.out).toContain("--absent-fonts");
    expect(result.out).not.toContain("native export and");
    expect(result.err).toBe("");
  });

  test("all format routes resolve file paths against the caller's directory", () => {
    for (const format of ["pptx", "xlsx", "docx"] as const) {
      const command = parseToolCommand(["review", `file with spaces.${format}`, "--absent-fonts"], tmpdir());
      expect(command.kind).toBe("review");
      if (command.kind !== "review") throw new Error("expected review");
      expect(command.options.file).toBe(resolve(tmpdir(), `file with spaces.${format}`));
      expect(command.options.format).toBe(format);
      expect(command.options.absentFonts).toBe(true);
    }
  });

  test("review validates format-specific flags and malformed options", () => {
    for (const args of [
      ["review"], ["review", "a.pdf"], ["review", "a.docx", "--part", "1:top=A1:B2"],
      ["review", "a.xlsx", "--width", "1280"], ["review", "a.pptx", "--pair-full"],
      ["review", "a.pptx", "--width", "0"], ["review", "a.pptx", "--width"],
      ["review", "a.docx", "--absent-font"], ["review", "a.xlsx", "--part", "1:top=B1:G10"],
      ["review", "a.xlsx", "--part", "1:top=A10:G1"],
    ]) expect(() => parseToolCommand(args)).toThrow();
    const command = parseToolCommand(["review", "a.xlsx", "--part", "1:top band=A1:G10", "--part", "2:other=A1:D3", "--pair-full"]);
    expect(command.kind).toBe("review");
    if (command.kind === "review") {
      expect(command.options.parts).toHaveLength(2);
      expect(command.options.pairFull).toBe(true);
    }
  });

  test("part filename rules match native export and prevent collisions", () => {
    expect(partSlug(" --top band-- ")).toBe("top-band");
    expect(() => partSlug("!!!")).toThrow();
    expect(() => parseParts(["--part", "1:top band=A1:G10", "--part", "1:top-band=A11:G12"])).toThrow("same file name");
    expect(() => parseParts(["--part", "1:Top=A1:G10", "--part", "1:top=A11:G12"])).toThrow("same file name");
    expect(() => parseParts(["--part"])).toThrow();
  });

  test("filename-derived cleanup cannot resolve to an artifact root or its parent", () => {
    for (const file of ["..pptx", "...pptx", ".", ".."]) expect(() => outputDirectory(file, tmpdir())).toThrow();
    expect(outputDirectory("file name.pptx", tmpdir())).toBe(join(tmpdir(), "file-name"));
  });

  test("benchmark actions preserve measurement identity and explicit report targets", () => {
    expect(parseToolCommand(["bench", "run", "pptx-coverage.pptx:aared"])).toEqual({ kind: "bench", args: ["--run", "pptx-coverage.pptx:aared"] });
    expect(parseToolCommand(["bench", "report"])).toEqual({ kind: "bench", args: ["--render-md", "--render-site"] });
    expect(parseToolCommand(["bench", "report", "--site"])).toEqual({ kind: "bench", args: ["--render-site"] });
    for (const args of [["--run", "no-such-fixture:aared"], ["--run", "--size"], ["--size", "--render-md"], ["--search", "--all"], ["--run", "--al"]]) {
      expect(() => validateBenchmarkArgs(args)).toThrow();
    }
  });

  test("benchmark failures and empty passes cannot report shell success", () => {
    expect(benchmarkExitCode([{ status: "measured" }])).toBe(0);
    expect(benchmarkExitCode([{ status: "error" }])).toBe(1);
    expect(benchmarkExitCode([])).toBe(1);
    expect(benchmarkExitCode([{ status: "measured" }], 1)).toBe(1);
    const machine = { cpu: "test CPU", os: "test OS", cores: 8, ram_gb: 16, platform: "win32", arch: "x64" };
    expect(sameBenchmarkMachine(machine, { ...machine })).toBe(true);
    expect(sameBenchmarkMachine(machine, { ...machine, cpu: "another CPU" })).toBe(false);
  });

  test("invalid commands fail before browser launch", async () => {
    for (const args of [["unknown"], ["toString"], ["bench", "run", "unknown:aared"], ["review", "a.docx", "--typo"]]) {
      const result = await invoke(args);
      expect(result.code).toBe(1);
      expect(result.err).toContain("aared tools:");
      expect(result.out).toBe("");
    }
  });

  test("package declaration builds exclude tooling and tests", async () => {
    const config = await Bun.file(join(REPO_ROOT, "packages/aared/tsconfig.build.json")).json();
    expect(config.include).toEqual(["src"]);
    expect(config.compilerOptions.rootDir).toBe(".");
    const build = await Bun.file(join(REPO_ROOT, "packages/aared/tsdown.config.ts")).text();
    expect(build).toContain('dts: { tsconfig: "tsconfig.build.json" }');
    for (const file of ["review/office-export.d.ts", "shared/office.d.ts", "review/page.d.ts", "review/pptx.d.ts"]) {
      expect(await Bun.file(join(TOOLS_ROOT, file)).exists()).toBe(false);
    }
  });

  test("missing review input fails without exporting", async () => {
    const result = await invoke(["review", "missing-review-file.pptx"]);
    expect(result.code).toBe(1);
    expect(result.err).toContain("does not exist");
    expect(result.out).toBe("");
  });

  test("inspect rejects conflicting output/memory modes", () => {
    expect(() => parseToolCommand(["inspect", "file.docx", "--out", "result.json", "--bench"])).toThrow();
    const command = parseToolCommand(["inspect", "file.docx", "--out", "result.json"], tmpdir());
    expect(command).toEqual({ kind: "inspect", file: resolve(tmpdir(), "file.docx"), output: resolve(tmpdir(), "result.json"), bench: false });
  });

  test("packed consumers resolve tarballs against the caller and reject malformed options", () => {
    expect(parseToolCommand(["check", "consumer-build"])).toEqual({ kind: "check" });
    expect(parseToolCommand(["check", "consumer-build", "--package", "viewer release.tgz"], tmpdir()))
      .toEqual({ kind: "check", package: resolve(tmpdir(), "viewer release.tgz") });
    for (const args of [
      ["check", "consumer-build", "--package"], ["check", "consumer-build", "--package", "package.json"],
      ["check", "consumer-build", "--typo"], ["check", "unknown"],
    ]) expect(() => parseToolCommand(args)).toThrow();
  });

  test("report shell escapes filenames and supports different column counts", () => {
    const page = reviewPage('<file & "name">', '<section class="pair"><figure>one</figure><figure>two</figure></section>');
    expect(page).toContain("&lt;file &amp; &quot;name&quot;&gt;");
    expect(page).toContain("flex-direction:column");
    expect(page).not.toContain("grid-auto-flow");
  });

  test("Word review keeps page-count differences visible instead of suppressing the report", async () => {
    await mkdir(join(tmpdir(), "opencode"), { recursive: true });
    const dir = await mkdtemp(join(tmpdir(), "opencode", "aared-tool-review-"));
    try {
      await Promise.all(["word-page-1.png", "aared-page-1.png", "aared-page-2.png"].map((name) => writeFile(join(dir, name), "report test placeholder")));
      const child = Bun.spawn([process.execPath, join(TOOLS_ROOT, "review/docx/review.ts"), dir], { stdout: "pipe", stderr: "pipe" });
      const [code, , err] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()]);
      expect(code).toBe(0);
      expect(err).toBe("");
      const html = await Bun.file(join(dir, "review.html")).text();
      expect(html).toContain("pages: Word 1, aared 2");
      expect(html).toContain("aared-page-2.png");
      expect(html).toContain("No corresponding page.");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
