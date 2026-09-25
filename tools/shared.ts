import { basename, extname, join } from "node:path";

const PYMUPDF_PIN = "PyMuPDF==1.27.2";

/**
 * Spawn a command and collect its output. Used by Word, PowerPoint, and Excel
 * export tools. Throws if exit code is non-zero, including stderr in the error.
 */
export async function command(label: string, command: string, args: string[]): Promise<string> {
  const child = Bun.spawn([command, ...args], { stdout: "pipe", stderr: "pipe" });
  const [exitCode, stdout, stderr] = await Promise.all([
    child.exited,
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
  ]);
  if (exitCode === 0) return stdout;
  const detail = [stderr.trim(), stdout.trim()].filter(Boolean).join("\n").slice(0, 4_000);
  throw new Error(
    `aared ${label} export: native export failed with exit code ${exitCode}${detail === "" ? "" : `:\n${detail}`}`,
  );
}

/**
 * Derive a clean output directory from a fixture file path by keeping only
 * alphanumeric characters, dots, hyphens, and underscores in the stem.
 */
export function outputDirectory(fixturePath: string, outputRoot: string): string {
  const stem = basename(fixturePath, extname(fixturePath)).replace(/[^A-Za-z0-9._-]+/g, "-");
  if (stem === "") throw new Error(`aared export: fixture path has no usable file name`);
  return join(outputRoot, stem);
}

/**
 * Verify that the required Python interpreter with PyMuPDF is installed,
 * and return its path and version.
 */
export async function requireRasterizer(label: string): Promise<{ python: string; version: string }> {
  const python = Bun.which("python") ?? Bun.which("python3");
  const missing = `aared ${label} export: Python with ${PYMUPDF_PIN} is required to rasterize the PDF (pip install "${PYMUPDF_PIN}")`;
  if (python === null) throw new Error(missing);
  let version: string;
  try {
    version = (await command(label, python, ["-c", "import fitz; print(fitz.VersionBind)"])).trim();
  } catch {
    throw new Error(missing);
  }
  if (version === "") throw new Error(`aared ${label} export: PyMuPDF did not report a version`);
  return { python, version };
}
