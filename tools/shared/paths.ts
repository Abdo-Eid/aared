import { join, resolve } from "node:path";

/** Stable homes shared by CLI, review and benchmark backends; never depend on cwd. */
export const REPO_ROOT = resolve(import.meta.dir, "../..");
export const TOOLS_ROOT = join(REPO_ROOT, "tools");
export const OUTPUT_ROOT = join(TOOLS_ROOT, "out");

export function reviewOutputRoot(format: "pptx" | "xlsx" | "docx"): string {
  return join(OUTPUT_ROOT, format);
}
