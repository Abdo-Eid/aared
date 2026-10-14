import { expect, test } from "bun:test";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { assertDocxReviewBuild } from "./browser-server.js";

test("Word capture preflight accepts the complete production build", async () => {
  await expect(assertDocxReviewBuild()).resolves.toBeUndefined();
});

test("Word capture preflight rejects missing, empty and partial builds before export", async () => {
  const scratch = join(tmpdir(), "opencode");
  await mkdir(scratch, { recursive: true });
  const dist = await mkdtemp(join(scratch, "aared-word-preflight-"));
  const core = join(dist, "core");
  try {
    await expect(assertDocxReviewBuild(dist)).rejects.toThrow("run bun run build first");
    await writeFile(core, "not a directory");
    await expect(assertDocxReviewBuild(dist)).rejects.toThrow("run bun run build first");
    await rm(core);
    await mkdir(core);
    await expect(assertDocxReviewBuild(dist)).rejects.toThrow("core entry is missing");
    await mkdir(join(core, "index.mjs"));
    await expect(assertDocxReviewBuild(dist)).rejects.toThrow("core entry is missing");
    await rm(join(core, "index.mjs"), { recursive: true });
    await writeFile(join(core, "index.mjs"), 'export { value } from "./shared.mjs";');
    await expect(assertDocxReviewBuild(dist)).rejects.toThrow("index.mjs -> ./shared.mjs");
    await writeFile(join(core, "shared.mjs"), "export const value = 1;");
    await expect(assertDocxReviewBuild(dist)).resolves.toBeUndefined();
  } finally {
    await rm(dist, { recursive: true, force: true });
  }
});
