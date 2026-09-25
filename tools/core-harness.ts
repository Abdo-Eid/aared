/**
 * Drives `core/worker.ts` in-process (Node/Bun worker message loop stub) and
 * dumps every unit/layout/media hash, or just timings in `bench` mode.
 *
 *   bun tools/core-harness.ts <file> [out-json] [bench]
 *
 * `bench` mode skips the dump and prints open/unit timings plus heap/RSS, so
 * large files can be timed without materializing the whole model as JSON.
 * Otherwise the full result is written to `out-json` (when given) and every
 * referenced layout/media part is fetched and hashed.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";

const [file, ...rest] = process.argv.slice(2);
if (file === undefined) {
  console.error("usage: bun tools/core-harness.ts <file> [out-json] [bench]");
  process.exit(1);
}
const mode = rest.includes("bench") ? "bench" : undefined;
const out = rest.find((arg) => arg !== "bench");
const g = globalThis as never as {
  postMessage: (message: unknown) => void;
  onmessage: ((event: { data: unknown }) => void) | null;
  gc?: () => void;
};
let waiting: ((message: never) => void) | undefined;
g.postMessage = (message: unknown) => {
  if ((message as { type?: string }).type === "ready") return;
  const w = waiting;
  waiting = undefined;
  w?.(message as never);
};
await import("../packages/aared/src/core/worker.ts");
let rid = 1;
function call(message: Record<string, unknown>): Promise<{
  ok: boolean;
  result?: never;
  error?: unknown;
}> {
  return new Promise((resolve) => {
    waiting = resolve;
    g.onmessage?.({ data: { ...message, sessionId: 1, requestId: rid++ } });
  });
}
const bytes = readFileSync(file);
const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const t0 = performance.now();
const info = await call({ type: "open", bytes: ab });
const tOpen = performance.now() - t0;
if (!info.ok) {
  console.error(info.error);
  process.exit(1);
}
const result: {
  info: unknown;
  units: unknown[];
  layouts: Record<string, unknown>;
  media: Record<string, unknown>;
} = { info: info.result, units: [], layouts: {}, media: {} };
const count = (info.result as { format: string; count: number }).format === "docx"
  ? 1
  : (info.result as { format: string; count: number }).count;
const times: number[] = [];
const mediaIds = new Set<string>();
const layoutIds = new Set<string>();
const scan = (value: unknown): void => {
  if (Array.isArray(value)) value.forEach(scan);
  else if (value !== null && typeof value === "object") {
    for (const k in value as Record<string, unknown>) {
      const v = (value as Record<string, unknown>)[k];
      if (k === "mediaId") mediaIds.add(v as string);
      else if (k === "layoutId") layoutIds.add(v as string);
      else scan(v);
    }
  }
};
for (let i = 0; i < count; i++) {
  const t = performance.now();
  const u = await call({ type: "unit", index: i });
  times.push(performance.now() - t);
  // re-read (cache path)
  const t2 = performance.now();
  await call({ type: "unit", index: i });
  times.push(performance.now() - t2);
  if (mode !== "bench") {
    result.units.push(u.ok ? u.result : { error: u.error });
    scan(u.result);
  }
  if ((info.result as { format: string }).format === "xlsx" && mode !== "bench") {
    const r = await call({
      type: "unit",
      index: i,
      range: { start: { row: 1, col: 1 }, end: { row: 20, col: 4 } },
    });
    result.units.push(r.ok ? r.result : { error: r.error });
  }
}
for (const id of layoutIds) {
  const l = await call({ type: "layout", layoutId: id });
  result.layouts[id] = l.ok ? l.result : l.error;
  scan(l.result);
}
for (const id of [...mediaIds].sort()) {
  const m = await call({ type: "media", mediaId: id });
  result.media[id] = m.ok
    ? createHash("sha1").update(new Uint8Array(m.result as ArrayBuffer)).digest("hex")
    : m.error;
}
g.gc?.();
const mem = process.memoryUsage();
console.log(JSON.stringify({
  file,
  open: +tOpen.toFixed(1),
  units: times.map((t) => +t.toFixed(1)),
  heapMB: +(mem.heapUsed / 1e6).toFixed(1),
  rssMB: +(mem.rss / 1e6).toFixed(1),
}));
if (out !== undefined) writeFileSync(out, JSON.stringify(result, null, 1));
// The imported worker module can leave stream handles that keep the Node/Bun
// event loop alive after the dump; the harness is done by this point.
process.exit(0);
