import type { WorkerInit } from "./index.js";
import { AaredError } from "./errors.js";

/** Normalized, validated before construction so invalid input is not a Worker failure. */
export type WorkerInput =
  | { kind: "default" }
  | { kind: "url"; url: string; bootstrap: boolean }
  | { kind: "source"; source: string }
  | { kind: "create"; create: () => Worker };

export function workerInput(scope: string, init: WorkerInit | undefined): WorkerInput {
  if (init === undefined) return { kind: "default" };
  const fail = (): never => { throw new AaredError("invalid-argument",
    `${scope}: worker requires exactly one of { url }, { source }, or { create }; only url accepts crossOrigin: "blob"`); };
  if (typeof init !== "object" || init === null) return fail();
  const input = init as Record<string, unknown>;
  const keys = ["url", "source", "create"].filter((key) => input[key] !== undefined);
  if (keys.length !== 1 || Object.keys(input).some((key) => !["url", "source", "create", "crossOrigin"].includes(key))) return fail();
  if (input.url !== undefined) {
    if (!(typeof input.url === "string" && input.url.trim() !== "" || input.url instanceof URL)) return fail();
    if (input.crossOrigin !== undefined && input.crossOrigin !== "blob") return fail();
    const raw = String(input.url);
    // Headless imports remain safe; relative URLs resolve only when opened in a page.
    const base = (globalThis as { location?: { href: string } }).location?.href;
    let url = raw;
    try { if (base !== undefined) url = new URL(raw, base).href; } catch { return fail(); }
    return { kind: "url", url, bootstrap: input.crossOrigin === "blob" };
  }
  if (input.crossOrigin !== undefined) return fail();
  if (input.source !== undefined) {
    if (typeof input.source !== "string" || input.source.trim() === "") return fail();
    return { kind: "source", source: input.source };
  }
  if (typeof input.create !== "function") return fail();
  return { kind: "create", create: input.create as () => Worker };
}
