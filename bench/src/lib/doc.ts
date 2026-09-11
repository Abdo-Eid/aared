import type { Format } from "./format";

export type DocOrigin = "sample" | "upload";

export interface DocMeta {
  key: string;
  name: string;
  format: Format;
  sizeBytes: number;
  /** Same-origin sample path or object URL; doubles as the download href. */
  url: string;
  bytes: Uint8Array;
  origin: DocOrigin;
}

export type Renderer = "baseline" | "aared";

export interface RunRecord {
  key: string;
  name: string;
  format: Format;
  renderer: Renderer;
  /** Null only if the open never got far enough to time. */
  ms: number | null;
  sizeBytes: number;
  error?: string;
  at: number;
}

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}
