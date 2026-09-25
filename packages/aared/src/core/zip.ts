/**
 * Minimal ZIP reader for OOXML packages (OPC).
 *
 * Central directory only, inflate on demand — the whole decompressed
 * package is never held in memory (see docs/ARCHITECTURE.md and the
 * D16 decision). Decompression uses the native
 * `DecompressionStream('deflate-raw')`, so `core` keeps zero
 * dependencies. Works headless (plain Bun) and in Workers — no DOM.
 *
 * Deliberately unsupported, with actionable errors instead of silent
 * misreads: encryption, ZIP64, multi-disk archives, and compression
 * methods other than stored (0) / deflated (8). Real Office files
 * never need those for preview.
 */

import { throwIfAborted } from "./errors.js";

// Kept here for existing importers; the helpers live in `errors.ts` so the
// main-thread bundle never needs the ZIP reader to get them.
export { abortError, throwIfAborted } from "./errors.js";

export type ZipErrorCode = "not-a-zip" | "unsupported" | "corrupt";

export class ZipError extends Error {
  readonly code: ZipErrorCode;
  constructor(code: ZipErrorCode, message: string) {
    super(`aared/core: ${message}`);
    this.name = "ZipError";
    this.code = code;
  }
}

export interface ZipEntry {
  /** Entry name as stored, e.g. `ppt/slides/slide1.xml`. */
  name: string;
  /** 0 = stored, 8 = deflated. Anything else is rejected at read time. */
  method: number;
  compressedSize: number;
  size: number;
  localHeaderOffset: number;
}

const utf8 = new TextDecoder("utf-8");

const EOCD_SIG = 0x0605_4b50;
const CENTRAL_SIG = 0x0201_4b50;
const LOCAL_SIG = 0x0403_4b50;

/**
 * Defensive resource limits (see the API's ZIP/decompression safety
 * contract). Office XML parts are kilobytes to low megabytes; these caps
 * sit orders of magnitude above any legitimate preview input and only
 * trip on corrupt or malicious archives. Failures surface as `ZipError`
 * and are mapped to typed `AaredError`s at the worker boundary.
 */
export const MAX_ZIP_ENTRIES = 100_000;
export const MAX_ZIP_ENTRY_BYTES = 256 * 1024 * 1024;

function u16(view: DataView, offset: number): number {
  return view.getUint16(offset, true);
}

function u32(view: DataView, offset: number): number {
  return view.getUint32(offset, true);
}

/**
 * Locate the End Of Central Directory record by scanning backwards
 * (a trailing archive comment of any length is allowed).
 */
function findEocd(view: DataView): number {
  const minPos = Math.max(0, view.byteLength - 22 - 0xffff);
  for (let pos = view.byteLength - 22; pos >= minPos; pos--) {
    if (u32(view, pos) === EOCD_SIG) return pos;
  }
  throw new ZipError("not-a-zip", "not a ZIP archive (no End Of Central Directory record)");
}

/** Parse the central directory; returns entries in archive order. */
export function readEntries(data: Uint8Array, signal?: AbortSignal): ZipEntry[] {
  throwIfAborted(signal);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data.byteLength < 22) {
    throw new ZipError("not-a-zip", "not a ZIP archive (too short)");
  }
  const eocd = findEocd(view);
  const entriesTotal = u16(view, eocd + 10);
  const centralSize = u32(view, eocd + 12);
  const centralOffset = u32(view, eocd + 16);
  if (entriesTotal === 0xffff || centralSize === 0xffff_ffff || centralOffset === 0xffff_ffff) {
    throw new ZipError("unsupported", "ZIP64 archives are not supported");
  }
  if (centralOffset + centralSize > data.byteLength) {
    throw new ZipError("corrupt", "central directory runs past end of file");
  }
  if (entriesTotal > MAX_ZIP_ENTRIES) {
    throw new ZipError("unsupported", `archive lists ${entriesTotal} entries, above the ${MAX_ZIP_ENTRIES} preview limit`);
  }

  const entries: ZipEntry[] = [];
  let pos = centralOffset;
  for (let i = 0; i < entriesTotal; i++) {
    throwIfAborted(signal);
    if (u32(view, pos) !== CENTRAL_SIG) {
      throw new ZipError("corrupt", `central directory entry ${i} has a bad signature`);
    }
    const flags = u16(view, pos + 8);
    const method = u16(view, pos + 10);
    const compressedSize = u32(view, pos + 20);
    const size = u32(view, pos + 24);
    const nameLen = u16(view, pos + 28);
    const extraLen = u16(view, pos + 30);
    const commentLen = u16(view, pos + 32);
    const diskStart = u16(view, pos + 34);
    const localHeaderOffset = u32(view, pos + 42);
    if (flags & 0x01) {
      throw new ZipError("unsupported", "encrypted ZIP entries are not supported");
    }
    if (diskStart !== 0) {
      throw new ZipError("unsupported", "multi-disk ZIP archives are not supported");
    }
    if (compressedSize === 0xffff_ffff || size === 0xffff_ffff || localHeaderOffset === 0xffff_ffff) {
      throw new ZipError("unsupported", "ZIP64 archives are not supported");
    }
    // Office files always set the UTF-8 flag; decode as UTF-8 regardless.
    const name = utf8.decode(data.subarray(pos + 46, pos + 46 + nameLen));
    entries.push({ name, method, compressedSize, size, localHeaderOffset });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  throwIfAborted(signal);
  return entries;
}

async function inflateRaw(compressed: Uint8Array, declaredSize: number, signal?: AbortSignal): Promise<Uint8Array> {
  throwIfAborted(signal);
  const source = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(compressed);
      controller.close();
    },
  });
  // Cast: TS's DOM lib types DecompressionStream's writable as
  // WritableStream<BufferSource>, which its own generics then reject
  // here. At runtime any Uint8Array chunk is valid input (verified in
  // Bun and per the web spec), so this is a typing workaround, not a
  // behavior change.
  const ds = new DecompressionStream("deflate-raw") as unknown as ReadableWritablePair<
    Uint8Array,
    Uint8Array
  >;
  // Count output bytes while streaming so a decompression bomb trips the
  // cap before its full output is ever materialized — `arrayBuffer()` on
  // the whole stream would hold the bomb in memory first.
  //
  // The central directory's declared size is almost always right, so write
  // straight into one buffer of that size: no chunk list, no final concat,
  // and half the peak memory for a large part. A lying header (larger
  // output than declared) falls back to collecting chunks.
  const reader = source.pipeThrough(ds).getReader();
  // Deflate cannot expand past ~1032:1, so a header claiming more than that
  // never gets its claim allocated up front.
  let out = new Uint8Array(Math.min(declaredSize, compressed.byteLength * 1032 + 1024));
  let chunks: Uint8Array[] | undefined;
  let total = 0;
  for (;;) {
    const chunk = await reader.read();
    if (chunk.done) break;
    const value = chunk.value;
    const next = total + value.byteLength;
    if (next > MAX_ZIP_ENTRY_BYTES) {
      void reader.cancel();
      throw new ZipError(
        "unsupported",
        `entry decompresses past the ${MAX_ZIP_ENTRY_BYTES}-byte preview limit`,
      );
    }
    if (chunks === undefined && next <= out.byteLength) {
      out.set(value, total);
    } else {
      chunks ??= [out.subarray(0, total)];
      chunks.push(value);
    }
    total = next;
  }
  if (chunks !== undefined) {
    out = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.byteLength;
    }
  } else if (total !== out.byteLength) {
    out = out.slice(0, total);
  }
  // DecompressionStream cannot be synchronously interrupted. Stop before
  // handing its completed bytes to XML or any later parser stage.
  throwIfAborted(signal);
  return out;
}

/**
 * Read one entry's (decompressed) bytes. The local file header is
 * consulted only to locate the data — sizes come from the central
 * directory, so data-descriptor entries work too.
 *
 * Stored (method 0) entries come back as a view into `data`, not a copy:
 * callers that decode immediately pay nothing, and callers that hand the
 * bytes on (media) copy explicitly — see `OpcPackage.readPart`.
 */
export async function getEntryData(data: Uint8Array, entry: ZipEntry, signal?: AbortSignal): Promise<Uint8Array> {
  throwIfAborted(signal);
  if (entry.size > MAX_ZIP_ENTRY_BYTES) {
    throw new ZipError(
      "unsupported",
      `entry "${entry.name}" declares ${entry.size} bytes, above the ${MAX_ZIP_ENTRY_BYTES}-byte preview limit`,
    );
  }
  if (entry.method !== 0 && entry.method !== 8) {
    throw new ZipError(
      "unsupported",
      `entry "${entry.name}" uses compression method ${entry.method}; only stored (0) and deflated (8) are supported`,
    );
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const lh = entry.localHeaderOffset;
  if (lh + 30 > data.byteLength || u32(view, lh) !== LOCAL_SIG) {
    throw new ZipError("corrupt", `entry "${entry.name}" has a bad local header`);
  }
  const nameLen = u16(view, lh + 26);
  const extraLen = u16(view, lh + 28);
  const dataStart = lh + 30 + nameLen + extraLen;
  if (dataStart + entry.compressedSize > data.byteLength) {
    throw new ZipError("corrupt", `entry "${entry.name}" runs past end of file`);
  }
  const compressed = data.subarray(dataStart, dataStart + entry.compressedSize);
  if (entry.method === 0) return compressed;
  return inflateRaw(compressed, entry.size, signal);
}
