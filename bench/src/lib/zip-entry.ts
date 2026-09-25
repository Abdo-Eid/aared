/**
 * Read one part of an OOXML package as text — the bench's own side reads
 * (a sheet's RTL flag, a document's widest section), done without any
 * library under measurement. It used to borrow JSZip, which loaded a
 * baseline's own dependency ahead of it and charged the harness's work
 * (and a kept JSZip instance) to that baseline's time, memory and download
 * size. The browser's `DecompressionStream` needs no library at all.
 */

const EOCD = 0x06054b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** The named part's text, or null when the package has no such part or can't be read. */
export async function readZipEntryText(bytes: Uint8Array, name: string): Promise<string | null> {
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65557); at -= 1) {
      if (view.getUint32(at, true) === EOCD) {
        eocd = at;
        break;
      }
    }
    if (eocd < 0) return null;
    const count = view.getUint16(eocd + 10, true);
    let at = view.getUint32(eocd + 16, true);
    const decoder = new TextDecoder();
    for (let index = 0; index < count; index += 1) {
      if (view.getUint32(at, true) !== CENTRAL) return null;
      const method = view.getUint16(at + 10, true);
      const compressed = view.getUint32(at + 20, true);
      const nameLength = view.getUint16(at + 28, true);
      const extraLength = view.getUint16(at + 30, true);
      const commentLength = view.getUint16(at + 32, true);
      const local = view.getUint32(at + 42, true);
      const entry = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
      if (entry === name) {
        if (view.getUint32(local, true) !== LOCAL) return null;
        const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
        const data = bytes.subarray(start, start + compressed);
        if (method === 0) return decoder.decode(data);
        if (method === 8) return decoder.decode(await inflateRaw(data));
        return null;
      }
      at += 46 + nameLength + extraLength + commentLength;
    }
    return null;
  } catch {
    return null;
  }
}
