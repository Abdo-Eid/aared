/**
 * Shared ZIP utilities for tests. Synthetic OPC packages are used to test
 * parser edge cases without needing real Office files. This module is
 * imported by test files but is not itself a test (no `.test.` in the name).
 */

let crc32Table: Uint32Array | undefined;

/**
 * Compute the CRC32 checksum of bytes, used to verify ZIP entry integrity.
 * The table is computed once and cached.
 */
function crc32(bytes: Uint8Array): number {
  let table = crc32Table;
  if (table === undefined) {
    table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c % 2 === 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
    crc32Table = table;
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = table[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Create a minimal stored (uncompressed) ZIP file from an array of
 * [filename, content] tuples. Converts to the record form and returns an
 * ArrayBuffer for compatibility with preview tests.
 */
export function storedZipFromTuples(parts: Array<[name: string, text: string]>): ArrayBuffer {
  const files: Record<string, string> = {};
  for (const [name, text] of parts) files[name] = text;
  return storedZip(files).buffer as ArrayBuffer;
}

/**
 * Create a minimal stored (uncompressed) ZIP file from a record of file names
 * and contents. Used to synthesize OPC packages for parser tests. Returns a
 * `Uint8Array` containing a structurally valid ZIP: local file headers +
 * central directory + end of central directory.
 */
export function storedZip(files: Record<string, Uint8Array | string>): Uint8Array {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const push = (target: Uint8Array[], view: DataView | Uint8Array): void => {
    target.push(view instanceof Uint8Array ? view : new Uint8Array(view.buffer));
  };
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = typeof content === "string" ? encoder.encode(content) : content;
    const header = new DataView(new ArrayBuffer(30));
    header.setUint32(0, 0x04034b50, true);
    header.setUint16(4, 20, true);
    header.setUint16(6, 0x0800, true);
    header.setUint16(8, 0, true);
    header.setUint32(14, crc32(data), true);
    header.setUint32(18, data.byteLength, true);
    header.setUint32(22, data.byteLength, true);
    header.setUint16(26, nameBytes.byteLength, true);
    push(chunks, header);
    push(chunks, nameBytes);
    push(chunks, data);
    const entry = new DataView(new ArrayBuffer(46));
    entry.setUint32(0, 0x02014b50, true);
    entry.setUint16(10, 0, true);
    entry.setUint32(20, data.byteLength, true);
    entry.setUint32(24, data.byteLength, true);
    entry.setUint16(28, nameBytes.byteLength, true);
    entry.setUint32(42, offset, true);
    push(central, entry);
    push(central, nameBytes);
    offset += 30 + nameBytes.byteLength + data.byteLength;
  }
  for (const part of central) chunks.push(part);
  const centralSize = central.reduce((sum, part) => sum + part.byteLength, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, Object.keys(files).length, true);
  end.setUint16(10, Object.keys(files).length, true);
  end.setUint32(12, centralSize, true);
  end.setUint32(16, offset, true);
  push(chunks, end);
  const out = new Uint8Array(offset + centralSize + 22);
  let position = 0;
  for (const part of chunks) {
    out.set(part, position);
    position += part.byteLength;
  }
  return out;
}
