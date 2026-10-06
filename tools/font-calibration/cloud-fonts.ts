// Office's cloud fonts (Aptos and the rest) are downloaded on demand into a
// per-user cache and never installed as system fonts; files are named by hash.
import { readdirSync } from "node:fs";

export const CLOUD_FONTS = `${process.env.LOCALAPPDATA ?? "C:/Users/abdulhamed/AppData/Local"}/Microsoft/FontCache/4/CloudFonts`;

/** Full name (name ID 4, Windows Unicode) of a TrueType file. */
function fullName(bytes: Uint8Array): string | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tables = view.getUint16(4);
  for (let i = 0; i < tables; i++) {
    const record = 12 + i * 16;
    if (String.fromCharCode(...bytes.subarray(record, record + 4)) !== "name") continue;
    const table = view.getUint32(record + 8);
    const count = view.getUint16(table + 2);
    const strings = table + view.getUint16(table + 4);
    for (let n = 0; n < count; n++) {
      const entry = table + 6 + n * 12;
      if (view.getUint16(entry) !== 3 || view.getUint16(entry + 6) !== 4) continue;
      const length = view.getUint16(entry + 8), offset = strings + view.getUint16(entry + 10);
      let out = "";
      for (let c = 0; c < length; c += 2) out += String.fromCharCode(view.getUint16(offset + c));
      return out;
    }
  }
  return undefined;
}

/** The cached file whose full name is `name` (versions 2.00 and 2.01 in the cache have identical advances). */
export async function cloudFontFile(family: string, name: string): Promise<string> {
  const dir = `${CLOUD_FONTS}/${family}`;
  let files: string[];
  try {
    files = readdirSync(dir).filter((file) => file.endsWith(".ttf"));
  } catch {
    throw new Error(`no ${family} in Office's cloud-font cache (${dir}): open a document using it in Office first`);
  }
  for (const file of files) {
    if (fullName(new Uint8Array(await Bun.file(`${dir}/${file}`).arrayBuffer())) === name) return `${dir}/${file}`;
  }
  throw new Error(`no "${name}" among ${dir}'s ${files.length} files`);
}
