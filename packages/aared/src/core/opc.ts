/**
 * OPC (Open Packaging Conventions) package reader: the shared substrate
 * of PPTX/XLSX/DOCX — ZIP entries + `[Content_Types].xml` + `.rels`.
 *
 * Internal to `core` (not re-exported from the package entrypoint):
 * format code reaches `core` only through the resolved `UnitModel`
 * (see AGENTS.md), and the worker RPC is what turns this
 * into the public `open()` surface.
 */

import { getEntryData, readEntries, throwIfAborted } from "./zip.js";
import type { ZipEntry } from "./zip.js";
import { AaredError } from "./errors.js";
import { attr, childElements, parseXml } from "./xml.js";
import type { XmlElement } from "./xml.js";

export interface OpcRelationship {
  id: string;
  type: string;
  /** As written: may be relative to the rels file's directory. */
  target: string;
  targetMode: "Internal" | "External";
}

function normalize(name: string): string {
  return name.startsWith("/") ? name.slice(1) : name;
}

function parseContentTypes(root: XmlElement): { defaults: Map<string, string>; overrides: Map<string, string> } {
  if (root.local !== "Types") throw new Error('aared/core: [Content_Types].xml root must be <Types>');
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const child of root.children) {
    if (child.local === "Default") {
      const ext = attr(child, "Extension");
      const type = attr(child, "ContentType");
      if (ext !== undefined && type !== undefined) defaults.set(ext.toLowerCase(), type);
    } else if (child.local === "Override") {
      const part = attr(child, "PartName");
      const type = attr(child, "ContentType");
      if (part !== undefined && type !== undefined) overrides.set(normalize(part), type);
    }
  }
  return { defaults, overrides };
}

function parseRelationships(root: XmlElement): OpcRelationship[] {
  if (root.local !== "Relationships") throw new Error("aared/core: rels root must be <Relationships>");
  const out: OpcRelationship[] = [];
  for (const child of childElements(root, "Relationship")) {
    const id = attr(child, "Id");
    const type = attr(child, "Type");
    const target = attr(child, "Target");
    if (id === undefined || type === undefined || target === undefined) continue;
    const mode = attr(child, "TargetMode");
    out.push({ id, type, target, targetMode: mode === "External" ? "External" : "Internal" });
  }
  return out;
}

/** Resolve a relative rels target against the rels file's directory. */
export function resolveTarget(relsDir: string, target: string): string {
  if (target.startsWith("/")) return normalize(target);
  const parts = relsDir === "" ? [] : relsDir.split("/");
  for (const seg of target.split("/")) {
    if (seg === "." || seg === "") continue;
    if (seg === "..") parts.pop();
    else parts.push(seg);
  }
  return parts.join("/");
}

function relsPathFor(partName: string): string {
  const slash = partName.lastIndexOf("/");
  const dir = slash === -1 ? "_rels" : `${partName.slice(0, slash)}/_rels`;
  const base = slash === -1 ? partName : partName.slice(slash + 1);
  return `${dir}/${base}.rels`;
}

export class OpcPackage {
  private constructor(
    private readonly data: Uint8Array,
    private readonly entries: Map<string, ZipEntry>,
    private readonly defaults: Map<string, string>,
    private readonly overrides: Map<string, string>,
  ) {}

  static async open(source: Uint8Array | ArrayBuffer, signal?: AbortSignal): Promise<OpcPackage> {
    throwIfAborted(signal);
    const data = source instanceof Uint8Array ? source : new Uint8Array(source);
    // Password-protected Office files are OLE compound documents, not ZIP
    // archives. Name that case instead of reporting a confusing ZIP error.
    if (
      data.byteLength >= 8
      && data[0] === 0xd0 && data[1] === 0xcf && data[2] === 0x11 && data[3] === 0xe0
      && data[4] === 0xa1 && data[5] === 0xb1 && data[6] === 0x1a && data[7] === 0xe1
    ) {
      throw new AaredError(
        "encrypted-document",
        "aared/core: this file is password-protected (OLE compound document); encrypted documents cannot be previewed",
      );
    }
    const map = new Map<string, ZipEntry>();
    for (const e of readEntries(data, signal)) {
      throwIfAborted(signal);
      map.set(e.name, e);
    }
    const ct = map.get("[Content_Types].xml");
    if (ct === undefined) throw new Error("aared/core: missing [Content_Types].xml — not an Office document");
    const bytes = await getEntryData(data, ct, signal);
    throwIfAborted(signal);
    const xml = new TextDecoder("utf-8").decode(bytes);
    throwIfAborted(signal);
    const { defaults, overrides } = parseContentTypes(parseXml(xml));
    throwIfAborted(signal);
    return new OpcPackage(data, map, defaults, overrides);
  }

  /** All part names in archive order. */
  partNames(): string[] {
    return [...this.entries.keys()];
  }

  hasPart(name: string): boolean {
    return this.entries.has(normalize(name));
  }

  async readPart(name: string, signal?: AbortSignal): Promise<Uint8Array> {
    throwIfAborted(signal);
    const entry = this.entries.get(normalize(name));
    if (entry === undefined) throw new Error(`aared/core: part "${name}" not found in package`);
    const bytes = await getEntryData(this.data, entry, signal);
    throwIfAborted(signal);
    return bytes;
  }

  async readText(name: string, signal?: AbortSignal): Promise<string> {
    throwIfAborted(signal);
    const bytes = await this.readPart(name, signal);
    throwIfAborted(signal);
    const text = new TextDecoder("utf-8").decode(bytes);
    throwIfAborted(signal);
    return text;
  }

  /** Content type via Override match, else Default by extension. */
  contentTypeOf(partName: string): string | undefined {
    const name = normalize(partName);
    const override = this.overrides.get(name);
    if (override !== undefined) return override;
    const dot = name.lastIndexOf(".");
    if (dot === -1) return undefined;
    return this.defaults.get(name.slice(dot + 1).toLowerCase());
  }

  /** Relationships of one part; parts without a rels file yield []. */
  async relationships(partName: string, signal?: AbortSignal): Promise<OpcRelationship[]> {
    throwIfAborted(signal);
    const name = normalize(partName);
    const slash = name.lastIndexOf("/");
    const dir = slash === -1 ? "" : name.slice(0, slash);
    return this.readRels(relsPathFor(name), dir, signal);
  }

  /** Package-level relationships (`_rels/.rels`). */
  async packageRelationships(signal?: AbortSignal): Promise<OpcRelationship[]> {
    throwIfAborted(signal);
    return this.readRels("_rels/.rels", "", signal);
  }

  private async readRels(path: string, dir: string, signal?: AbortSignal): Promise<OpcRelationship[]> {
    throwIfAborted(signal);
    const entry = this.entries.get(path);
    if (entry === undefined) return [];
    const bytes = await getEntryData(this.data, entry, signal);
    throwIfAborted(signal);
    const xml = new TextDecoder("utf-8").decode(bytes);
    throwIfAborted(signal);
    const relationships = parseRelationships(parseXml(xml)).map((r) => ({
      ...r,
      target: r.targetMode === "External" ? r.target : resolveTarget(dir, r.target),
    }));
    throwIfAborted(signal);
    return relationships;
  }
}
