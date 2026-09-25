/**
 * OPC (Open Packaging Conventions) package reader: the shared substrate
 * of PPTX/XLSX/DOCX — ZIP entries + `[Content_Types].xml` + `.rels`.
 *
 * Internal to `core` (not re-exported from the package entrypoint):
 * format code reaches `core` only through the resolved `UnitModel`
 * (see AGENTS.md), and the worker RPC is what turns this
 * into the public `open()` surface.
 */

import { getEntryData, readEntries } from "./zip.js";
import type { ZipEntry } from "./zip.js";
import { AaredError, throwIfAborted } from "./errors.js";
import { attr, childElements, firstChild, parseXml, textContent } from "./xml.js";
import type { XmlElement, XmlStreamOptions } from "./xml.js";

export type OfficeFormat = "pptx" | "xlsx" | "docx";

const OFFICE_DOCUMENT = "application/vnd.openxmlformats-officedocument";

/**
 * Content types of each format's main part: the one place format identity is
 * decided, so the Worker's sniffing and every format's gate agree. Word and
 * Excel templates open as their documents; a PowerPoint template
 * (`presentationml.template`) carries layouts but no slides, so it is not
 * listed and stays unsupported.
 */
const MAIN_PART_TYPES: Record<OfficeFormat, readonly string[]> = {
  pptx: [`${OFFICE_DOCUMENT}.presentationml.presentation.main+xml`],
  xlsx: [`${OFFICE_DOCUMENT}.spreadsheetml.sheet.main+xml`, `${OFFICE_DOCUMENT}.spreadsheetml.template.main+xml`],
  docx: [`${OFFICE_DOCUMENT}.wordprocessingml.document.main+xml`, `${OFFICE_DOCUMENT}.wordprocessingml.template.main+xml`],
};

const FORMAT_NOUN: Record<OfficeFormat, string> = {
  pptx: "a PPTX presentation",
  xlsx: "an XLSX workbook",
  docx: "a Word document",
};

/** Image types by extension, for parts a package leaves without a declared content type. */
const IMAGE_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
};

/**
 * A media part's content type guessed from its extension alone, for callers
 * that hold no package; `application/octet-stream` when the extension is
 * unknown. `OpcPackage.mediaContentType` prefers the type the package declares.
 */
export function guessMediaContentType(partName: string): string {
  const dot = partName.lastIndexOf(".");
  return (dot === -1 ? undefined : IMAGE_TYPES[partName.slice(dot + 1).toLowerCase()]) ?? "application/octet-stream";
}

/** The format whose main part has this content type, if any. */
export function officeFormatOf(contentType: string | undefined): OfficeFormat | undefined {
  if (contentType === undefined) return undefined;
  for (const format of ["pptx", "xlsx", "docx"] as const) {
    if (MAIN_PART_TYPES[format].includes(contentType)) return format;
  }
  return undefined;
}

export interface OpcRelationship {
  id: string;
  type: string;
  /** As written: may be relative to the rels file's directory. */
  target: string;
  targetMode: "Internal" | "External";
}

/** One decoder for every part: `TextDecoder` is stateless for whole-buffer `decode()`. */
const utf8 = new TextDecoder("utf-8");

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
  /**
   * Parsed `.rels` parts by rels path. Relationship lists are tiny and
   * immutable, while format code asks for the same part's relationships many
   * times (every placeholder, drawing, and hyperlink lookup), so each rels
   * part is inflated and parsed once per package. Only finished reads are
   * stored: a cancelled read never poisons the cache.
   */
  private readonly relsCache = new Map<string, OpcRelationship[]>();

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
    if (ct === undefined) throw new Error("aared/core: missing [Content_Types].xml; not an Office document");
    const bytes = await getEntryData(data, ct, signal);
    throwIfAborted(signal);
    const xml = utf8.decode(bytes);
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

  /**
   * One part's bytes in a buffer the caller owns outright (offset 0, full
   * length), so it can be transferred to another thread without a copy.
   */
  async readPart(name: string, signal?: AbortSignal): Promise<Uint8Array> {
    const bytes = await this.readShared(name, signal);
    return bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength && bytes.buffer !== this.data.buffer
      ? bytes
      : bytes.slice();
  }

  /** Part bytes that may alias the archive: decode or copy, never transfer or mutate. */
  private async readShared(name: string, signal?: AbortSignal): Promise<Uint8Array> {
    throwIfAborted(signal);
    const entry = this.entries.get(normalize(name));
    if (entry === undefined) throw new Error(`aared/core: part "${name}" not found in package`);
    const bytes = await getEntryData(this.data, entry, signal);
    throwIfAborted(signal);
    return bytes;
  }

  async readText(name: string, signal?: AbortSignal): Promise<string> {
    throwIfAborted(signal);
    const bytes = await this.readShared(name, signal);
    const text = utf8.decode(bytes);
    throwIfAborted(signal);
    return text;
  }

  /**
   * One part parsed as XML. Cancellation is checked after the inflate and
   * again after the parse, so an aborted read never hands back a tree.
   */
  async readXml(name: string, signal?: AbortSignal, stream?: XmlStreamOptions): Promise<XmlElement> {
    const root = parseXml(await this.readText(name, signal), stream);
    throwIfAborted(signal);
    return root;
  }

  /** `readXml` for a part a relationship may name without the package carrying it: `undefined` when absent. */
  async readXmlIfPresent(name: string, signal?: AbortSignal, stream?: XmlStreamOptions): Promise<XmlElement | undefined> {
    throwIfAborted(signal);
    return this.hasPart(name) ? this.readXml(name, signal, stream) : undefined;
  }

  /**
   * The main document part of `format`: the package's internal
   * `officeDocument` relationship, provided its content type belongs to that
   * format. Throws a message naming the format otherwise.
   */
  async officeDocument(format: OfficeFormat, signal?: AbortSignal): Promise<string> {
    const relationships = await this.packageRelationships(signal);
    const office = relationships.find((relationship) => relationship.type.endsWith("/officeDocument"));
    if (office === undefined || office.targetMode === "External") {
      throw new Error(`aared/core: ${format.toUpperCase()} has no internal officeDocument relationship`);
    }
    if (officeFormatOf(this.contentTypeOf(office.target)) !== format) {
      throw new Error(`aared/core: package is not ${FORMAT_NOUN[format]}`);
    }
    return office.target;
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

  /**
   * A media part's content type: the declared one, else a guess from its
   * extension, else `application/octet-stream`. Every format resolves its
   * pictures through this, so an undeclared type is guessed the same way.
   */
  mediaContentType(partName: string): string {
    return this.contentTypeOf(partName) ?? guessMediaContentType(partName);
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

  /**
   * The document title from the package's core properties (`dc:title` in the
   * part the `core-properties` relationship names, usually
   * `docProps/core.xml`), trimmed; `undefined` when the file sets none.
   */
  async coreTitle(signal?: AbortSignal): Promise<string | undefined> {
    const core = (await this.packageRelationships(signal))
      .find((relationship) => relationship.targetMode !== "External" && relationship.type.endsWith("/core-properties"));
    if (core === undefined) return undefined;
    const properties = await this.readXmlIfPresent(core.target, signal);
    const title = properties === undefined ? undefined : firstChild(properties, "title");
    const text = title === undefined ? "" : textContent(title).trim();
    return text === "" ? undefined : text;
  }

  /**
   * Shared, read-only relationship list: callers must not mutate it (none
   * do; `find`/`filter`/iteration only).
   */
  private async readRels(path: string, dir: string, signal?: AbortSignal): Promise<OpcRelationship[]> {
    throwIfAborted(signal);
    const cached = this.relsCache.get(path);
    if (cached !== undefined) return cached;
    const entry = this.entries.get(path);
    if (entry === undefined) return [];
    const bytes = await getEntryData(this.data, entry, signal);
    throwIfAborted(signal);
    const xml = utf8.decode(bytes);
    throwIfAborted(signal);
    const relationships = parseRelationships(parseXml(xml));
    for (const r of relationships) {
      if (r.targetMode !== "External") r.target = resolveTarget(dir, r.target);
    }
    throwIfAborted(signal);
    this.relsCache.set(path, relationships);
    return relationships;
  }
}
