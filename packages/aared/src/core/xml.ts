/**
 * Minimal XML parser for OOXML parts (`[Content_Types].xml`, `.rels`,
 * DrawingML, SpreadsheetML, WordprocessingML). Headless by design: plain
 * string scanning, no `DOMParser`, so it runs in Bun unit tests and in
 * Workers alike.
 *
 * Scope is deliberately small — elements, attributes, character data,
 * CDATA, comments, processing instructions — which is everything real
 * Office parts contain. No DTDs, no entities beyond the five predefined
 * plus numeric references.
 *
 * Performance notes (this is the hottest code in the worker):
 * - scanning works on char codes, never per-character regexes or strings;
 * - entity decoding is skipped entirely for runs without `&`;
 * - leaf elements share one frozen `children` array and attribute-less
 *   elements one frozen `attrs` object, so a large part allocates roughly
 *   one object per element instead of three;
 * - element names are interned, so a 100k-element tree holds each distinct
 *   tag name once;
 * - `parseXml(src, { streamChildrenOf, onChild })` hands the direct children
 *   of one element to a callback instead of keeping them, so a worksheet's
 *   `sheetData` (or `sharedStrings`' `si` list) never exists as a whole tree.
 *
 * Parsed trees are read-only: nothing in `core` mutates them, and the shared
 * empty `children`/`attrs` are frozen to keep it that way.
 */

import { damagedDocument } from "./errors.js";

export interface XmlElement {
  /** Qualified name as written, e.g. `a:pPr`. */
  name: string;
  /** Local part after any prefix, e.g. `pPr`. */
  local: string;
  attrs: Record<string, string>;
  children: XmlElement[];
  /** Concatenated direct character data (entities decoded). */
  text: string;
}

export class XmlError extends Error {
  constructor(message: string) {
    super(`aared/core: invalid XML: ${message}`);
    this.name = "XmlError";
  }
}

/** Stream the direct children of every `<streamChildrenOf>` element to `onChild` instead of keeping them. */
export interface XmlStreamOptions {
  /** Local name of the element whose children are streamed (e.g. `sheetData`). */
  streamChildrenOf: string;
  /** Called once per direct child, in document order; the child is then dropped. */
  onChild: (child: XmlElement) => void;
}

const EMPTY_CHILDREN = Object.freeze([]) as unknown as XmlElement[];
const EMPTY_ATTRS = Object.freeze({}) as Record<string, string>;
/** Cap on the module-level name tables (they live as long as the Worker). */
const INTERN_LIMIT = 8192;

/* ------------------------------------------------------------------ */
/* Query helpers                                                      */
/* ------------------------------------------------------------------ */

/**
 * A copy of `s` that shares no storage with the part text it was sliced from.
 * Engines keep a slice of 13 or more characters as a view of its whole parent
 * string, so one attribute value kept from a parsed part keeps the entire
 * decoded part alive. Parsed trees are views of their source on purpose (they
 * are transient); anything kept past a parse — an index, a cache, a
 * module-level table — goes through `ownString` / `ownTree`.
 */
export function ownString(s: string): string {
  return s.length < 13 ? s : (` ${s}`).slice(1);
}

/** A deep copy of a subtree that outlives its part's parse (style and numbering definitions): every string owned. */
export function ownTree(el: XmlElement): XmlElement {
  let attrs = el.attrs;
  if (attrs !== EMPTY_ATTRS) {
    const copy: Record<string, string> = {};
    for (const key in attrs) copy[key] = ownString(attrs[key]!);
    attrs = copy;
  }
  const children = el.children.length === 0 ? EMPTY_CHILDREN : el.children.map(ownTree);
  return { name: ownString(el.name), local: ownString(el.local), attrs, children, text: ownString(el.text) };
}

// Module-level tables outlive every part, so they hold owned copies only.
const localNames = new Map<string, string>();

function localOf(name: string): string {
  let local = localNames.get(name);
  if (local === undefined) {
    const owned = ownString(name);
    const colon = owned.indexOf(":");
    local = colon === -1 ? owned : ownString(owned.slice(colon + 1));
    if (localNames.size < INTERN_LIMIT) localNames.set(owned, local);
  }
  return local;
}

/** Attribute lookup by exact name, falling back to local-name match. */
export function attr(el: XmlElement, name: string): string | undefined {
  const attrs = el.attrs;
  const exact = attrs[name];
  if (exact !== undefined) return exact;
  if (attrs === EMPTY_ATTRS) return undefined;
  const local = localOf(name);
  const n = local.length;
  for (const key in attrs) {
    // `key`'s local part equals `local`: same rule as splitting at its first
    // colon, without allocating. `at` is where that colon must sit.
    const at = key.length - n - 1;
    if (key === local || (at >= 0 && key.indexOf(":") === at && key.endsWith(local))) return attrs[key];
  }
  return undefined;
}

export function childElements(el: XmlElement, local: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const child of el.children) if (child.local === local) out.push(child);
  return out;
}

export function firstChild(el: XmlElement, local: string): XmlElement | undefined {
  for (const child of el.children) if (child.local === local) return child;
  return undefined;
}

/** Recursive text of an element and all descendants. */
export function textContent(el: XmlElement): string {
  const children = el.children;
  if (children.length === 0) return el.text;
  let out = el.text;
  for (const child of children) out += textContent(child);
  return out;
}

/** A finite number from an attribute value in hand; malformed input throws, naming the attribute. */
export function parseNumber(value: string, name: string, context: string): number {
  const number = Number(value);
  if (!Number.isFinite(number)) throw damagedDocument(`invalid ${name}="${value}" on ${context}`);
  return number;
}

/** A numeric attribute: `undefined` when absent, a throw when malformed (see `parseNumber`). */
export function numberAttr(el: XmlElement, name: string, context: string): number | undefined {
  const value = attr(el, name);
  return value === undefined ? undefined : parseNumber(value, name, context);
}

/* ------------------------------------------------------------------ */
/* Parser                                                             */
/* ------------------------------------------------------------------ */

const ENTITY = /&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g;

function decodeEntity(m: string, body: string): string {
  switch (body) {
    case "amp": return "&";
    case "lt": return "<";
    case "gt": return ">";
    case "quot": return '"';
    case "apos": return "'";
    default: {
      const code = body.charCodeAt(1) === 120 /* x */
        ? Number.parseInt(body.slice(2), 16)
        : Number.parseInt(body.slice(1), 10);
      if (!Number.isSafeInteger(code)) throw new XmlError(`bad character reference "${m}"`);
      return String.fromCodePoint(code);
    }
  }
}

function decodeEntities(s: string): string {
  return s.indexOf("&") === -1 ? s : s.replace(ENTITY, decodeEntity);
}

/** `[A-Za-z0-9_.:-]` plus any non-ASCII code unit (XML allows Unicode names). */
function isNameCode(c: number): boolean {
  return (c >= 97 && c <= 122) // a-z
    || (c >= 65 && c <= 90) // A-Z
    || (c >= 48 && c <= 58) // 0-9 and ':'
    || c === 95 || c === 45 || c === 46 // _ - .
    || c >= 0x80;
}

function isSpace(c: number): boolean {
  return c === 32 || c === 10 || c === 9 || c === 13;
}

/** Interned tag names with their local part: every element of one tag shares both strings. */
interface TagName {
  name: string;
  local: string;
}

const internedNames = new Map<string, TagName>();

function intern(name: string): TagName {
  const known = internedNames.get(name);
  if (known !== undefined) return known;
  const owned = ownString(name);
  const tag = { name: owned, local: localOf(owned) };
  if (internedNames.size < INTERN_LIMIT) internedNames.set(owned, tag);
  return tag;
}

class Parser {
  pos = 0;
  readonly len: number;

  constructor(
    readonly src: string,
    private readonly stream: XmlStreamOptions | undefined,
  ) {
    this.len = src.length;
  }

  skipSpace(): void {
    const { src, len } = this;
    let p = this.pos;
    while (p < len && isSpace(src.charCodeAt(p))) p++;
    this.pos = p;
  }

  readName(): string {
    const { src, len } = this;
    const start = this.pos;
    let p = start;
    while (p < len && isNameCode(src.charCodeAt(p))) p++;
    if (p === start) throw new XmlError(`expected a name at offset ${start}`);
    this.pos = p;
    return src.slice(start, p);
  }

  /** Skip `<!-- -->` or `<? ?>` at `pos`; returns false when neither starts here. */
  skipCommentOrPi(): boolean {
    const src = this.src;
    if (src.startsWith("<!--", this.pos)) {
      const end = src.indexOf("-->", this.pos + 4);
      if (end === -1) throw new XmlError("unterminated comment");
      this.pos = end + 3;
      return true;
    }
    if (src.startsWith("<?", this.pos)) {
      const end = src.indexOf("?>", this.pos + 2);
      if (end === -1) throw new XmlError("unterminated processing instruction");
      this.pos = end + 2;
      return true;
    }
    return false;
  }

  /** Skip whitespace, comments, and processing instructions (XML "misc"). */
  skipMisc(): void {
    do this.skipSpace();
    while (this.skipCommentOrPi());
  }

  parseAttributes(): Record<string, string> {
    const src = this.src;
    let attrs: Record<string, string> | undefined;
    for (;;) {
      this.skipSpace();
      const c = src.charCodeAt(this.pos);
      // '/', '>' or end of input
      if (c === 47 || c === 62 || Number.isNaN(c)) return attrs ?? EMPTY_ATTRS;
      const name = this.readName();
      this.skipSpace();
      if (src.charCodeAt(this.pos) !== 61 /* = */) throw new XmlError(`expected "=" after attribute "${name}"`);
      this.pos++;
      this.skipSpace();
      const quote = src.charCodeAt(this.pos);
      if (quote !== 34 && quote !== 39) throw new XmlError(`expected a quoted value for "${name}"`);
      const start = this.pos + 1;
      const end = src.indexOf(quote === 34 ? '"' : "'", start);
      if (end === -1) throw new XmlError(`unterminated value for "${name}"`);
      (attrs ??= {})[name] = decodeEntities(src.slice(start, end));
      this.pos = end + 1;
    }
  }

  parseElement(): XmlElement {
    // Assumes src[pos] === '<' and not a declaration/comment/CDATA/close.
    const src = this.src;
    this.pos++;
    const { name, local } = intern(this.readName());
    const attrs = this.parseAttributes();
    this.skipSpace();
    if (src.charCodeAt(this.pos) === 47 /* / */ && src.charCodeAt(this.pos + 1) === 62 /* > */) {
      this.pos += 2;
      return { name, local, attrs, children: EMPTY_CHILDREN, text: "" };
    }
    if (src.charCodeAt(this.pos) !== 62) throw new XmlError(`expected ">" to open <${name}>`);
    this.pos++;

    const onChild = this.stream !== undefined && this.stream.streamChildrenOf === local ? this.stream.onChild : undefined;
    let children: XmlElement[] | undefined;
    let text = "";
    for (;;) {
      const lt = src.indexOf("<", this.pos);
      if (lt === -1) throw new XmlError(`unterminated <${name}>`);
      // Streamed containers drop their inter-child whitespace with their children.
      if (lt > this.pos && onChild === undefined) text += decodeEntities(src.slice(this.pos, lt));
      this.pos = lt;
      const next = src.charCodeAt(lt + 1);
      if (next === 47 /* / */) {
        const start = lt + 2;
        const end = start + name.length;
        if (!src.startsWith(name, start) || (end < this.len && isNameCode(src.charCodeAt(end)))) {
          this.pos = start;
          throw new XmlError(`mismatched close tag </${this.readName()}> for <${name}>`);
        }
        this.pos = end;
        this.skipSpace();
        if (src.charCodeAt(this.pos) !== 62) throw new XmlError(`expected ">" to close </${name}>`);
        this.pos++;
        return { name, local, attrs, children: children ?? EMPTY_CHILDREN, text };
      }
      if (next === 33 /* ! */) {
        if (src.startsWith("<!--", lt)) {
          this.skipCommentOrPi();
          continue;
        }
        if (src.startsWith("<![CDATA[", lt)) {
          const close = src.indexOf("]]>", lt + 9);
          if (close === -1) throw new XmlError("unterminated CDATA section");
          if (onChild === undefined) text += src.slice(lt + 9, close);
          this.pos = close + 3;
          continue;
        }
        throw new XmlError("DOCTYPE and other declarations are not supported");
      }
      if (next === 63 /* ? */) {
        this.skipCommentOrPi();
        continue;
      }
      const child = this.parseElement();
      if (onChild !== undefined) onChild(child);
      else (children ??= []).push(child);
    }
  }
}

/**
 * Parse a whole XML document; exactly one root element is required. With
 * `stream`, the direct children of every `stream.streamChildrenOf` element
 * go to `stream.onChild` (in document order, synchronously) and that
 * element is returned with no children.
 */
export function parseXml(src: string, stream?: XmlStreamOptions): XmlElement {
  const p = new Parser(src, stream);
  if (src.charCodeAt(0) === 0xfeff) p.pos++;
  p.skipMisc();
  if (src.charCodeAt(p.pos) !== 60 /* < */ || src.charCodeAt(p.pos + 1) === 47) {
    throw new XmlError("no root element");
  }
  const root = p.parseElement();
  p.skipMisc();
  if (p.pos !== src.length) throw new XmlError("trailing content after root element");
  return root;
}
