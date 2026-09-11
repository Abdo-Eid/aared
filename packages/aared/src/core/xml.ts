/**
 * Minimal XML parser for OOXML parts (`[Content_Types].xml`, `.rels`,
 * and later DrawingML). Headless by design: plain string scanning, no
 * `DOMParser`, so it runs in Bun unit tests and in Workers alike.
 *
 * Scope is deliberately small — elements, attributes, character data,
 * CDATA, comments, processing instructions — which is everything real
 * Office parts contain. No DTDs, no entities beyond the five
 * predefined plus numeric references (anything else is a corrupt-file
 * error, not a silent pass-through).
 */

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
    super(`aared/core: invalid XML — ${message}`);
    this.name = "XmlError";
  }
}

/** Attribute lookup by exact name, falling back to local-name match. */
export function attr(el: XmlElement, name: string): string | undefined {
  const exact = el.attrs[name];
  if (exact !== undefined) return exact;
  const local = name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
  for (const key of Object.keys(el.attrs)) {
    const keyLocal = key.includes(":") ? key.slice(key.indexOf(":") + 1) : key;
    if (keyLocal === local) return el.attrs[key];
  }
  return undefined;
}

export function childElements(el: XmlElement, local: string): XmlElement[] {
  return el.children.filter((c) => c.local === local);
}

export function firstChild(el: XmlElement, local: string): XmlElement | undefined {
  return el.children.find((c) => c.local === local);
}

/** Recursive text of an element and all descendants. */
export function textContent(el: XmlElement): string {
  let out = el.text;
  for (const c of el.children) out += textContent(c);
  return out;
}

function decodeEntities(s: string): string {
  return s.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, body: string) => {
    switch (body) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default: {
        const code = body.startsWith("#x") || body.startsWith("#X")
          ? Number.parseInt(body.slice(2), 16)
          : Number.parseInt(body.slice(1), 10);
        if (!Number.isSafeInteger(code)) throw new XmlError(`bad character reference "${m}"`);
        return String.fromCodePoint(code);
      }
    }
  });
}

function isNameChar(ch: string): boolean {
  return /[A-Za-z0-9_.:-]/.test(ch);
}

interface Cursor {
  src: string;
  pos: number;
}

function skipWhitespace(c: Cursor): void {
  while (c.pos < c.src.length && /\s/.test(c.src[c.pos] ?? "")) c.pos++;
}

function readName(c: Cursor): string {
  const start = c.pos;
  while (c.pos < c.src.length && isNameChar(c.src[c.pos] ?? "")) c.pos++;
  if (c.pos === start) throw new XmlError(`expected a name at offset ${c.pos}`);
  return c.src.slice(start, c.pos);
}

function parseAttributes(c: Cursor): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (;;) {
    skipWhitespace(c);
    const ch = c.src[c.pos];
    if (ch === "/" || ch === ">" || ch === undefined) return attrs;
    const name = readName(c);
    skipWhitespace(c);
    if (c.src[c.pos] !== "=") throw new XmlError(`expected "=" after attribute "${name}"`);
    c.pos++;
    skipWhitespace(c);
    const quote = c.src[c.pos];
    if (quote !== '"' && quote !== "'") throw new XmlError(`expected a quoted value for "${name}"`);
    c.pos++;
    const end = c.src.indexOf(quote, c.pos);
    if (end === -1) throw new XmlError(`unterminated value for "${name}"`);
    attrs[name] = decodeEntities(c.src.slice(c.pos, end));
    c.pos = end + 1;
  }
}

/** Skip whitespace, comments, and processing instructions (XML "misc"). */
function skipMisc(c: Cursor): void {
  for (;;) {
    skipWhitespace(c);
    if (c.src.startsWith("<!--", c.pos)) {
      const end = c.src.indexOf("-->", c.pos + 4);
      if (end === -1) throw new XmlError("unterminated comment");
      c.pos = end + 3;
    } else if (c.src.startsWith("<?", c.pos)) {
      const end = c.src.indexOf("?>", c.pos + 2);
      if (end === -1) throw new XmlError("unterminated processing instruction");
      c.pos = end + 2;
    } else {
      return;
    }
  }
}

function splitName(name: string): { local: string } {
  const i = name.indexOf(":");
  return { local: i === -1 ? name : name.slice(i + 1) };
}

function parseElement(c: Cursor): XmlElement {
  // Assumes src[pos] === '<' and not a declaration/comment/CDATA/close.
  c.pos++;
  const name = readName(c);
  const attrs = parseAttributes(c);
  skipWhitespace(c);
  if (c.src.startsWith("/>", c.pos)) {
    c.pos += 2;
    return { name, ...splitName(name), attrs, children: [], text: "" };
  }
  if (c.src[c.pos] !== ">") throw new XmlError(`expected ">" to open <${name}>`);
  c.pos++;

  const children: XmlElement[] = [];
  let text = "";
  for (;;) {
    const lt = c.src.indexOf("<", c.pos);
    if (lt === -1) throw new XmlError(`unterminated <${name}>`);
    text += decodeEntities(c.src.slice(c.pos, lt));
    c.pos = lt;
    if (c.src.startsWith("</", c.pos)) {
      c.pos += 2;
      const close = readName(c);
      if (close !== name) throw new XmlError(`mismatched close tag </${close}> for <${name}>`);
      skipWhitespace(c);
      if (c.src[c.pos] !== ">") throw new XmlError(`expected ">" to close </${close}>`);
      c.pos++;
      return { name, ...splitName(name), attrs, children, text };
    }
    if (c.src.startsWith("<!--", c.pos)) {
      const end = c.src.indexOf("-->", c.pos + 4);
      if (end === -1) throw new XmlError("unterminated comment");
      c.pos = end + 3;
      continue;
    }
    if (c.src.startsWith("<![CDATA[", c.pos)) {
      const end = c.src.indexOf("]]>", c.pos + 9);
      if (end === -1) throw new XmlError("unterminated CDATA section");
      text += c.src.slice(c.pos + 9, end);
      c.pos = end + 3;
      continue;
    }
    if (c.src.startsWith("<?", c.pos)) {
      const end = c.src.indexOf("?>", c.pos + 2);
      if (end === -1) throw new XmlError("unterminated processing instruction");
      c.pos = end + 2;
      continue;
    }
    if (c.src.startsWith("<!", c.pos)) {
      throw new XmlError("DOCTYPE and other declarations are not supported");
    }
    children.push(parseElement(c));
  }
}

/** Parse a whole XML document; exactly one root element is required. */
export function parseXml(src: string): XmlElement {
  const c: Cursor = { src, pos: 0 };
  if (src.charCodeAt(0) === 0xfeff) c.pos++;
  skipMisc(c);
  if (!c.src.startsWith("<", c.pos) || c.src.startsWith("</", c.pos)) {
    throw new XmlError("no root element");
  }
  const root = parseElement(c);
  skipMisc(c);
  if (c.pos !== c.src.length) throw new XmlError("trailing content after root element");
  return root;
}
