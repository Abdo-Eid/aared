/**
 * Word `styles.xml` reader: paragraph/character style inheritance for the DOCX
 * route. Headless, plain data — mirrors how `core/pptx.ts` resolves
 * DrawingML placeholder inheritance, but over `w:basedOn` chains.
 *
 * Precedence (Word): direct formatting > character style (+ its bases) >
 * paragraph style (+ its bases) > `w:docDefaults`. Duplicate `w:styleId`
 * definitions resolve last-wins; reference `.docx` files in the wild repeat
 * `Heading1`/`Heading2` blocks.
 */

import { attr, childElements, firstChild, ownString, ownTree } from "../xml.js";
import type { XmlElement } from "../xml.js";
import { throwIfAborted } from "../errors.js";

export interface DocxStyle {
  type: "paragraph" | "character";
  basedOn?: string;
  name?: string;
  pPr?: XmlElement;
  rPr?: XmlElement;
}

export interface StyleData {
  defaultPPr?: XmlElement;
  defaultRPr?: XmlElement;
  /**
   * The paragraph style Word marks `w:default="1"` (usually `Normal`): what a
   * paragraph with no `w:pStyle` uses, so its font, size and spacing are this
   * style's, not the bare document defaults'.
   */
  defaultParagraphStyleId?: string;
  paragraph: Map<string, DocxStyle>;
  character: Map<string, DocxStyle>;
}

/** The spellings of an on/off attribute that mean on (`w:default="1"`, `"true"`, `"on"`). */
const ON_VALUES = new Set(["1", "true", "on"]);

function readStyle(style: XmlElement): { id: string; parsed: DocxStyle } | undefined {
  const id = attr(style, "styleId");
  const type = attr(style, "type");
  if (id === undefined) return undefined;
  if (type !== "paragraph" && type !== "character") return undefined;
  const basedOn = firstChild(style, "basedOn");
  const name = firstChild(style, "name");
  const pPr = firstChild(style, "pPr");
  const rPr = firstChild(style, "rPr");
  const basedOnId = basedOn === undefined ? undefined : attr(basedOn, "val");
  const displayName = name === undefined ? undefined : attr(name, "val");
  // The session keeps styles for its lifetime: own every string, so they don't pin styles.xml.
  return {
    id: ownString(id),
    parsed: {
      type,
      ...(basedOnId === undefined ? {} : { basedOn: ownString(basedOnId) }),
      ...(displayName === undefined ? {} : { name: ownString(displayName) }),
      ...(pPr === undefined ? {} : { pPr: ownTree(pPr) }),
      ...(rPr === undefined ? {} : { rPr: ownTree(rPr) }),
    },
  };
}

export function parseStyles(root: XmlElement | undefined, signal?: AbortSignal): StyleData {
  throwIfAborted(signal);
  const data: StyleData = { paragraph: new Map(), character: new Map() };
  if (root === undefined) return data;
  const defaults = firstChild(root, "docDefaults");
  const rPrDefault = defaults === undefined ? undefined : firstChild(defaults, "rPrDefault");
  const pPrDefault = defaults === undefined ? undefined : firstChild(defaults, "pPrDefault");
  const defaultRPr = rPrDefault === undefined ? undefined : firstChild(rPrDefault, "rPr");
  const defaultPPr = pPrDefault === undefined ? undefined : firstChild(pPrDefault, "pPr");
  if (defaultRPr !== undefined) data.defaultRPr = ownTree(defaultRPr);
  if (defaultPPr !== undefined) data.defaultPPr = ownTree(defaultPPr);
  for (const style of childElements(root, "style")) {
    throwIfAborted(signal);
    const parsed = readStyle(style);
    if (parsed === undefined) continue;
    // Last definition wins: real files repeat Heading1/Heading2 blocks.
    if (parsed.parsed.type === "paragraph") {
      data.paragraph.set(parsed.id, parsed.parsed);
      if (ON_VALUES.has(attr(style, "default") ?? "")) data.defaultParagraphStyleId = parsed.id;
    } else {
      data.character.set(parsed.id, parsed.parsed);
    }
  }
  throwIfAborted(signal);
  return data;
}

/**
 * Style ancestry base-first (so callers apply direct formatting last).
 * Cycle-safe: a repeated `basedOn` target ends the chain instead of looping.
 */
function ancestry(styles: ReadonlyMap<string, DocxStyle>, styleId: string | undefined): DocxStyle[] {
  const chain: DocxStyle[] = [];
  const seen = new Set<string>();
  let current = styleId;
  while (current !== undefined && !seen.has(current)) {
    seen.add(current);
    const style = styles.get(current);
    if (style === undefined) break;
    chain.unshift(style);
    current = style.basedOn;
  }
  return chain;
}

export function paragraphAncestry(styles: StyleData, styleId: string | undefined): DocxStyle[] {
  return ancestry(styles.paragraph, styleId);
}

export function characterAncestry(styles: StyleData, styleId: string | undefined): DocxStyle[] {
  return ancestry(styles.character, styleId);
}

const HEADING_STYLE_ID = /^Heading([1-9])$/;
const HEADING_STYLE_NAME = /^heading ([1-9])$/i;

/** Heading level 1–9 from a paragraph style id or display name, if it names one. */
export function headingLevelForStyle(styleId: string | undefined, styles: StyleData): number | undefined {
  if (styleId === undefined) return undefined;
  const idMatch = HEADING_STYLE_ID.exec(styleId);
  if (idMatch !== null) return Number(idMatch[1]);
  const name = styles.paragraph.get(styleId)?.name;
  if (name === undefined) return undefined;
  const nameMatch = HEADING_STYLE_NAME.exec(name.trim());
  return nameMatch === null ? undefined : Number(nameMatch[1]);
}
