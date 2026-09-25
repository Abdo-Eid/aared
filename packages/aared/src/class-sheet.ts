/**
 * Renderer defaults as classes in one adopted stylesheet per render root.
 *
 * A default every cell, paragraph or page shares (a grid cell's padding and
 * gridlines, a paragraph's `margin: 0`) used to be written inline on every
 * element: one declaration block, one style attribute and one
 * CSSStyleDeclaration per element — on a spreadsheet grid or a 121-page
 * document that is most of the page's memory. Now each such default is one
 * rule, interned by `cssClass()` at module scope, and an element carries a
 * class name. Only what varies per element stays inline, and inline still
 * overrides a class default without extra work.
 *
 * The rules live in one constructed `CSSStyleSheet` per root (the document,
 * or the shadow root a render lives in), counted by how many renders use it:
 * adopted before a render attaches its DOM, filled with one `replaceSync`
 * when the registry has grown (never `insertRule` per page), removed when
 * the last render on that root is destroyed. No `<style>` text and no style
 * attribute is ever written, so a strict `style-src` needs no
 * `'unsafe-inline'`.
 *
 * Each rule is `.aared-<hash>:not(#aared-0#aared-1)` — specificity (2,1,0) —
 * so an embedder's own CSS only beats a default with `!important` or a
 * selector with at least two IDs plus a class, close to how inline styles
 * behaved. The class names are internal, not API.
 */

/** Matches every element; adds two IDs' worth of specificity to a rule. */
const BUMP = ":not(#aared-0#aared-1)";

const registry = new Map<string, string>();
let rules = "";
let version = 0;

/** FNV-1a: stable across aared copies on one page, so identical rules share a name. */
function hash(text: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i += 1) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(36);
}

/**
 * The class for a block of constant declarations. Call it at module scope
 * with fixed strings only — a per-render value (a width, a colour) would grow
 * the registry without bound; those stay inline or go through the
 * `--aared-*` tokens. Rules are emitted in registration order, so a modifier
 * declared after its base wins at equal specificity.
 */
export function cssClass(declarations: string): string {
  const key = declarations.trim();
  let name = registry.get(key);
  if (name === undefined) {
    name = `aared-${hash(key)}`;
    registry.set(key, name);
    rules += `.${name}${BUMP}{${key}}\n`;
    version += 1;
  }
  return name;
}

/** Class names joined for `setAttribute("class", …)`; falsy entries are skipped. */
export function classes(...names: Array<string | false | null | undefined>): string {
  return names.filter((name): name is string => typeof name === "string" && name.length > 0).join(" ");
}

type Root = Document | ShadowRoot;

interface Entry {
  sheet: CSSStyleSheet;
  version: number;
  uses: number;
}

const entries = new WeakMap<Root, Entry>();

function rootOf(node: Node): Root | null {
  const root = node.getRootNode();
  if (root.nodeType === 9) return root as Document;
  if (root.nodeType === 11 && "host" in root) return root as ShadowRoot;
  // Detached: the node's document for now; `follow()` moves the scope once it is attached.
  return node.ownerDocument;
}

const documentOf = (root: Root): Document => (root.nodeType === 9 ? (root as Document) : (root as ShadowRoot).ownerDocument);

function ensureAdopted(root: Root, entry: Entry): void {
  if (entry.version !== version) {
    entry.sheet.replaceSync(rules);
    entry.version = version;
  }
  if (!root.adoptedStyleSheets.includes(entry.sheet)) root.adoptedStyleSheets = [...root.adoptedStyleSheets, entry.sheet];
}

function acquire(root: Root): boolean {
  let entry = entries.get(root);
  if (entry === undefined) {
    // The root realm's own constructor, so a render into an iframe works.
    const view = documentOf(root).defaultView as (Window & typeof globalThis) | null;
    if (view === null || typeof view.CSSStyleSheet !== "function") return false;
    entry = { sheet: new view.CSSStyleSheet(), version: -1, uses: 0 };
    entries.set(root, entry);
  }
  entry.uses += 1;
  ensureAdopted(root, entry);
  return true;
}

function release(root: Root): void {
  const entry = entries.get(root);
  if (entry === undefined || entry.uses === 0) return;
  entry.uses -= 1;
  if (entry.uses === 0) root.adoptedStyleSheets = root.adoptedStyleSheets.filter((sheet) => sheet !== entry.sheet);
}

/** One render's hold on its root's class sheet. */
export interface ClassSheetScope {
  /**
   * Re-check where the target lives (it may have been attached or moved) and
   * that the sheet is still adopted. Cheap; renderers call it whenever they
   * paint, measure or re-zoom, so a move or a replaced `adoptedStyleSheets`
   * is repaired on the next of those, not instantly.
   */
  follow(): void;
  /** Let go; the last render on a root takes the sheet away. Idempotent. */
  release(): void;
}

/** Adopt the class sheet on the root `target` lives in. Call before attaching the render's DOM. */
export function adoptClassSheet(target: Node): ClassSheetScope {
  let root = rootOf(target);
  let held = root !== null && acquire(root);
  let done = false;
  return {
    follow() {
      if (done) return;
      const next = rootOf(target);
      if (next !== root) {
        const previous = held ? root : null;
        root = next;
        held = next !== null && acquire(next);
        if (previous !== null) release(previous);
        return;
      }
      // An embedder may have replaced `adoptedStyleSheets` wholesale, or a
      // module registered classes since: put the current sheet back.
      const entry = root === null ? undefined : entries.get(root);
      if (held && root !== null && entry !== undefined) ensureAdopted(root, entry);
    },
    release() {
      if (done) return;
      done = true;
      if (held && root !== null) release(root);
    },
  };
}

/** A constant sheet per document (e.g. `<aared-viewer>`'s shadow chrome), parsed once. */
const constantSheets = new WeakMap<Document, Map<string, CSSStyleSheet>>();

export function constantSheet(doc: Document, text: string): CSSStyleSheet | null {
  const view = doc.defaultView as (Window & typeof globalThis) | null;
  if (view === null || typeof view.CSSStyleSheet !== "function") return null;
  let byText = constantSheets.get(doc);
  if (byText === undefined) {
    byText = new Map();
    constantSheets.set(doc, byText);
  }
  let sheet = byText.get(text);
  if (sheet === undefined) {
    sheet = new view.CSSStyleSheet();
    sheet.replaceSync(text);
    byText.set(text, sheet);
  }
  return sheet;
}

/** The rules as text, for tests. */
export function classSheetText(): string {
  return rules;
}
