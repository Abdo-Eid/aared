import { describe, expect, test } from "bun:test";
import { adoptClassSheet, classSheetText, classes, cssClass } from "../src/class-sheet.js";

/**
 * The class sheet's bookkeeping, on fake roots (pure: no browser). Where the
 * rules land and what they paint is re-read from real DOM by the XLSX and
 * DOCX visual suites.
 */

class FakeSheet {
  static replaced = 0;
  text = "";
  replaceSync(text: string): void {
    FakeSheet.replaced += 1;
    this.text = text;
  }
}

function fakeDocument(): Document {
  const doc = { nodeType: 9, adoptedStyleSheets: [] as unknown[], defaultView: { CSSStyleSheet: FakeSheet } };
  return doc as unknown as Document;
}

function fakeShadow(doc: Document): ShadowRoot {
  return { nodeType: 11, host: {}, ownerDocument: doc, adoptedStyleSheets: [] as unknown[] } as unknown as ShadowRoot;
}

function nodeIn(root: Document | ShadowRoot, doc: Document): Node {
  let current: Node | Document | ShadowRoot = root;
  const node = { getRootNode: () => current, ownerDocument: doc, move(to: Node | Document | ShadowRoot) { current = to; } };
  return node as unknown as Node;
}

describe("class sheet", () => {
  test("a class is interned by its declarations, and rules keep registration order", () => {
    const base = cssClass("margin:0;padding:0");
    expect(cssClass("  margin:0;padding:0  ")).toBe(base);
    const modifier = cssClass("padding:2px");
    expect(modifier).not.toBe(base);
    const text = classSheetText();
    expect(text).toContain(`.${base}:not(#aared-0#aared-1){margin:0;padding:0}`);
    expect(text.indexOf(`.${base}`)).toBeLessThan(text.indexOf(`.${modifier}`));
    expect(classes(base, false, undefined, modifier, null)).toBe(`${base} ${modifier}`);
  });

  test("renders on one root share one adopted sheet; the last release takes it away", () => {
    const doc = fakeDocument();
    const embedder = new FakeSheet();
    doc.adoptedStyleSheets = [embedder as unknown as CSSStyleSheet];
    const first = adoptClassSheet(nodeIn(doc, doc));
    const second = adoptClassSheet(nodeIn(doc, doc));
    expect(doc.adoptedStyleSheets).toHaveLength(2);
    expect(doc.adoptedStyleSheets[0]).toBe(embedder as unknown as CSSStyleSheet);
    first.release();
    first.release();
    expect(doc.adoptedStyleSheets).toHaveLength(2);
    second.release();
    expect(doc.adoptedStyleSheets).toEqual([embedder as unknown as CSSStyleSheet]);
  });

  test("pseudo-element classes are distinct and a hash collision fails rather than changing another class", () => {
    expect(cssClass('content:"x"', ["::before"])).not.toBe(cssClass('content:"x"', ["::after"]));
    // Two valid declaration blocks with colliding FNV-1a hashes for their
    // encoded keys. A registration-order suffix would not be cross-copy safe.
    const first = cssClass("--probe:229599");
    expect(() => cssClass("--probe:432382")).toThrow("renderer style hash collision");
    expect(cssClass("--probe:229599")).toBe(first);
    expect(classSheetText()).not.toContain("--probe:432382");
  });

  test("the sheet is filled once and refilled only when a class was added since", () => {
    const doc = fakeDocument();
    cssClass("color:red");
    const before = FakeSheet.replaced;
    const scope = adoptClassSheet(nodeIn(doc, doc));
    expect(FakeSheet.replaced).toBe(before + 1);
    adoptClassSheet(nodeIn(doc, doc)).release();
    expect(FakeSheet.replaced).toBe(before + 1);
    const late = cssClass("color:blue");
    scope.follow();
    expect(FakeSheet.replaced).toBe(before + 2);
    expect((doc.adoptedStyleSheets[0] as unknown as FakeSheet).text).toContain(late);
    scope.release();
  });

  test("a render inside a shadow root adopts there, and follows its target when it moves", () => {
    const doc = fakeDocument();
    const shadow = fakeShadow(doc);
    const target = nodeIn(target_detached(), doc) as Node & { move(to: unknown): void };
    function target_detached(): Document {
      return { nodeType: 1 } as unknown as Document;
    }
    const scope = adoptClassSheet(target);
    // Detached: the sheet goes to the owner document for now.
    expect(doc.adoptedStyleSheets).toHaveLength(1);
    target.move(shadow);
    scope.follow();
    expect(shadow.adoptedStyleSheets).toHaveLength(1);
    expect(doc.adoptedStyleSheets).toHaveLength(0);
    // An embedder that replaces the list wholesale gets ours back on the next follow.
    shadow.adoptedStyleSheets = [];
    scope.follow();
    expect(shadow.adoptedStyleSheets).toHaveLength(1);
    scope.release();
    expect(shadow.adoptedStyleSheets).toHaveLength(0);
  });

  test("a document without a window is left alone", () => {
    const doc = { nodeType: 9, adoptedStyleSheets: [], defaultView: null } as unknown as Document;
    const scope = adoptClassSheet(nodeIn(doc, doc));
    expect(doc.adoptedStyleSheets).toHaveLength(0);
    scope.follow();
    scope.release();
  });
});
