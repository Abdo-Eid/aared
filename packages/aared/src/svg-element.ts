/**
 * One SVG element factory for renderer-side painters. It lives apart from
 * `fill-paint.ts` so the effect painter (and WordArt outlines) can build SVG
 * without fetching the fill painter: a document with effects but no gradient,
 * pattern or picture fill never downloads fill code.
 */

const SVG_NS = "http://www.w3.org/2000/svg";

function numberText(value: number): string {
  return `${Math.round(value * 1_000_000) / 1_000_000}`;
}

export function svgElement<K extends keyof SVGElementTagNameMap>(name: K, attributes: Record<string, string | number> = {}): SVGElementTagNameMap[K] {
  const element = document.createElementNS(SVG_NS, name);
  for (const [key, value] of Object.entries(attributes)) {
    element.setAttribute(key, typeof value === "number" ? numberText(value) : value);
  }
  return element;
}
