/**
 * The ratio between painted pixels and layout pixels for anything under
 * `root`: `getBoundingClientRect()` reports the first, computed styles,
 * `offsetWidth` and the model's own page geometry the second, and a CSS `zoom`
 * on any ancestor scales only the first.
 *
 * aared sets that zoom itself (`fit-width` and `fit-page` put it on the render
 * root), and a consumer's browser zoom reaches the same code as a display scale,
 * so every measurement that compares a rect against a layout-pixel number has to
 * divide by this. Reading a rect straight inflated each block in the DOCX
 * pagination probe: measured in Chrome, `zoom: 1.5` on the container turned this
 * repository's 7-page coverage fixture into 10, and Chromium rounds the two
 * spaces differently per zoom factor, which is why the engines disagreed.
 *
 * `tabs.ts` applies the same rule inline, from its own root's width.
 */
export function paintedToLayoutZoom(root: HTMLElement): number {
  const ruler = document.createElement("div");
  ruler.style.cssText = "position:absolute;left:0;top:0;width:100px;height:100px;visibility:hidden";
  root.append(ruler);
  const painted = ruler.getBoundingClientRect().width;
  ruler.remove();
  return painted > 0 ? painted / 100 : 1;
}