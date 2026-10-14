/**
 * One browser bundle for the tab-stop tests: the renderer and the tab pass
 * must share a module instance, because a paragraph's stops live in a
 * module-level WeakMap that `render()` fills and the pass reads.
 */
export { render } from "../src/docx/renderer.js";
export { layoutTabParagraphs } from "../src/docx/tabs.js";
