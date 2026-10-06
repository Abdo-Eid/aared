import "aared/element";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * One-tag embedding: import the element once, then drop it on the page and point it at a file.
 * Listens for `load` once, then hands back the summary and a teardown.
 */
export function runElementDemo(host: HTMLElement, source: string | File): Promise<DemoOutcome> {
  host.replaceChildren();
  // Typed by `aared/element`: an `AaredViewer`, with each event's `detail` typed.
  const viewer = document.createElement("aared-viewer");
  viewer.controls = true;
  const done = new Promise<DemoOutcome>((resolve, reject) => {
    viewer.addEventListener(
      "load",
      (event) => {
        resolve({
          summary: { format: event.detail.format, units: event.detail.count },
          dispose: () => viewer.remove(),
        });
      },
      { once: true },
    );
    viewer.addEventListener("error", (event) => reject(event.detail.cause), { once: true });
  });
  host.append(viewer);
  if (typeof source === "string") viewer.src = source;
  else viewer.file = source;
  return done;
}
