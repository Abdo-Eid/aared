import "aared/element";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Zero-JS embedding: drop the element on the page and point it at a file.
 * Listens for `load` once, then hands back the summary and a teardown.
 */
export function runElementDemo(host: HTMLElement, source: string | File): Promise<DemoOutcome> {
  host.replaceChildren();
  const viewer = document.createElement("aared-viewer");
  viewer.setAttribute("controls", "");
  const done = new Promise<DemoOutcome>((resolve, reject) => {
    viewer.addEventListener(
      "load",
      (event) => {
        const info = (event as CustomEvent).detail as { format: string; count: number };
        resolve({
          summary: { format: info.format, units: info.count },
          dispose: () => viewer.remove(),
        });
      },
      { once: true },
    );
    viewer.addEventListener(
      "error",
      (event) => {
        // "error" is a well-known DOM event name, so `event` narrows to
        // `ErrorEvent` here — go through `unknown` since it doesn't
        // structurally overlap `CustomEvent`, unlike the generic "load".
        const detail = (event as unknown as CustomEvent).detail as { message: string };
        reject(new Error(detail.message));
      },
      { once: true },
    );
  });
  host.append(viewer);
  if (typeof source === "string") viewer.setAttribute("src", source);
  else (viewer as unknown as { file: File }).file = source;
  return done;
}
