/** Auto-detection is kept outside the shared controller and fixed-format entries. */
import type { OpenViewerInput } from "./embed-viewer.js";

export async function openAutoViewer(input: Omit<OpenViewerInput, "route">) {
  // Keep every entry's controller edge lazy. In a multi-page consumer an
  // eager React/element edge lets Rollup fold the controller into that page's
  // bootstrap, which a fixed-format preview must never execute.
  const { openViewer } = await import("./embed-viewer.js");
  return openViewer({ ...input, route: async (format) => {
    if (format === "pptx") return (await import("./pptx/viewer.js")).route;
    if (format === "xlsx") return (await import("./xlsx/viewer.js")).route;
    return (await import("./docx/viewer.js")).route;
  } });
}
