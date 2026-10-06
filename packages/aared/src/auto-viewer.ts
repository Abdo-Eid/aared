/** Auto-detection is kept outside the shared controller and fixed-format entries. */
import { openViewer } from "./embed-viewer.js";
import type { OpenViewerInput } from "./embed-viewer.js";
import { WorkerSession } from "./core/session.js";
import { throwIfAborted } from "./core/errors.js";

export function openAutoViewer(input: Omit<OpenViewerInput, "route">) {
  return openViewer({ ...input, route: async (bytes, options) => {
    const signal = options?.signal;
    const format = options?.type ?? await WorkerSession.detectFormat(input.scope, bytes, {
      ...(signal === undefined ? {} : { signal }),
      ...(options?.worker === undefined ? {} : { worker: options.worker }),
    });
    throwIfAborted(signal);
    if (format === "pptx") return (await import("./pptx/viewer.js")).route;
    if (format === "xlsx") return (await import("./xlsx/viewer.js")).route;
    return (await import("./docx/viewer.js")).route;
  } });
}
