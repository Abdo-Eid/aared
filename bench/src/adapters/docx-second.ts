import { abortError, type OpenedDoc } from "./types";
import { mountDocxPreviewLineage } from "./docx";

/**
 * docx-renderer (Recon fork of docx-preview-sync → docx-preview). Same
 * `render(data, body, styles?, options?)` shape as the docx-preview baseline
 * and the same paragraph/class output, so the mount is shared — this column
 * measures what the fork's pagination/overlay work costs.
 */
export async function openDocxRenderer(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { render } = await import("docx-renderer");
  if (signal.aborted) throw abortError();
  return mountDocxPreviewLineage(bytes, host, signal, t0, (bodyHost, styleHost) =>
    render(bytes.slice(), bodyHost, styleHost, {
      className: "bench-docx",
      breakPages: true,
    }),
  );
}
