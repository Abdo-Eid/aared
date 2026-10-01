import { abortError, type OpenedDoc } from "./types";
import { assertZip, mountDocxPreviewLineage } from "./docx";

/**
 * @apollo-design/docx-preview — a reworked docx-preview fork whose
 * `paginate` pass reflows the rendered document into real page-sized
 * sections (tables split by row, paragraphs by line) instead of breaking
 * only where the file stored a page break. Same `renderAsync` shape and
 * paragraph/class output as docx-preview, so the mount is shared; the
 * column is here to measure what measured pagination costs and how close
 * its page breaks land to Word's.
 */
export async function openApolloDocx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  assertZip(bytes);
  const { renderAsync } = await import("@apollo-design/docx-preview");
  if (signal.aborted) throw abortError();
  return mountDocxPreviewLineage(bytes, host, signal, t0, (bodyHost, styleHost) =>
    renderAsync(bytes.slice(), bodyHost, styleHost, {
      className: "bench-docx",
      inWrapper: true,
      breakPages: true,
      paginate: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      useBase64URL: false,
    }),
  );
}
