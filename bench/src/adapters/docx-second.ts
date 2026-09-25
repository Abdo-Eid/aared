import {
  abortError,
  type OpenedDoc,
  type OutlineEntry,
} from "./types";
import { maxSectionWidthPx, scrapeDocxOutline } from "./docx";

/**
 * docx-renderer (Recon fork of docx-preview-sync → docx-preview). Same
 * `render(data, body, styles?, options?)` shape as the docx-preview baseline
 * and the same paragraph/class output, so the outline scraper is shared —
 * this column measures what the fork's pagination/overlay work costs.
 */
export async function openDocxRenderer(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { render } = await import("docx-renderer");
  if (signal.aborted) throw abortError();

  host.replaceChildren();
  const styleHost = document.createElement("div");
  styleHost.className = "bench-docx-styles";
  styleHost.hidden = true;
  const bodyHost = document.createElement("div");
  bodyHost.className = "bench-docx";
  host.append(styleHost, bodyHost);

  // Section widths parse concurrently with the render so the fit width
  // costs no extra wall time.
  const sectionWidth = maxSectionWidthPx(bytes);
  await render(bytes.slice(), bodyHost, styleHost, {
    className: "bench-docx",
    breakPages: true,
  });
  if (signal.aborted) throw abortError();

  const sections = Array.from(bodyHost.querySelectorAll("section.docx"));
  const pages =
    sections.length > 0
      ? sections
      : Array.from(bodyHost.querySelectorAll("section"));
  const first = pages[0];
  // Fit the document's true widest page (file sections), not just the
  // widest rendered one: a renderer that flows a landscape section onto
  // portrait-width pages must still open at the aared column's zoom.
  const paperWidth = Math.max(
    first instanceof HTMLElement ? first.offsetWidth : 0,
    await sectionWidth,
  );

  const outline: OutlineEntry[] = [];
  scrapeDocxOutline(bodyHost, outline);

  const ms = Math.round(performance.now() - t0);
  return {
    format: "docx",
    pageCount: Math.max(pages.length, 1),
    pages: pages
      .filter(
        (element): element is HTMLElement => element instanceof HTMLElement,
      )
      .map((element) => ({
        width: element.offsetWidth,
        height: element.offsetHeight,
      })),
    ms,
    outline,
    paperWidth,
    destroy: () => {
      host.replaceChildren();
    },
  };
}
