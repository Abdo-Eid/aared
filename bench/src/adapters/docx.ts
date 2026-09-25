import { readZipEntryText } from "../lib/zip-entry";
import { abortError, type OpenedDoc, type OutlineEntry } from "./types";

function hasZipSignature(bytes: Uint8Array): boolean {
  return (
    bytes.length > 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}

/**
 * Shared outline scraper for the docx-preview lineage (docx-preview and its
 * docx-renderer fork render the same paragraph/class shape): paragraphs stay
 * `<p>`, the Word style id lands in the class
 * (`${className}_${styleId}`, lowercased — e.g. `bench-docx_heading1`), and
 * real `<h1>-<h6>` stay as a fallback.
 */
export function scrapeDocxOutline(
  bodyHost: HTMLElement,
  outline: OutlineEntry[],
): void {
  const headingClass = /heading[\s_-]?0?([1-9])/i;
  bodyHost.querySelectorAll("h1,h2,h3,h4,h5,h6,p").forEach((el) => {
    const tag = el.tagName.toLowerCase();
    const tagMatch = /^h([1-6])$/.exec(tag);
    if (tagMatch?.[1] !== undefined) {
      const text = (el.textContent ?? "").trim();
      if (text) {
        outline.push({ level: Number(tagMatch[1]), text, element: el });
      }
      return;
    }
    if (tag !== "p") return;
    const classes = el.getAttribute("class") ?? "";
    const classMatch = headingClass.exec(classes);
    if (classMatch?.[1] === undefined) return;
    const text = (el.textContent ?? "").trim();
    if (text) {
      outline.push({ level: Number(classMatch[1]), text, element: el });
    }
  });
}

/** A `w:pgSz` width in twips (twentieths of a point) as CSS pixels at 96 dpi. */
const TWIPS_TO_PX = (0.05 * 96) / 72;

/**
 * Widest `w:pgSz` section in the file, in CSS px (dxa → pt → px with the
 * same factors docx-preview itself uses, so the number matches a renderer's
 * own page width exactly). Baselines that flow every section onto
 * portrait-width pages still fit the document's true widest page this way,
 * so all bench columns open at the same zoom. Harness work, not the
 * baseline's: read once per file (the measurement page primes it before
 * measuring) and never through a library under measurement. Never throws:
 * 0 means "unparseable, fall back to rendered widths".
 */
const sectionWidths = new WeakMap<Uint8Array, Promise<number>>();

export function maxSectionWidthPx(bytes: Uint8Array): Promise<number> {
  let width = sectionWidths.get(bytes);
  if (width === undefined) {
    width = readMaxSectionWidthPx(bytes);
    sectionWidths.set(bytes, width);
  }
  return width;
}

async function readMaxSectionWidthPx(bytes: Uint8Array): Promise<number> {
  const xml = await readZipEntryText(bytes, "word/document.xml");
  if (xml === null) return 0;
  let max = 0;
  for (const match of xml.matchAll(/<w:pgSz[^>]*>/g)) {
    const width = /w:w="(\d+)"/.exec(match[0]);
    if (width?.[1] !== undefined) max = Math.max(max, Number(width[1]));
  }
  return max > 0 ? max * TWIPS_TO_PX : 0;
}
/**
 * Baseline DOCX renderer (docx-preview). Bytes go in directly — never a Blob,
 * because docx-preview's JSZip would read it through `new FileReader()`, which
 * pptxviewjs replaces (see ../adapters/pptx.ts). The style container lives beside the
 * body container so clearing one document never wipes the next one's styles.
 */
export async function openDocx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  if (!hasZipSignature(bytes)) {
    throw new Error("Not a readable Word document (missing zip signature).");
  }
  const { renderAsync } = await import("docx-preview");
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
  await renderAsync(bytes, bodyHost, styleHost, {
    className: "bench-docx",
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: false,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    useBase64URL: false,
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
      .filter((element): element is HTMLElement => element instanceof HTMLElement)
      .map((element) => ({ width: element.offsetWidth, height: element.offsetHeight })),
    ms,
    outline,
    paperWidth,
    destroy: () => {
      host.replaceChildren();
    },
  };
}
