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
 * Baseline DOCX renderer (docx-preview). Bytes go in directly — never a Blob,
 * because JSZip would read it through `new FileReader()`, which pptxviewjs
 * replaces (see ../adapters/pptx.ts). The style container lives beside the
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
  const paperWidth =
    first instanceof HTMLElement ? first.offsetWidth : 0;

  const outline: OutlineEntry[] = [];
  // docx-preview renders paragraphs as `<p>` (see renderParagraph in
  // docx-preview.mjs), never as `<h1>-<h6>`: the Word style id lands in the
  // class via processStyleName (`${className}_${escapeClassName(styleId)}`
  // with className "bench-docx", lowercased by escapeClassName), e.g. style
  // `Heading1` becomes `bench-docx_heading1`. Match Heading1-9 classes in
  // document order; real `<h1>-<h6>` stay as a fallback.
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
