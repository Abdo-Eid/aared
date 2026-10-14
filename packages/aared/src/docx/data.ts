import type { DocumentSection } from "../core/model.js";

/** Source section metadata, not a physical rendered page. All lengths are inches. */
export interface SectionInfo {
  index: number;
  breakType: "continuous" | "nextPage" | "evenPage" | "oddPage";
  page: {
    width: number;
    height: number;
    orientation: "portrait" | "landscape";
    margins: { top: number; right: number; bottom: number; left: number; header: number; footer: number; gutter: number };
  };
}

export function sectionInfo(section: DocumentSection, index: number): SectionInfo {
  const { breakType, page } = section.properties;
  const m = page.margins;
  return { index, breakType, page: {
    width: page.widthTwips / 1440, height: page.heightTwips / 1440, orientation: page.orientation,
    margins: { top: m.top / 1440, right: m.right / 1440, bottom: m.bottom / 1440, left: m.left / 1440,
      header: m.header / 1440, footer: m.footer / 1440, gutter: m.gutter / 1440 },
  } };
}
