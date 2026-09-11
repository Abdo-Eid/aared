export type Format = "pdf" | "docx" | "pptx" | "xlsx";

export const FORMATS: readonly Format[] = ["pdf", "docx", "pptx", "xlsx"];

/** Solid badge fills, matching the product's format colours. */
export const FORMAT_BADGE: Record<Format, string> = {
  pdf: "bg-[#dc2626]",
  pptx: "bg-[#e8560a]",
  xlsx: "bg-[#047857]",
  docx: "bg-[#1d4ed8]",
};

export function formatOfFile(name: string): Format | null {
  const ext = name.split(".").pop()?.toLowerCase() ?? "";
  switch (ext) {
    case "pdf":
      return "pdf";
    case "docx":
      return "docx";
    case "pptx":
      return "pptx";
    case "xlsx":
      return "xlsx";
    default:
      return null;
  }
}

export interface SampleDef {
  name: string;
  path: string;
  note: string;
}

/** Canonical fixtures, served directly by the bench's Vite public directory. */
export const SAMPLES: readonly SampleDef[] = [
  {
    name: "arabic-pptx-rtl-coverage.pptx",
    path: "/arabic-pptx-rtl-coverage.pptx",
    note: "Flagship: Arabic RTL, table, shapes, gradients",
  },
  {
    name: "post-production-review.pptx",
    path: "/post-production-review.pptx",
    note: "Generated deck",
  },
  {
    name: "post-production-tracker.xlsx",
    path: "/post-production-tracker.xlsx",
    note: "Merges + widths",
  },
  {
    name: "arabic-xlsx-rtl-coverage.xlsx",
    path: "/arabic-xlsx-rtl-coverage.xlsx",
    note: "Arabic RTL, merges + frozen panes, deep log",
  },
  {
    name: "arabic-docx-bidi-coverage.docx",
    path: "/arabic-docx-bidi-coverage.docx",
    note: "Flow + headings, Arabic RTL",
  },
  {
    name: "delivery-report.pdf",
    path: "/delivery-report.pdf",
    note: "Control group",
  },
];
