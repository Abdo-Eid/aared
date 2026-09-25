export type Format = "pdf" | "docx" | "pptx" | "xlsx";

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
}

/** Canonical fixtures, served directly by the bench's Vite public directory. */
export const SAMPLES: readonly SampleDef[] = [
  { name: "pptx-coverage.pptx", path: "/pptx-coverage.pptx" },
  { name: "post-production-review.pptx", path: "/post-production-review.pptx" },
  { name: "post-production-tracker.xlsx", path: "/post-production-tracker.xlsx" },
  { name: "xlsx-coverage.xlsx", path: "/xlsx-coverage.xlsx" },
  { name: "docx-coverage.docx", path: "/docx-coverage.docx" },
  { name: "delivery-report.pdf", path: "/delivery-report.pdf" },
];
