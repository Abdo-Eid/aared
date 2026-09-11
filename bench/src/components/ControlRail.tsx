import { useState } from "react";
import { Trash2, Upload } from "lucide-react";
import { SAMPLES, type Format, type SampleDef } from "../lib/format";
import { formatSize, type RunRecord } from "../lib/doc";
import { FORMAT_BADGE } from "../lib/format";
import { cn, FileBadge } from "./chrome";

/** What the eye should verify, per format. The human is the oracle here. */
const CHECKS: Record<Format, string[]> = {
  docx: [
    "Headings rail comes from the document's own heading styles.",
    "Text selects — real DOM, not a picture.",
    "Arabic: RTL flow, bullets on the right; the English paragraph after must not inherit the direction.",
    "aared layout mode paginates into physical pages; the pager count agrees with the page stack.",
  ],
  pptx: [
    "Compare Arabic paragraph direction and bullet side between the DOM/SVG and canvas renderers.",
    "RTL tables: check column order — reversed columns misread the data.",
    "aared text should select, search, and reach assistive technology; canvas output cannot.",
    "Pager steps slides; the count agrees with the deck.",
  ],
  xlsx: [
    "Values read formatted (dates, currency), not as serial numbers.",
    "Compare merged cells, widths, fills, fonts, borders, frozen panes, and RTL layout.",
    "Tabs switch sheets; aared virtualizes oversized sheets instead of truncating them.",
  ],
  pdf: [
    "Control group: text selects and find-in-page works.",
    "Rail shows page thumbnails; the pager follows the scroll.",
  ],
};

interface ControlRailProps {
  activeSample: string | null;
  docFormat: Format | null;
  onSample: (sample: SampleDef) => void;
  onFile: (file: File | undefined) => void;
  history: RunRecord[];
  onClearHistory: () => void;
  notice: string | null;
}

export function ControlRail({
  activeSample,
  docFormat,
  onSample,
  onFile,
  history,
  onClearHistory,
  notice,
}: ControlRailProps) {
  const [dragging, setDragging] = useState(false);
  const checks = docFormat ? CHECKS[docFormat] : null;

  return (
    <div className="grid content-start gap-5">
      {notice ? (
        <p
          role="alert"
          className="rounded-lg border border-warning/40 bg-warning-soft px-3 py-2 text-xs text-warning"
        >
          {notice}
        </p>
      ) : null}

      <section>
        <h2 className="mb-2 font-mono text-[0.6875rem] uppercase tracking-widest text-muted-foreground">
          Samples
        </h2>
        <ul className="grid gap-1.5">
          {SAMPLES.map((sample) => {
            const format = sample.name.split(".").pop() ?? "";
            const active = activeSample === sample.path;
            return (
              <li key={sample.path}>
                <button
                  type="button"
                  onClick={() => onSample(sample)}
                  aria-pressed={active}
                  className={cn(
                    "focus-ring flex w-full cursor-pointer items-center gap-2.5 rounded-lg border p-2 text-left transition-colors",
                    active
                      ? "border-accent-blue bg-accent-blue/5"
                      : "border-border bg-card hover:border-accent-blue/50",
                  )}
                >
                  <FileBadge
                    className={
                      FORMAT_BADGE[format as Format] ?? "bg-neutral-soft"
                    }
                    label={format}
                  />
                  <span className="min-w-0">
                    <span className="block truncate text-[0.8125rem] font-semibold">
                      {sample.name}
                    </span>
                    <span className="block truncate text-xs text-muted-foreground">
                      {sample.note}
                    </span>
                  </span>
                </button>
              </li>
            );
          })}
        </ul>
      </section>

      <section>
        <h2 className="mb-2 font-mono text-[0.6875rem] uppercase tracking-widest text-muted-foreground">
          Open a file
        </h2>
        <label
          className={cn(
            "flex cursor-pointer flex-col items-center gap-1.5 rounded-lg border border-dashed bg-card p-5 text-center transition-colors",
            dragging
              ? "border-accent-blue bg-accent-blue/5"
              : "border-border hover:border-accent-blue/60",
          )}
          onDragEnter={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={(e) => {
            e.preventDefault();
            setDragging(false);
          }}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            onFile(e.dataTransfer.files[0]);
          }}
        >
          <Upload
            aria-hidden="true"
            className="size-4 text-muted-foreground"
          />
          <span className="text-[0.8125rem] font-medium">
            Drop a file, or click to choose
          </span>
          <span className="font-mono text-[0.6875rem] text-muted-foreground">
            PDF · DOCX · PPTX · XLSX
          </span>
          <input
            accept=".pdf,.docx,.pptx,.xlsx"
            className="sr-only"
            onChange={(e) => onFile(e.target.files?.[0])}
            type="file"
          />
        </label>
      </section>

      <section>
        <h2 className="mb-2 font-mono text-[0.6875rem] uppercase tracking-widest text-muted-foreground">
          What to look for
        </h2>
        {checks ? (
          <ul className="grid gap-2.5">
            {checks.map((check) => (
              <li
                key={check}
                className="grid grid-cols-[0.1875rem_minmax(0,1fr)] gap-2 text-[0.8125rem] leading-snug text-muted-foreground"
              >
                <span aria-hidden="true" className="rounded-sm bg-border" />
                <span>{check}</span>
              </li>
            ))}
          </ul>
        ) : (
          <p className="text-[0.8125rem] leading-snug text-muted-foreground">
            Open a file to get format-specific checks.
          </p>
        )}
      </section>

      <section>
        <div className="mb-2 flex items-center justify-between">
          <h2 className="font-mono text-[0.6875rem] uppercase tracking-widest text-muted-foreground">
            Runs
          </h2>
          {history.length > 0 ? (
            <button
              type="button"
              onClick={onClearHistory}
              className="focus-ring flex cursor-pointer items-center gap-1 rounded-md px-1.5 py-1 font-mono text-[0.6875rem] text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
            >
              <Trash2 aria-hidden="true" className="size-3" />
              Clear
            </button>
          ) : null}
        </div>
        {history.length === 0 ? (
          <p className="text-[0.8125rem] leading-snug text-muted-foreground">
            Every open lands here with its milliseconds.
          </p>
        ) : (
          <ul className="grid gap-1">
            {history.map((run) => (
              <li
                key={run.key}
                title={run.error ?? `${run.name} · ${formatSize(run.sizeBytes)}`}
                className="flex items-center gap-2 rounded-md bg-card px-2 py-1.5 font-mono text-[0.6875rem]"
              >
                <span
                  className={cn(
                    "shrink-0 rounded px-1 py-px text-[0.625rem] font-bold uppercase",
                    run.renderer === "aared"
                      ? "bg-accent-blue/10 text-accent-blue"
                      : "bg-muted text-muted-foreground",
                  )}
                >
                  {run.renderer === "aared" ? "aared" : "base"}
                </span>
                <span className="min-w-0 flex-1 truncate text-foreground">
                  {run.name}
                </span>
                <span
                  className={cn(
                    "shrink-0",
                    run.error ? "text-warning" : "text-muted-foreground",
                  )}
                >
                  {run.ms === null
                    ? "—"
                    : `${run.ms.toLocaleString("en-US")} ms`}
                </span>
              </li>
            ))}
          </ul>
        )}
      </section>
    </div>
  );
}
