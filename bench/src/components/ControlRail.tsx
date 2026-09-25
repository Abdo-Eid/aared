import { useState } from "react";
import { Upload } from "lucide-react";
import { SAMPLES, type Format, type SampleDef } from "../lib/format";
import { FORMAT_BADGE } from "../lib/format";
import { cn, FileBadge } from "./chrome";

interface ControlRailProps {
  activeSample: string | null;
  onSample: (sample: SampleDef) => void;
  onFile: (file: File | undefined) => void;
  notice: string | null;
}

export function ControlRail({
  activeSample,
  onSample,
  onFile,
  notice,
}: ControlRailProps) {
  const [dragging, setDragging] = useState(false);

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

    </div>
  );
}
