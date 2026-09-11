import { useState } from "react";
import { usePreview } from "aared/react";

export interface HooksFixture {
  label: string;
  url: string;
}

/**
 * React embedding: the hook owns the mount lifecycle — abort handling,
 * loading/error state, and teardown on unmount or source change. The app
 * renders a target div and reads state.
 */
export function HooksDemo({ fixtures }: { fixtures: HooksFixture[] }) {
  const [choice, setChoice] = useState(0);
  const [file, setFile] = useState<File | null>(null);
  const source: string | File = file ?? fixtures[choice]?.url ?? "";
  const { ref, status, viewer, error } = usePreview({ source });

  return (
    <div
      data-hooks-frame
      data-state={status}
      data-format={viewer?.format ?? ""}
      data-units={viewer ? String(viewer.info.count) : ""}
    >
      <div className="demo-picker">
        <select
          aria-label="Fixture"
          value={choice}
          disabled={file !== null}
          onChange={(event) => setChoice(Number(event.target.value))}
        >
          {fixtures.map((fixture, index) => (
            <option key={fixture.label} value={index}>
              {fixture.label}
            </option>
          ))}
        </select>
        <input
          type="file"
          accept=".pptx,.xlsx,.docx"
          aria-label="Local file"
          onChange={(event) => setFile(event.target.files?.[0] ?? null)}
        />
        {file !== null && (
          <button type="button" onClick={() => setFile(null)}>
            Back to fixtures
          </button>
        )}
      </div>
      <p data-hooks-status>
        {status === "loading" && "Loading…"}
        {status === "error" && (error instanceof Error ? error.message : String(error))}
        {status === "ready" && viewer !== null && `${viewer.format} · ${viewer.info.count} units`}
      </p>
      <div ref={ref} data-stage data-state={status} className="demo-stage" />
    </div>
  );
}
