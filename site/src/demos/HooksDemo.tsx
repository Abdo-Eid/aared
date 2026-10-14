import { useEffect, useState } from "react";
import { usePreview } from "@aared/viewer/react";
import { messageFor } from "@aared/viewer/core";
import { viewerToolbar } from "../lib/demo-controls.js";
import { describeUnits } from "../lib/demo-frame.js";

export interface HooksFixture {
  label: string;
  url: string;
  format: string;
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
  const { ref, status, viewer, error, retry } = usePreview({ source });
  // A paging and zoom bar: the viewer's own `next()`, `zoomIn()`, … behind buttons.
  useEffect(() => {
    if (viewer === null || ref.current === null) return;
    return viewerToolbar(ref.current, viewer);
  }, [viewer, ref]);

  return (
    <div
      className="demo-frame"
      data-hooks-frame
      data-state={status}
      data-format={viewer?.format ?? ""}
      data-units={viewer ? String(viewer.info.count) : ""}
    >
      <div className="demo-picker">
        <div className="demo-pills" role="group" aria-label="File format">
          {fixtures.map((fixture, index) => (
            <button
              key={fixture.label}
              type="button"
              aria-pressed={file === null && index === choice}
              disabled={file !== null}
              onClick={() => setChoice(index)}
            >
              {fixture.format.toUpperCase()}
            </button>
          ))}
        </div>
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
      <p data-hooks-status role={status === "error" ? "alert" : "status"}>
        {status === "loading" && "Loading…"}
        {status === "error" && <>{messageFor(error)} <button type="button" onClick={retry}>Retry</button></>}
        {status === "ready" && `${viewer.format} · ${describeUnits(viewer.format, viewer.info.count)}`}
      </p>
      <div ref={ref} data-stage data-state={status} className="demo-stage" />
    </div>
  );
}
