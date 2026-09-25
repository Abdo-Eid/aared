import { useCallback, useEffect, useRef, useState } from "react";
import { Moon, PanelLeft, Sun } from "lucide-react";
import { BASELINES, DEFAULT_BASELINE } from "./adapters";
import { formatOfFile, type SampleDef } from "./lib/format";
import type { DocMeta, Renderer } from "./lib/doc";
import { IconButton } from "./components/chrome";
import { ControlRail } from "./components/ControlRail";
import { ViewerCard } from "./components/ViewerCard";
import { StageMessage } from "./components/chrome";
import { cn } from "./components/chrome";

type Mode = "single" | "compare";

export function App() {
  const [doc, setDoc] = useState<DocMeta | null>(null);
  const [mode, setMode] = useState<Mode>("single");
  const [singleRenderer, setSingleRenderer] = useState<Renderer>("aared");
  const [dark, setDark] = useState(false);
  const [railOpen, setRailOpen] = useState(true);
  const [notice, setNotice] = useState<string | null>(null);
  const [activeSample, setActiveSample] = useState<string | null>(null);
  const uploadUrlRef = useRef<string | null>(null);

  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
  }, [dark]);

  useEffect(() => {
    const url = uploadUrlRef.current;
    return () => {
      if (url) URL.revokeObjectURL(url);
    };
  }, []);

  const openBytes = useCallback(
    (
      name: string,
      bytes: Uint8Array,
      url: string,
      origin: DocMeta["origin"],
    ) => {
      const format = formatOfFile(name);
      if (!format) {
        setNotice(`No renderer for this file type: ${name}`);
        return;
      }
      setNotice(null);
      setDoc({
        key: `${origin}:${name}:${bytes.byteLength}:${Date.now()}`,
        name,
        format,
        sizeBytes: bytes.byteLength,
        url,
        bytes,
        origin,
      });
    },
    [],
  );

  const openSample = useCallback(
    async (sample: SampleDef) => {
      setNotice(null);
      try {
        const response = await fetch(sample.path);
        if (!response.ok) {
          throw new Error(`Sample missing (${response.status}): ${sample.name}`);
        }
        const buffer = await response.arrayBuffer();
        openBytes(sample.name, new Uint8Array(buffer), sample.path, "sample");
        setActiveSample(sample.path);
      } catch (error) {
        setNotice(
          error instanceof Error ? error.message : "Could not load sample.",
        );
      }
    },
    [openBytes],
  );

  const openFile = useCallback(
    async (file: File | undefined) => {
      if (!file) return;
      setNotice(null);
      const buffer = await file.arrayBuffer();
      if (uploadUrlRef.current) URL.revokeObjectURL(uploadUrlRef.current);
      const url = URL.createObjectURL(file);
      uploadUrlRef.current = url;
      openBytes(file.name, new Uint8Array(buffer), url, "upload");
      setActiveSample(null);
    },
    [openBytes],
  );

  const closeDoc = useCallback(() => {
    setDoc(null);
  }, []);

  const baselineOptions = doc !== null ? Object.keys(BASELINES[doc.format]) : [];
  const baselineDefault =
    doc !== null ? DEFAULT_BASELINE[doc.format] : "baseline";
  // Single-mode value: "aared" or an explicit baseline id ("baseline" legacy
  // maps to the per-format default). Compare mode steers its left card with
  // the same control — left is "the other one" when aared is selected.
  const effectiveSingle: Renderer =
    singleRenderer === "baseline" ? (baselineDefault as Renderer) : singleRenderer;
  const compareLeft: Renderer =
    effectiveSingle === "aared" ? (baselineDefault as Renderer) : effectiveSingle;

  return (
    <div className="flex h-dvh flex-col bg-background text-foreground">
      <header className="flex h-12 shrink-0 items-center gap-3 border-b border-border bg-card px-4">
        <IconButton
          label={railOpen ? "Hide sidebar" : "Show sidebar"}
          onClick={() => setRailOpen((value) => !value)}
          pressed={railOpen}
        >
          <PanelLeft className="size-4" />
        </IconButton>
        <div className="flex min-w-0 items-baseline gap-2">
          <h1 className="text-sm font-extrabold tracking-tight">
            aared bench
          </h1>
          <p className="hidden truncate text-xs text-muted-foreground sm:block">
            baseline renderers · open-speed harness
          </p>
        </div>
        <div className="ml-auto flex shrink-0 items-center gap-2">
          <div
            role="group"
            aria-label="View mode"
            className="flex items-center gap-0.5 rounded-lg border border-border p-0.5"
          >
            {(["single", "compare"] as const).map((value) => (
              <button
                key={value}
                type="button"
                onClick={() => setMode(value)}
                aria-pressed={mode === value}
                className={cn(
                  "focus-ring cursor-pointer rounded-md px-2.5 py-1 text-xs font-semibold capitalize transition-colors",
                  mode === value
                    ? "bg-foreground text-background"
                    : "text-muted-foreground hover:text-foreground",
                )}
              >
                {value}
              </button>
            ))}
          </div>
          {doc !== null && doc.format !== "pdf" ? (
            <label className="flex items-center gap-1.5 rounded-lg border border-border p-0.5 pl-2 text-xs">
              <span className="font-semibold text-muted-foreground">
                {mode === "single" ? "Renderer" : "Compare vs"}
              </span>
              <select
                aria-label={mode === "single" ? "Single renderer" : "Baseline renderer to compare against"}
                className="focus-ring cursor-pointer rounded-md bg-transparent px-1 py-1 font-semibold"
                value={mode === "single" ? effectiveSingle : compareLeft}
                onChange={(e) => setSingleRenderer(e.target.value as Renderer)}
              >
                {mode === "single" ? (
                  <option value="aared">aared</option>
                ) : null}
                {baselineOptions.map((id) => (
                  <option key={id} value={id}>
                    {id}
                  </option>
                ))}
              </select>
            </label>
          ) : null}
          <IconButton
            label={dark ? "Switch to light theme" : "Switch to dark theme"}
            onClick={() => setDark((value) => !value)}
            pressed={dark}
          >
            {dark ? (
              <Sun className="size-4" />
            ) : (
              <Moon className="size-4" />
            )}
          </IconButton>
        </div>
      </header>

      <div className="flex min-h-0 flex-1">
        {railOpen ? (
          <aside className="w-80 shrink-0 overflow-y-auto border-r border-border bg-card p-5">
            <ControlRail
              activeSample={activeSample}
              onSample={(sample) => void openSample(sample)}
              onFile={(file) => void openFile(file)}
              notice={notice}
            />
          </aside>
        ) : null}

        <main
          className="flex min-w-0 flex-1 gap-2 overflow-hidden p-2"
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => {
            e.preventDefault();
            void openFile(e.dataTransfer.files[0]);
          }}
        >
          {!doc ? (
            <div className="flex flex-1 flex-col overflow-hidden rounded-2xl border border-dashed border-border bg-card">
              <StageMessage>
                Open a sample on the left — or drop a PDF, DOCX, PPTX or XLSX
                anywhere here.
              </StageMessage>
            </div>
          ) : mode === "single" ? (
            <ViewerCard
              key={doc.key}
              doc={doc}
              renderer={doc.format === "pdf" ? "baseline" : singleRenderer}
              theme={dark ? "dark" : "light"}
              onClose={closeDoc}
            />
          ) : (
            <>
              <ViewerCard
                key={`${doc.key}:${compareLeft}`}
                doc={doc}
                renderer={doc.format === "pdf" ? "baseline" : compareLeft}
                theme={dark ? "dark" : "light"}
                onClose={closeDoc}
              />
              <ViewerCard
                key={`${doc.key}:aared`}
                doc={doc}
                renderer="aared"
                theme={dark ? "dark" : "light"}
                onClose={closeDoc}
              />
            </>
          )}
        </main>
      </div>
    </div>
  );
}
