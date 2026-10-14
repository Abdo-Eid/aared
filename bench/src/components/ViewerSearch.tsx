import { useEffect, useState } from "react";
import type { SearchNavigationState } from "@aared/viewer/preview";
import type { UsePreviewReady } from "@aared/viewer/react";
import { ChevronDown, ChevronUp, Search } from "lucide-react";
import { IconButton } from "./chrome";

/** App-owned find bar using the shared viewer search and navigation APIs. */
export function ViewerSearch({ viewer, index }: { viewer: UsePreviewReady["viewer"]; index: number }) {
  const [query, setQuery] = useState("");
  const [position, setPosition] = useState<SearchNavigationState>({ index: null, count: 0 });
  const [searching, setSearching] = useState(false);
  const [moving, setMoving] = useState(false);
  const [error, setError] = useState(false);
  const sheetIndex = viewer.format === "xlsx" ? index : 0;
  const label = viewer.format === "xlsx" ? "Search this sheet" : "Search this document";

  useEffect(() => {
    const update = (): void => {
      const { index, count } = viewer.searchNavigation;
      setPosition({ index, count });
    };
    viewer.addEventListener("searchchange", update);
    update();
    return () => viewer.removeEventListener("searchchange", update);
  }, [viewer]);

  useEffect(() => {
    const controller = new AbortController();
    setSearching(query.length > 0);
    setError(false);
    void (async () => {
      try {
        await viewer.search(query, { signal: controller.signal });
        if (controller.signal.aborted) return;
        if (viewer.searchNavigation.count > 0 && viewer.searchNavigation.index === null) {
          await viewer.searchNavigation.next();
        }
      } catch {
        if (!controller.signal.aborted) setError(true);
      } finally {
        if (!controller.signal.aborted) setSearching(false);
      }
    })();
    return () => controller.abort();
  }, [query, viewer, sheetIndex]);

  useEffect(() => () => { void viewer.search("").catch(() => {}); }, [viewer]);

  const step = (direction: "next" | "previous"): void => {
    const focus = document.activeElement;
    setMoving(true);
    setError(false);
    void viewer.searchNavigation[direction]().catch(() => setError(true)).finally(() => {
      setMoving(false);
      if (focus instanceof HTMLElement && focus.isConnected) focus.focus({ preventScroll: true });
    });
  };
  const disabled = searching || moving || position.count === 0;

  return (
    <section aria-label={label} className="shrink-0 rounded-xl border border-border/60 bg-card px-3 py-2">
      <div className="flex items-center gap-2">
        <label className="flex h-9 min-w-0 flex-1 items-center gap-2 rounded-lg border border-border bg-background px-2.5 focus-within:ring-2 focus-within:ring-accent-blue/40">
          <Search aria-hidden="true" className="size-4 shrink-0 text-muted-foreground" />
          <span className="sr-only">{label}</span>
          <input type="search" autoFocus value={query}
            onChange={(event) => setQuery(event.currentTarget.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter" && !disabled) { event.preventDefault(); step(event.shiftKey ? "previous" : "next"); }
            }}
            placeholder={viewer.format === "xlsx" ? "Search sheet…" : "Find…"}
            className="min-w-0 flex-1 bg-transparent text-sm outline-none placeholder:text-muted-foreground"
          />
        </label>
        <output aria-live="polite" className="min-w-16 shrink-0 text-center text-xs tabular-nums text-muted-foreground">
          {error ? "Search failed" : searching ? "Searching…" : query.length === 0 ? "" : position.count === 0 ? "No results" : `${(position.index ?? -1) + 1} of ${position.count}`}
        </output>
        <IconButton label="Previous match" disabled={disabled} onClick={() => step("previous")}>
          <ChevronUp aria-hidden="true" className="size-4" />
        </IconButton>
        <IconButton label="Next match" disabled={disabled} onClick={() => step("next")}>
          <ChevronDown aria-hidden="true" className="size-4" />
        </IconButton>
      </div>
    </section>
  );
}
