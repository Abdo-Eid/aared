import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { fileURLToPath } from "node:url";

/**
 * Standalone visual bench. No aliases into other repos, no shims: every
 * renderer here is an npm dependency of this package.
 */
export default defineConfig({
  plugins: [react(), tailwindcss()],
  publicDir: fileURLToPath(new URL("../fixtures", import.meta.url)),
  server: {
    port: 4321,
    open: false,
  },
});
