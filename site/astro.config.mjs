import { defineConfig } from 'astro/config';
import react from '@astrojs/react';
import mdx from '@astrojs/mdx';

// Domain-root deploy (no subpath `base`): routes are `/`, `/playground/`,
// `/api/`. If this ever moves under a project subpath again, every
// root-relative URL (asset `src="/…"`, header links, the render Worker's
// URL) must be re-prefixed together — a wrong `base` silently breaks the
// Worker only in production builds, never in dev.
const site = 'https://abdo-eid.github.io/';
const base = '/';

export default defineConfig({
  site,
  base,
  integrations: [
    react(),
    mdx(),
  ],
  markdown: {
    shikiConfig: {
      theme: 'github-dark',
    },
  },
});
