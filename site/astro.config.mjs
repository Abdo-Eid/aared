import { defineConfig } from 'astro/config';
import react from '@astrojs/react';
import mdx from '@astrojs/mdx';
import sitemap from '@astrojs/sitemap';
import { codeThemes } from './src/lib/code-themes.ts';

// GitHub project Pages serves the built directory beneath /aared/. Astro/Vite
// prefixes emitted assets; sitePath handles UI URLs and MDX uses sibling links.
const site = 'https://abdo-eid.github.io/aared/';
const base = '/aared/';

export default defineConfig({
  site,
  base,
  trailingSlash: 'always',
  integrations: [
    react(),
    mdx(),
    sitemap(),
  ],
  markdown: {
    // Light and dark token colours both ship; the site theme picks one.
    shikiConfig: {
      themes: codeThemes,
      defaultColor: false,
    },
  },
});
