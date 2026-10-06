import { defineCollection } from "astro:content";
import { glob } from "astro/loaders";
import { z } from "astro/zod";

export const collections = {
  docs: defineCollection({
    loader: glob({ pattern: "*.mdx", base: "./src/content/docs" }),
    schema: z.object({
      title: z.string(),
      description: z.string().optional(),
      /** List the page's headings beside it; the playground's tabs need the width. */
      sidebar: z.boolean().default(true),
    }),
  }),
};
