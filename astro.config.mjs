import { defineConfig } from 'astro/config';
import mdx from '@astrojs/mdx';
import tailwind from "@astrojs/tailwind";

// https://astro.build/config
// Sitemap is maintained manually at public/sitemap.xml (only "/" and "/en"),
// which avoids the @astrojs/sitemap×Astro4 build bug and lets us include
// precise hreflang alternates.
export default defineConfig({
  site: 'https://gugi516.github.io',
  integrations: [mdx(), tailwind()],
});
