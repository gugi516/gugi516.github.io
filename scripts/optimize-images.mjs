// Image optimizer — runs before `astro build`.
//
// Strategy: ORIGINALS are preserved. For each cover image we write a
// downscaled, web-optimized WebP into a `web/` subfolder next to it.
// The site references the `web/` versions; originals stay untouched as masters.
//
//   public/papers/covers/2026_x.webp        ← original (any size, kept)
//   public/papers/covers/web/2026_x.webp    ← 800px web version (generated)
//
// Why: high-res (2K-3K) covers look blurry when the browser downsamples them
// into ~290px cards. A pre-scaled 800px WebP downsamples cleanly → crisp text.
//
// Run:  npm run optimize-images   (build runs it automatically)
//       add --force to ignore the up-to-date check and regenerate everything.

import { mkdir, readdir, stat, writeFile } from 'node:fs/promises';
import { join, extname, basename } from 'node:path';

const TARGET_WIDTH = 800;
const WEBP_QUALITY = 92;
const WEBP_EFFORT  = 6;

const FORCE = process.argv.includes('--force');

const ROOT = new URL('..', import.meta.url).pathname.replace(/^\//, '');
const DIRS = [
  'public/papers/covers',
  'public/patents/covers',
  'public/books/covers',
];

const SOURCE_EXT = new Set(['.jpg', '.jpeg', '.png', '.webp']);

async function main() {
  let sharp;
  try {
    sharp = (await import('sharp')).default;
  } catch {
    console.log('[optimize-images] sharp not available — skipping.');
    return;
  }

  let generated = 0;
  let skipped = 0;

  for (const dir of DIRS) {
    const abs = join(ROOT, dir);
    const webDir = join(abs, 'web');
    let files;
    try {
      files = await readdir(abs);
    } catch {
      continue;
    }
    await mkdir(webDir, { recursive: true });

    for (const name of files) {
      const ext = extname(name).toLowerCase();
      if (!SOURCE_EXT.has(ext)) continue;

      const srcPath = join(abs, name);
      const st = await stat(srcPath);
      if (!st.isFile()) continue; // skip the web/ subdir itself

      const outName = basename(name, ext) + '.webp';
      const outPath = join(webDir, outName);

      // Skip if web version is newer than the source (unless --force)
      if (!FORCE) {
        try {
          const outSt = await stat(outPath);
          if (outSt.mtimeMs >= st.mtimeMs) { skipped++; continue; }
        } catch { /* output missing → generate */ }
      }

      try {
        const meta = await sharp(srcPath).metadata();
        let pipeline = sharp(srcPath);
        if (meta.width && meta.width > TARGET_WIDTH) {
          pipeline = pipeline.resize({ width: TARGET_WIDTH, withoutEnlargement: true });
        }
        // Write directly to the web/ path (no temp-rename → avoids Windows file-lock issues)
        await pipeline.webp({ quality: WEBP_QUALITY, effort: WEBP_EFFORT }).toFile(outPath);
        const outMeta = await sharp(outPath).metadata();
        console.log(`  ${dir}/web/${outName}  (${meta.width}px → ${outMeta.width}px)`);
        generated++;
      } catch (err) {
        console.warn(`  ⚠ failed: ${dir}/${name} — ${err.message}`);
      }
    }
  }

  console.log(`[optimize-images] done — ${generated} web versions generated, ${skipped} up-to-date.`);
}

main().catch((err) => {
  console.error('[optimize-images] error:', err);
  process.exit(0); // never fail the build
});
