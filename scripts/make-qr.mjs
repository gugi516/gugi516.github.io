// QR code generator — creates KR / EN QR codes as SVG + PNG.
//
// QR codes point to the /go (KR) and /go-en (EN) redirect pages, NOT directly
// to the final URLs. The redirect destination lives in site.json → qrTargets,
// so printed QR codes stay valid even if the destination changes later.
//
// Run:  npm run qr
// Output: public/qr/qr_kr.svg, qr_kr.png, qr_en.svg, qr_en.png

import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

const ROOT = new URL('..', import.meta.url).pathname.replace(/^\//, '');
const OUT  = join(ROOT, 'public/qr');
const SITE = 'https://gugi516.github.io';

// QR points to redirect pages (permanent), not final URLs.
const TARGETS = [
  { name: 'qr_kr', url: `${SITE}/go` },
  { name: 'qr_en', url: `${SITE}/go-en` },
];

// Color variants. Default file (no suffix) = brand carbon blue; "_black" = pure black.
const VARIANTS = [
  { suffix: '',       dark: '#1e3a5f' },  // brand carbon blue
  { suffix: '_black', dark: '#000000' },  // pure black (universal print)
];

// High error correction (H) so the QR stays scannable with print wear / center logo.
const BASE = { errorCorrectionLevel: 'H', margin: 2 };

async function main() {
  let QRCode;
  try {
    QRCode = (await import('qrcode')).default;
  } catch {
    console.error('[make-qr] qrcode not installed. Run: npm install --save-dev qrcode');
    process.exit(1);
  }

  await mkdir(OUT, { recursive: true });

  for (const t of TARGETS) {
    for (const v of VARIANTS) {
      const opts = { ...BASE, color: { dark: v.dark, light: '#ffffff' } };
      const fname = `${t.name}${v.suffix}`;

      const svg = await QRCode.toString(t.url, { ...opts, type: 'svg' });
      await writeFile(join(OUT, `${fname}.svg`), svg, 'utf8');

      const pngBuf = await QRCode.toBuffer(t.url, { ...opts, type: 'png', width: 1024 });
      await writeFile(join(OUT, `${fname}.png`), pngBuf);

      console.log(`  ${fname}  → ${t.url}  (${v.dark}, svg + png 1024px)`);
    }
  }
  console.log('[make-qr] done — public/qr/');
}

main().catch((err) => { console.error('[make-qr] error:', err); process.exit(1); });
