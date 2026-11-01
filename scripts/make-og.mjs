// Generates a 1200×630 Open Graph share banner (public/og-image.jpg).
//
// Composition: deep boreal-blue gradient + azure accents (site palette),
// name / title / focus set as centred type. Text only — the site carries no
// profile photo. English text only (avoids missing-Korean-font issues in librsvg).
//
// Run:  node scripts/make-og.mjs   (re-run whenever the banner text changes)

import sharp from 'sharp';

const W = 1200, H = 630;
const OUT = 'public/og-image.jpg';
const MID = W / 2;                   // text is centred on the canvas

// ── background + text ──
const bg = Buffer.from(`
<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#0c1420"/>
      <stop offset="0.55" stop-color="#16202f"/>
      <stop offset="1" stop-color="#1e3a5f"/>
    </linearGradient>
  </defs>
  <rect width="${W}" height="${H}" fill="url(#g)"/>

  <!-- subtle grid texture -->
  <g opacity="0.05" stroke="#ffffff" stroke-width="1">
    ${Array.from({ length: 13 }, (_, i) => `<line x1="${i * 100}" y1="0" x2="${i * 100}" y2="${H}"/>`).join('')}
    ${Array.from({ length: 7 }, (_, i) => `<line x1="0" y1="${i * 100}" x2="${W}" y2="${i * 100}"/>`).join('')}
  </g>

  <!-- eyebrow + azure accent -->
  <text x="${MID}" y="196" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="22" letter-spacing="5" fill="#41a0e0" font-weight="700">RESEARCH PORTFOLIO</text>
  <rect x="${MID - 29}" y="216" width="58" height="5" rx="2.5" fill="#41a0e0"/>

  <!-- name -->
  <text x="${MID}" y="312" text-anchor="middle" font-family="Georgia, 'Times New Roman', serif" font-size="74" fill="#ffffff" font-weight="700">Gukhwa Jang</text>

  <!-- title + focus -->
  <text x="${MID}" y="392" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="30" fill="#dbdfe6">Ph.D. Researcher · Seoul National University</text>
  <text x="${MID}" y="436" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="25" fill="#a9b5c6">Urban Microclimate · Shrinking Cities · Urban Design</text>

  <!-- leaf-green secondary accent: a street-tree row (a hairline rule with
       five tree dots sitting on it) separating the focus line from the URL.
       Green is the site's secondary accent — the banner stays blue-dominant. -->
  <g fill="#4aa96c">
    <rect x="${MID - 210}" y="491" width="420" height="2" rx="1" opacity="0.5"/>
    ${[-160, -80, 0, 80, 160].map((dx) => `<circle cx="${MID + dx}" cy="492" r="5" opacity="0.95"/>`).join('')}
  </g>

  <!-- url -->
  <text x="${MID}" y="552" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="26" fill="#41a0e0" font-weight="600">gugi516.github.io</text>
</svg>`);

await sharp(bg)
  .jpeg({ quality: 90 })
  .toFile(OUT);

console.log(`[make-og] wrote ${OUT} (${W}x${H})`);
