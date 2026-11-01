// Build-time citation generator.
//
// Source of truth: cv/citations/publications.bib (normalized BibTeX, 9 papers).
// Output: src/data/citations.json  →  { "<doi-lowercase>": { apa, bibtex } }
//
// CardPaper.astro looks up citations[paper.doi] and, if present, shows a
// "Cite" button that opens a modal with APA + BibTeX (each copyable).
//
// APA conventions mirror cv/build_cv.py (Surname, I. I.; "A, B, & C"; title
// trailing-dash stripped, " - " → ": ").
//
// Run: npm run build-citations  (build runs it automatically before astro)

import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

const ROOT = new URL('..', import.meta.url).pathname.replace(/^\//, '');
const BIB  = join(ROOT, 'cv/citations/publications.bib');
const OUT  = join(ROOT, 'src/data/citations.json');

// ── BibTeX parsing ──
function parseEntries(text) {
  const entries = [];
  // Match @article{key, ... } with brace counting
  const re = /@(\w+)\s*\{\s*([^,]+),/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const type = m[1].toLowerCase();
    const key = m[2].trim();
    // Find the body from after the first comma until the matching closing brace
    let i = m.index + m[0].length;
    let depth = 1;
    let body = '';
    while (i < text.length && depth > 0) {
      const ch = text[i];
      if (ch === '{') depth++;
      else if (ch === '}') { depth--; if (depth === 0) break; }
      body += ch;
      i++;
    }
    entries.push({ type, key, fields: parseFields(body) });
  }
  return entries;
}

function parseFields(body) {
  const fields = {};
  // key = {value}  OR  key = value,
  const re = /(\w+)\s*=\s*(\{([^{}]*)\}|"([^"]*)"|([^,\n]+))/g;
  let m;
  while ((m = re.exec(body)) !== null) {
    const k = m[1].toLowerCase();
    const v = (m[3] ?? m[4] ?? m[5] ?? '').trim();
    fields[k] = v;
  }
  return fields;
}

// ── APA helpers ──
// Single initial per given-name token. Hyphenated romanized Korean names
// (e.g. "Youn-Ha") collapse to one initial ("Y."), not "Y.-H.".
// Space-separated given names (e.g. "Dennis Heejoon") keep one initial each
// ("D. H.").
function initials(given) {
  return given.split(/\s+/).filter(Boolean)
    .map(part => part[0].toUpperCase() + '.')
    .join(' ');
}
function apaAuthor(name) {
  let sur, giv;
  if (name.includes(',')) {
    const idx = name.indexOf(',');
    sur = name.slice(0, idx).trim();
    giv = name.slice(idx + 1).trim();
  } else {
    const ps = name.split(/\s+/);
    sur = ps[ps.length - 1];
    giv = ps.slice(0, -1).join(' ');
  }
  return `${sur}, ${initials(giv)}`.trim().replace(/,\s*$/, '');
}
function apaAuthorList(authorField) {
  const authors = authorField.split(/\s+and\s+/).map(a => apaAuthor(a.trim()));
  const n = authors.length;
  let out = '';
  for (let i = 0; i < n; i++) {
    if (i > 0) out += ', ';
    if (i === n - 1 && n > 1) out += '& ';
    out += authors[i];
  }
  return out;
}

function buildAPA(f) {
  const authors = apaAuthorList(f.author || '');
  const year = f.year || 'n.d.';
  let title = (f.title || '').replace(/\s*-\s*$/, '').replace(/ - /g, ': ').replace(/[.\s]+$/, '');
  const journal = (f.journal || '').replace(/\\&/g, '&');
  const vol = f.volume || '';
  const num = f.number || '';
  const pages = (f.pages || '').replace(/-/g, '–'); // en dash
  const doi = f.doi || '';

  let s = `${authors} (${year}). ${title}. ${journal}`;
  if (vol) {
    s += `, ${vol}`;
    if (num) s += `(${num})`;
    if (pages) s += `, ${pages}`;
  }
  s += '.';
  if (doi) s += ` https://doi.org/${doi}`;
  return s;
}

function buildBibtex(e) {
  const f = e.fields;
  const order = ['author', 'title', 'journal', 'volume', 'number', 'pages', 'year', 'month', 'doi'];
  const lines = [`@${e.type}{${e.key},`];
  const present = order.filter(k => f[k]);
  present.forEach((k, i) => {
    const comma = i < present.length - 1 ? ',' : '';
    lines.push(`  ${k.padEnd(7)} = {${f[k]}}${comma}`);
  });
  lines.push('}');
  return lines.join('\n');
}

async function main() {
  let text;
  try {
    text = await readFile(BIB, 'utf8');
  } catch {
    console.log('[build-citations] publications.bib not found — writing empty citations.json');
    await writeFile(OUT, '{}\n', 'utf8');
    return;
  }

  const entries = parseEntries(text);
  const out = {};
  for (const e of entries) {
    const doi = (e.fields.doi || '').toLowerCase().trim();
    if (!doi) continue;
    out[doi] = { apa: buildAPA(e.fields), bibtex: buildBibtex(e) };
  }

  await writeFile(OUT, JSON.stringify(out, null, 2) + '\n', 'utf8');
  console.log(`[build-citations] ${Object.keys(out).length} citations → src/data/citations.json`);
}

main().catch((err) => { console.error('[build-citations] error:', err); process.exit(0); });
