# -*- coding: utf-8 -*-
"""논문 11 (GECCO 2026) 웹 자산 생성 — 표지 webp / PDF 복사 / figure webp.

관례 근거 (실측):
  - 표지 마스터: 원본 해상도 유지(기존 covers/*.webp 는 2176~2480px 급),
    같은 저널 2025_GECCO_Assessing-Corvus.webp = 2268x3094 / 401KB
  - figure: 장변 2000px 캡, WEBP quality=90 method=6  (scripts/process_figures.py 와 동일)
  - 원본 고해상도는 _figures_inbox/(gitignore) 로 이동하지 않고 Downloads 에 그대로 둔다.

로그: _logs/process_paper11_gecco_<YYMMDD_HHMMSS>.csv (즉시 flush)
"""
import csv
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from PIL import Image

Image.MAX_IMAGE_PIXELS = None          # f8.jpg = 109.7M px (DecompressionBomb 경고 회피)

ROOT = Path(__file__).resolve().parents[1]
SRC = Path(r"C:/Users/claylee/Downloads/gecco")
SLUG = "2026_GECCO_Integrating-unmanned-aerial"
FIG_MAX = 2000
QUALITY, METHOD = 90, 6

LOGDIR = ROOT / "_logs"
LOGDIR.mkdir(exist_ok=True)
logpath = LOGDIR / f"process_paper11_gecco_{datetime.now():%y%m%d_%H%M%S}.csv"
log = open(logpath, "w", encoding="utf-8", newline="")
w = csv.writer(log)
w.writerow(["kind", "src", "src_size_px", "src_bytes", "out", "out_size_px", "out_bytes", "quality"])
log.flush()


def record(kind, s, ss, sb, o, os_, ob, q):
    w.writerow([kind, s, ss, sb, o, os_, ob, q])
    log.flush()
    print(f"{kind:7s} {s:16s} {ss:>12s} {sb//1024:>6d}KB -> {o:52s} {os_:>12s} {ob//1024:>6d}KB")


# ── 표지 ──────────────────────────────────────────────────────────────
cov_src = SRC / "표지.jpg"
cov_out = ROOT / "public" / "papers" / "covers" / f"{SLUG}.webp"
im = Image.open(cov_src).convert("RGB")
im.save(cov_out, "WEBP", quality=QUALITY, method=METHOD)
record("cover", cov_src.name, f"{im.size[0]}x{im.size[1]}", cov_src.stat().st_size,
       str(cov_out.relative_to(ROOT)), f"{im.size[0]}x{im.size[1]}", cov_out.stat().st_size, QUALITY)

# ── PDF ───────────────────────────────────────────────────────────────
pdf_src = SRC / "1-s2.0-S235198942600380X-main.pdf"
pdf_out = ROOT / "public" / "papers" / "pdfs" / f"{SLUG}.pdf"
shutil.copy2(pdf_src, pdf_out)
assert pdf_out.stat().st_size == pdf_src.stat().st_size
record("pdf", pdf_src.name[:16], "-", pdf_src.stat().st_size,
       str(pdf_out.relative_to(ROOT)), "-", pdf_out.stat().st_size, "-")

# ── figures ───────────────────────────────────────────────────────────
figdir = ROOT / "public" / "papers" / "figures" / SLUG
figdir.mkdir(parents=True, exist_ok=True)
srcs = sorted((SRC / "figures").glob("f*.*"), key=lambda p: int("".join(c for c in p.stem if c.isdigit())))
if not srcs:
    sys.exit("figures 원본 없음")
for p in srcs:
    n = int("".join(c for c in p.stem if c.isdigit()))
    im = Image.open(p).convert("RGB")
    sw, sh = im.size
    scale = min(1.0, FIG_MAX / max(sw, sh))
    if scale < 1.0:
        im = im.resize((round(sw * scale), round(sh * scale)), Image.LANCZOS)
    out = figdir / f"fig{n}.webp"
    im.save(out, "WEBP", quality=QUALITY, method=METHOD)
    record("figure", p.name, f"{sw}x{sh}", p.stat().st_size,
           str(out.relative_to(ROOT)), f"{im.size[0]}x{im.size[1]}", out.stat().st_size, QUALITY)

log.close()
print(f"\nfigure {len(srcs)}장 (원본 f1~f{len(srcs)}) / 로그: {logpath.relative_to(ROOT)}")
