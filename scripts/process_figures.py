# -*- coding: utf-8 -*-
import os, json, io, shutil
from pathlib import Path
from PIL import Image
ROOT = str(Path(__file__).resolve().parents[1])
SLUG = "2026_FST_Bi-temporal-structures"
src_dir = os.path.join(ROOT, "public", "papers", "figures", "fst")
out_dir = os.path.join(ROOT, "public", "papers", "figures", SLUG)
inbox   = os.path.join(ROOT, "_figures_inbox", SLUG)
os.makedirs(out_dir, exist_ok=True); os.makedirs(inbox, exist_ok=True)

MAX = 2000
for n in range(1, 5):
    srcf = os.path.join(src_dir, f"Figure {n}.jpg")
    im = Image.open(srcf).convert("RGB")
    w, h = im.size
    scale = min(1.0, MAX / max(w, h))
    if scale < 1.0:
        im = im.resize((round(w*scale), round(h*scale)), Image.LANCZOS)
    outf = os.path.join(out_dir, f"fig{n}.webp")
    im.save(outf, "WEBP", quality=90, method=6)
    print(f"fig{n}.webp  {w}x{h} -> {im.size}  {os.path.getsize(outf)//1024}KB")
    shutil.move(srcf, os.path.join(inbox, f"Figure {n}.jpg"))  # 원본 → inbox(gitignore)
# 빈 fst 폴더 제거
try: os.rmdir(src_dir)
except OSError: pass

captions = [
 "Location of the study site and tree species distribution; (a) study site location in the Republic of Korea; (b) distribution of eight major tree species in the study site with expert field investigation polygon borders (thick orange outlines). Numbers in brackets in the legend indicate the number of polygons.",
 "Comparison of field investigation and ALS-derived coverage for three layers by tree species; (a) radar chart for mean coverage comparison by species (solid line with circle: field investigation; dashed line with square markers: ALS results; dark green, green, and yellow-green lines represent tree, subtree, and shrub layer results, respectively); (b) boxplot for coverage difference (Field − ALS) of each polygon for three layers by tree species.",
 "Assessment of ALS-derived vertical point distributions and horizontal coverage by tree species and layers across late-autumn and early-spring periods. (a) Vertical point distribution assessment: (a1) late-autumn period; (a2) early-spring period; (a3) difference (late-autumn minus early-spring). (b) Horizontal coverage assessment: (b1) late-autumn period; (b2) early-spring period; (b3) difference (late-autumn minus early-spring).",
 "Correlation analysis of ALS-derived vertical point distribution and horizontal coverage by layers and periods for eight tree species. The x-axis represents horizontal coverage and the y-axis represents vertical point distribution. Thick black rectangles indicate correlations analyzed for the same period.",
]

fj = os.path.join(ROOT, "src", "data", "figures.json")
data = json.load(io.open(fj, encoding="utf-8"))
data["10"] = [{"src": f"/papers/figures/{SLUG}/fig{i+1}.webp", "caption": captions[i]} for i in range(4)]
io.open(fj, "w", encoding="utf-8", newline="\n").write(json.dumps(data, ensure_ascii=False, indent=1) + "\n")
print("figures.json: id=10 (fig1-4) 추가")
