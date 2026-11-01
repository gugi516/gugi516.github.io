# -*- coding: utf-8 -*-
import io
from pathlib import Path
p = str(Path(__file__).resolve().parents[1] / 'src' / 'data' / 'site.json')
s = io.open(p, encoding="utf-8").read()
lines = s.split("\n")

TITLE = "Bi-temporal ALS assessment of vertical–horizontal forest structures and their structural association in a temperate urban forest"
new_item = [
'      {',
'        "id": 10,',
'        "date": "2026/08/17",',
'        "year": 2026,',
'        "role": "first",',
'        "journal": "Forest Science and Technology",',
'        "if": "2.4 (2026)",',
'        "quartile": "Q2 (2026)",',
'        "doi": "10.1080/21580103.2026.2712965",',
'        "authors": ["Seunghyeon Lee", "Yonghwan Kim", "Dohee Kim", "Hansoo Kim", "Youngkeun Song"],',
'        "title": {',
f'          "ko": "{TITLE}",',
f'          "en": "{TITLE}"',
'        }',
'      },',
]

# 1) publications id=9 앞의 '{' 앞에 신규 항목 삽입
i9 = next(i for i,l in enumerate(lines) if l.strip()=='"id": 9,')
b = i9
while lines[b].strip() != '{':
    b -= 1
assert 0 < b < i9
lines = lines[:b] + new_item + lines[b:]
print(f"publications id=10 삽입 (line {b})")

# 2) underReview R2 (Forest Science and Technology + status) 줄 제거
rm = [i for i,l in enumerate(lines) if 'Forest Science and Technology' in l and '"status"' in l]
assert len(rm)==1, f"제거 대상 {len(rm)}건 (1이어야 함)"
print("underReview 제거 line:", rm[0], '→', lines[rm[0]].strip()[:60])
del lines[rm[0]]

s2 = "\n".join(lines)
# 3) note 카운트 5 → 4
assert s2.count("투고 · 심사 중 (5편)")==1 and s2.count("under peer review (5)")==1
s2 = s2.replace("투고 · 심사 중 (5편)", "투고 · 심사 중 (4편)").replace("under peer review (5)", "under peer review (4)")
print("note 카운트 5 → 4")

io.open(p, "w", encoding="utf-8", newline="\n").write(s2)
print("site.json 저장 완료")

# JSON 파싱 검증
import json
json.load(io.open(p, encoding="utf-8"))
print("JSON 파싱 OK")
