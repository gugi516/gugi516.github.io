# -*- coding: utf-8 -*-
"""
normalize_bib.py — individual/ 의 .bib 들을 '통일된 양식'으로 정규화(in-place).

하는 일 (기계적 처리):
  - 깨진 export 수리: pages 뒤 콤마 누락 보정
  - 저자명 통일 → "Last, First" 형식
      · "First ... Last"  → "Last, First ..."
      · 붙은 이름 "LeeSeunghyeon" → "Lee, Seunghyeon"
      · 본인/지도교수 표기 정규화 (Seunghyeon / Youngkeun)
  - DOI 통일: URL 제거, 없으면 citekey에서 추출
  - 군더더기 필드(abstract/keywords/url/issn/publisher) 제거
  - 필드 순서 통일: author, title, journal, volume, number, pages, year, month, doi
  - citekey = 파일명(스템)

주의: 출판사 BibTeX가 month를 빼는 경우가 있음 → month 없으면 경고 출력(수동/Claude 보강 필요).
실행:  python normalize_bib.py
"""
import re, glob, os
import merge_bib as M

MN = ["", "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]
NAME_FIX = {
    "Lee, Seung-Hyeon": "Lee, Seunghyeon", "Lee, Seunghyun": "Lee, Seunghyeon",
    "Song, Young-Keun": "Song, Youngkeun",
}

def fix_author(a):
    a = a.strip()
    if "," in a:                       # already "Last, First"
        sur, giv = a.split(",", 1)
        out = f"{sur.strip()}, {giv.strip()}"
    elif " " in a:                     # "First ... Last" -> "Last, First ..."
        p = a.split()
        out = f"{p[-1]}, {' '.join(p[:-1])}"
    else:                              # concatenated "SurnameGiven"
        out = a
        for i in range(1, len(a)):
            if a[i].isupper() and a[i - 1].islower():
                out = f"{a[:i]}, {a[i:]}"
                break
    return NAME_FIX.get(out, out)

def normalize(raw, key):
    raw = re.sub(r"(\bpages\s*=\s*\{[^}]*\})(\s*\n\s*year)", r"\1,\2", raw, flags=re.I)  # comma fix
    au = " and ".join(fix_author(a) for a in M.field(raw, "author").split(" and "))
    moN = M.month_num(raw)
    f = {
        "author": au, "title": M.field(raw, "title"), "journal": M.field(raw, "journal"),
        "volume": M.field(raw, "volume"), "number": M.field(raw, "number"),
        "pages": M.field(raw, "pages") or M.field(raw, "article-number") or M.field(raw, "eid"),
        "year": M.field(raw, "year"),
        "month": MN[moN] if 0 < moN < 13 else "", "doi": M.get_doi(raw),
    }
    lines = [f"@article{{{key},"]
    lines.append(f"  author  = {{{f['author']}}},")
    lines.append(f"  title   = {{{f['title']}}},")
    lines.append(f"  journal = {{{f['journal']}}},")
    vnp = []
    if f["volume"]: vnp.append(f"volume = {{{f['volume']}}}")
    if f["number"]: vnp.append(f"number = {{{f['number']}}}")
    if f["pages"]:  vnp.append(f"pages = {{{f['pages']}}}")
    if vnp: lines.append("  " + ", ".join(vnp) + ",")
    ym = [f"year = {{{f['year']}}}"]
    if f["month"]: ym.append(f"month = {{{f['month']}}}")
    lines.append("  " + ", ".join(ym) + ",")
    lines.append(f"  doi     = {{{f['doi']}}}")
    lines.append("}")
    return "\n".join(lines) + "\n", f

def main():
    no_month = []
    for path in sorted(glob.glob(os.path.join(os.path.dirname(__file__), "individual", "*.bib"))):
        key = os.path.splitext(os.path.basename(path))[0]
        ents = M.split_entries(open(path, encoding="utf-8").read())
        out = []
        for raw in ents:
            text, f = normalize(raw, key if len(ents) == 1 else (key + "_" + (f and "x")))
            out.append(text)
            if not f["month"]:
                no_month.append(key)
        open(path, "w", encoding="utf-8").write("\n".join(out))
        print("normalized", os.path.basename(path))
    if no_month:
        print("WARN month 없음(보강 필요):", ", ".join(no_month))

if __name__ == "__main__":
    main()
