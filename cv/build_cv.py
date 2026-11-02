# -*- coding: utf-8 -*-
"""
Academic CV builder — Gukhwa Jang
Forest remote sensing / LiDAR · GEDI / ecology — postdoc application

Structure follows field-standard academic CVs (Pascual UMD / Harvard GSAS).
All content is sourced from the verified portfolio data in
  ../src/data/site.json
Edit the data blocks below, then run:  python build_cv.py
Output: Jang_Gukhwa_CV_v1.docx  ->  PDF via make_pdf.ps1 -Publish (LibreOffice; never Word/docx2pdf) -> ../public/cv/
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, re, json
from datetime import date as _dtdate

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------- data source
# 사실 데이터의 정본은 application_info.xlsx 다. 이 파일에 값을 직접 적지 않는다.
#   - 원자 필드로 CV 표기가 재현되면 규칙으로 조합한다
#   - 재현되지 않는 curated 표기는 xlsx 의 'CV 표기(EN)' / 'CV 서술(EN)' 컬럼을 그대로 쓴다
#   - 수록 대상은 xlsx 의 'CV 표시' / 'CV 구분' 플래그로 고른다
import openpyxl as _oxl

_WB = _oxl.load_workbook(os.path.join(HERE, "application_info.xlsx"), data_only=True)


def sheet(name, where=None, order=None):
    """xlsx 시트를 dict 리스트로. where 로 거르고 order 로 정렬한다.
    엑셀의 행 순서에 의존하지 않도록 정렬 키를 명시하는 것을 원칙으로 한다."""
    ws = _WB[name]
    hdr = [str(c.value) for c in ws[1]]
    out = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    out = [r for r in out if any(v is not None for v in r.values())]
    if where:
        out = [r for r in out if where(r)]
    if order:
        out.sort(key=order)
    return out


def years_only(period):
    """'2022/09 — 2025' -> '2022 — 2025' (CV 는 장학 기간을 연도로만 표기)"""
    return re.sub(r"(\d{4})/\d{2}", r"\1", str(period or "")).strip()


def _dot(s):
    """제목 끝의 마침표는 표기 규칙이므로 렌더러가 붙인다 (xlsx 에는 넣지 않는다)."""
    return str(s or "").rstrip(". ") + "."


def _ym(v):
    """'2024/07/22' -> '2024.07'  (CV 는 경력 기간을 연.월로 표기)"""
    m = re.match(r"(\d{4})/(\d{2})", str(v or ""))
    return f"{m.group(1)}.{m.group(2)}" if m else ""


def _yr(v):
    """'2024/07/22' · '2024.07' -> '2024'. 학력·직위 기간은 연도만 (영문 학술 CV 관행)."""
    m = re.match(r"(\d{4})", str(v or ""))
    return m.group(1) if m else ""


def _year_range(start, end):
    """연도 범위. 같은 해면 한 번만('2024'), 종료가 없으면 Present."""
    a, b = _yr(start), (_yr(end) or "Present")
    return a if a == b else f"{a} — {b}"


def _mon_yyyy(v):
    """'2026/08' -> 'Aug 2026' (수여 월 같은 특정 시점만 Mon YYYY).
    'Expected 2026/08 ...' -> 'expected Aug 2026'."""
    s = str(v or "")
    m = re.search(r"(\d{4})[/.](\d{2})", s)
    if not m:
        return s
    mon = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][int(m.group(2))]
    t = f"{mon} {m.group(1)}"
    return f"expected {t}" if s.startswith("Expected") else t


def _period_years(v):
    """'2025/07 — 2025/11' -> '2025', '2023/04 — 2027/12' -> '2023 — 2027' (과제 기간도 연도만)."""
    head, _, tail = str(v or "").partition(" — ")
    return _year_range(head, tail) if tail else _yr(head)


def _work(kind):
    """Work & Internships 에서 CV 구분으로 골라 (직위, 소속, 기간, 서술, 시작일) 튜플로.
    마지막 '시작일' 은 표기에 쓰지 않고 연표(Position)의 시작월 산출에만 쓴다."""
    out = []
    for r in sheet("Work & Internships", where=lambda r: r["CV 구분"] == kind,
                   order=lambda r: _desc(r["Start Date"])):
        out.append((_txt(r["CV 직위(EN)"]), _txt(r["CV 소속(EN)"]),
                    _year_range(r["Start Date"], r["End Date"]), _txt(r["CV 서술(EN)"]),
                    _txt(r["Start Date"])))
    return out


# 문장형 표기에서 소문자로 내리면 안 되는 것: 약어 · 고유명사 · 상표
_KEEP_WORDS = ["LiDAR", "GEDI", "XGBoost", "UAV", "ALS", "MLS", "RGB", "DEM", "DTM", "CHM",
               "Sentinel-2", "Landsat", "PlanetScope", "Gyeonggi-do", "Korea", "Korean",
               "Japan", "China", "Seoul", "AGU", "ESA", "ICLEE", "JCK", "CES", "KOSERT"]
_KEEP = {re.sub(r"[^a-z0-9]", "", w.lower()): w for w in _KEEP_WORDS}

# 여러 단어로 된 고유명사 — 단어 단위 보호로는 잡히지 않아 변환 후 되돌린다
_KEEP_PHRASES = ["Gyeonggi-do Province", "Republic of Korea", "World Forestry Congress",
                 "Consumer Electronics Show", "Seoul National University"]


def sentence_case(title):
    """제목형 -> 문장형 (APA 7th). 첫 단어와 보호 목록만 대문자를 유지한다.
    토큰 전체를 먼저 보호 목록과 대조하므로 Gyeonggi-do / Sentinel-2 처럼
    하이픈이 든 고유명사가 쪼개지지 않는다."""
    def one(tok, first):
        m = re.match(r"^(\W*)(.*?)(\W*)$", tok, re.S)
        pre, mid, post = m.group(1), m.group(2), m.group(3)
        if not mid:
            return tok
        bare = re.sub(r"[^a-z0-9]", "", mid.lower())
        if bare in _KEEP:
            return pre + _KEEP[bare] + post
        out = []
        for i, part in enumerate(mid.split("-")):
            pb = re.sub(r"[^a-z0-9]", "", part.lower())
            if pb in _KEEP:
                out.append(_KEEP[pb])
            elif first and i == 0:
                out.append(part[:1].upper() + part[1:].lower())
            else:
                out.append(part.lower())
        return pre + "-".join(out) + post

    out = " ".join(one(t, i == 0) for i, t in enumerate(str(title or "").split(" ")))
    for ph in _KEEP_PHRASES:
        out = re.sub(re.escape(ph), ph, out, flags=re.I)
    return out


def _lower_first(s):
    """괄호 안 설명은 소문자로 시작한다. 단 약어·고유명사(GEDI 등)는 그대로 둔다."""
    s = str(s or "")
    head = re.sub(r"[^a-z0-9]", "", s.split(" ")[0].lower())
    return s if head in _KEEP else s[:1].lower() + s[1:]


def _desc(v):
    """문자열 날짜를 내림차순 정렬 키로. CV 는 모든 섹션이 최신순이다."""
    return tuple(-ord(c) for c in str(v or ""))


def _txt(v):
    """빈 셀(None)을 빈 문자열로. xlsx 가 아직 비어 있는 컬럼은 '빈 채로' 인쇄한다 —
    str(None) 이 'None' 으로 찍히는 것만 막고, 값을 지어내거나 행을 건너뛰지 않는다."""
    return "" if v is None else str(v).strip()

# ---------------------------------------------------------------- timeline collector
# 맨 아래 'Timeline (2023 — Present)' 섹션의 재료. 각 섹션의 '렌더 루프 안에서' _tl() 을
# 부르므로 연표 항목은 본문 항목과 1:1 로 대응한다 (xlsx 를 새 필터로 다시 읽지 않는다).
TL_FROM = 2023          # 수록 기준: 2023-01 이후에 '일어난' 확정 실적
TIMELINE = []


def _ym2(v):
    """날짜 셀 -> (year, month). xlsx 에 섞여 있는 표기를 모두 받는다.
      '2026/10/28' · '2025.09.17' · '2025.03' · '2025/06 — 2027/05' · datetime
    월을 읽을 수 없으면 month=0 (연표에서 '월 미상'으로 그 해 끝에 둔다)."""
    if hasattr(v, "year") and hasattr(v, "month"):
        return v.year, v.month
    s = str(v or "")
    m = re.match(r"\s*(\d{4})\s*[/.\-]\s*(\d{1,2})\b", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"\s*(\d{4})", s)
    return (int(m.group(1)), 0) if m else (0, 0)


def _tl(year, month, cat, main, meta=""):
    """연표에 한 줄 추가. TL_FROM 이전 항목과 연도 미상 항목은 조용히 빠진다."""
    y = int(year) if str(year).strip().isdigit() else 0
    if y < TL_FROM or not str(main).strip():
        return
    TIMELINE.append({"y": y, "m": int(month or 0), "cat": cat,
                     "main": str(main).strip().rstrip(". "), "meta": str(meta or "").strip()})

# ---------------------------------------------------------------- palette
INK    = RGBColor(0x1a, 0x1a, 0x1a)
ACCENT = RGBColor(0x1f, 0x4e, 0x3d)   # deep forest green
GREY   = RGBColor(0x44, 0x44, 0x44)
LIGHT  = RGBColor(0x6f, 0x6f, 0x6f)
BODYFONT = "Cambria"   # academic serif, high legibility at small sizes; clean via LibreOffice
HEADFONT = "Cambria"

doc = Document()

# page margins
for s in doc.sections:
    s.top_margin = Inches(0.75); s.bottom_margin = Inches(0.75)
    s.left_margin = Inches(0.7); s.right_margin = Inches(0.7)

RIGHT_TAB = Inches(7.1)  # within 7.1" text width

# Disable OpenType kerning/ligatures — Word applies them unevenly when exporting to PDF,
# producing the "some letters too tight, some too loose" spacing. Turning the feature off
# gives uniform metric spacing.
for _cs in doc.settings.element.iter(qn('w:compatSetting')):
    if _cs.get(qn('w:name')) == 'enableOpenTypeFeatures':
        _cs.set(qn('w:val'), '0')

# base style
st = doc.styles["Normal"]
st.font.name = BODYFONT
st.font.size = Pt(10)
st.font.color.rgb = INK
st.paragraph_format.space_after = Pt(0)
st.paragraph_format.line_spacing = 1.13

# ---- font-size scale (single source of truth for the v8 readability pass) ----
# Maps every legacy point size used in the call sites below to its new value, so the
# whole document rescales consistently without touching ~100 individual calls.
#   8/8.5 = notes/meta · 9/9.5/10 = body & titles · 10.5 = "CV" tag · 11 = section head · 22 = name
_SZ = {8: 8, 8.5: 9, 9: 10, 9.5: 10.5, 10: 10.5, 10.5: 11, 11: 13.5, 12: 12, 22: 24}

def _force_font(rPr, font):
    """Set the font for ALL script categories (ascii/hAnsi/cs/eastAsia) so Word never
    falls back to the theme font for punctuation (·, —, “”, †, →) — that fallback splits
    the text across font subsets and produces uneven letter/word spacing in the PDF."""
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rPr.insert(0, rf)
    for _a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(_a), font)

def _set_run(r, size=9.5, bold=False, italic=False, color=INK, font=BODYFONT, caps=False, spacing=None):
    r.font.size = Pt(_SZ.get(size, size)); r.bold = bold; r.italic = italic
    r.font.color.rgb = color
    if caps: r.font.all_caps = True
    rPr = r._element.get_or_add_rPr()
    _force_font(rPr, font)
    if spacing is not None:
        sp = OxmlElement('w:spacing'); sp.set(qn('w:val'), str(spacing)); rPr.append(sp)

def add_hyperlink(paragraph, url, text, color="444444", size=9,
                  underline=False, italic=False, bold=False, font=BODYFONT):
    """Append a real, clickable hyperlink run to a paragraph (works in Word and exported PDF)."""
    r_id = paragraph.part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    link = OxmlElement('w:hyperlink'); link.set(qn('r:id'), r_id)
    run = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(a), font)
    rPr.append(rf)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(_SZ.get(size, size) * 2))); rPr.append(sz)
    col = OxmlElement('w:color'); col.set(qn('w:val'), color); rPr.append(col)
    if underline:
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    if italic: rPr.append(OxmlElement('w:i'))
    if bold:   rPr.append(OxmlElement('w:b'))
    run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    run.append(t); link.append(run); paragraph._p.append(link)
    return link

def add_right_tab(p):
    p.paragraph_format.tab_stops.add_tab_stop(RIGHT_TAB, WD_TAB_ALIGNMENT.RIGHT)

def bottom_border(p):
    pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '6')
    b.set(qn('w:space'), '2'); b.set(qn('w:color'), '1f4e3d')
    pbdr.append(b); pPr.append(pbdr)

# ---------------------------------------------------------------- running header / footer (every page)
_sec = doc.sections[0]
_hp = _sec.header.paragraphs[0]
_hp.paragraph_format.tab_stops.add_tab_stop(RIGHT_TAB, WD_TAB_ALIGNMENT.RIGHT)
_set_run(_hp.add_run("\tGukhwa Jang  —  Curriculum Vitae"), 8, color=LIGHT)
_fp = _sec.footer.paragraphs[0]
_fp.paragraph_format.tab_stops.add_tab_stop(RIGHT_TAB, WD_TAB_ALIGNMENT.RIGHT)
_set_run(_fp.add_run("\tGukhwa Jang    "), 8, color=LIGHT)
_pg = _fp.add_run()
for _t, _v in (("begin", None), ("instr", " PAGE "), ("end", None)):
    if _t == "instr":
        e = OxmlElement('w:instrText'); e.set(qn('xml:space'), 'preserve'); e.text = _v
    else:
        e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), _t)
    _pg._r.append(e)
_set_run(_pg, 8, color=LIGHT)

# ---------------------------------------------------------------- header
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p.add_run("Gukhwa Jang"); _set_run(r, 22, bold=True, color=INK)
r = p.add_run("    ·    Curriculum Vitae"); _set_run(r, 10.5, color=LIGHT, caps=True, spacing=20)
p.paragraph_format.space_after = Pt(1)

p = doc.add_paragraph()
_phd = next(r for r in sheet("Education") if str(r["Degree(EN)"]).startswith("Ph.D"))
r = p.add_run(f"Ph.D., {_mon_yyyy(_phd['Date Degree Received'])}")
_set_run(r, 10, color=ACCENT, bold=True)
p.paragraph_format.space_after = Pt(1)

p = doc.add_paragraph()
# 소속은 Education 의 Ph.D. 행에서 만든다 (하드코딩 금지 — 2절). 빈 칸은 조용히 빠진다.
_aff = ", ".join(x for x in (str(_phd[_k] or "").strip() for _k in
                             ("연구실(EN)", "Institution(EN)", "City(EN)", "Country(EN)")) if x)
if _aff:
    r = p.add_run(_aff)
    _set_run(r, 9.5, color=GREY)
p.paragraph_format.space_after = Pt(2)

contact = [
    ("rnrghk6811@snu.ac.kr", "mailto:rnrghk6811@snu.ac.kr"),
    ("ORCID 0000-0001-6653-9285", "https://orcid.org/0000-0001-6653-9285"),
    ("Google Scholar", "https://scholar.google.com/citations?user=dghWqqcAAAAJ"),
    ("gugi516.github.io", "https://gugi516.github.io/en"),
]
p = doc.add_paragraph()
for i, (disp, url) in enumerate(contact):
    if i:
        sep = p.add_run("  |  "); _set_run(sep, 9, color=LIGHT)
    if url:
        add_hyperlink(p, url, disp, color="444444", size=9)
    else:
        _set_run(p.add_run(disp), 9, color=GREY)
bottom_border(p); p.paragraph_format.space_after = Pt(6)

# ---------------------------------------------------------------- section helper
def section(title):
    p = doc.add_paragraph()
    r = p.add_run(title); _set_run(r, 11, bold=True, color=ACCENT, caps=True, spacing=30)
    bottom_border(p)
    p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(6)

def entry(left_bold, right="", sub="", bullets=None, left_rest=""):
    """One CV entry: bold lead + optional right-aligned date, italic subline, bullets."""
    p = doc.add_paragraph(); add_right_tab(p)
    r = p.add_run(left_bold); _set_run(r, 9.5, bold=True)
    if left_rest:
        r = p.add_run(left_rest); _set_run(r, 9.5)
    if right:
        r = p.add_run("\t" + right); _set_run(r, 9, color=GREY)
    p.paragraph_format.space_before = Pt(5)
    if sub:
        ps = doc.add_paragraph(); r = ps.add_run(sub); _set_run(r, 9, italic=True, color=GREY)
    if bullets:
        for b in bullets:
            pb = doc.add_paragraph(style="List Bullet")
            pb.paragraph_format.left_indent = Inches(0.25)
            pb.paragraph_format.space_after = Pt(0)
            r = pb.add_run(b); _set_run(r, 9)

def authors_run(p, authors, me="Gukhwa Jang", mark_first=False):
    """Render byline; bold `me`. If mark_first and `me` is the first author, add † after the name."""
    for i, a in enumerate(authors):
        if i:
            s = p.add_run(", "); _set_run(s, 9, color=INK)
        r = p.add_run(a); _set_run(r, 9, bold=(a == me))
        if mark_first and i == 0 and a == me:
            d = p.add_run("†"); _set_run(d, 9, bold=True, color=ACCENT)

# ---------------------------------------------------------------- Research Interests
section("Research Interests")
# 첫 항목만 대문자로 시작하고 나머지는 이어쓰기 (한 문장으로 읽히게)
_ri = [x for x in (_txt(r["Interest(EN)"])
                   for r in sheet("Research Interests", order=lambda r: r["No."])) if x]
p = doc.add_paragraph()
r = p.add_run(" · ".join([_ri[0]] + [_lower_first(x) for x in _ri[1:]]) + "." if _ri else "")
_set_run(r, 9.5)
p.paragraph_format.space_after = Pt(2)

# ---------------------------------------------------------------- Education
section("Education")
def _edu_period(r):
    """학력 기간은 연도만('2022 — 2026') — 영문 학술 CV 관행. xlsx 의 YYYY.MM 은 렌더러가 연도로 줄인다.
    수여 예정이면 '(expected)' 를 덧붙인다."""
    head, _, tail = str(r["Start–End"] or "").partition(" — ")
    se = _year_range(head, tail)
    if str(r["Date Degree Received"] or "").startswith("Expected"):
        return f"{se} (expected)"
    return se


def _edu_bullets(r):
    """학위논문 + (있으면) 심사위원. 박사는 Dissertation, 그 외는 Thesis."""
    out = []
    if r["학위논문(EN)"]:
        label = "Dissertation" if "Doctoral" in str(r["Level of Study"]) else "Thesis"
        out.append(f'{label}: “{r["학위논문(EN)"]}”')
    if r["심사위원(EN)"]:
        out.append(f'Committee: {r["심사위원(EN)"]}')
    return out or None


_MONTHS = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
           7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}

edu = [(", ".join(x for x in (_txt(r["Degree(EN)"]),
                              _txt(r["Program(EN)"]).split(" · ")[0]) if x),
        _edu_period(r),
        # 기관명을 부제 줄에 함께 — 학위마다 대학이 다르므로(SNU/KNU) 생략하면 구분이 안 된다 (2026-10-21)
        " · ".join(x for x in (_txt(r["Institution(EN)"]), _txt(r["CV 서술(EN)"])) if x),
        _edu_bullets(r))
       for r in sheet("Education", where=lambda r: r["CV 표시"] == "Y",
                      order=lambda r: r["No."])]
for lead, date, sub, bl in edu:
    entry(lead, right=date, sub=sub, bullets=bl)

# ---------------------------------------------------------------- Research Experience
section("Research Experience")
exp = _work("research")
for lead, rest, date, sub, _start in exp:
    entry(lead, right=date, sub=sub, left_rest=rest)
    _y, _m = _ym2(_start)
    _tl(_y, _m, "Position", lead, rest)

# ---------------------------------------------------------------- Publications
section("Peer-Reviewed Publications")
p = doc.add_paragraph()
r = p.add_run("APA 7th.  Name in bold = author; † = first author.  "
              "Impact Factor / quartile shown where available.  (SCIE / ESCI, then KCI; newest first)  "
              "Korean-journal (KCI) papers are listed in Korean.")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)

# --- source: citations/publications.bib (ascending) + IF/quartile from site.json (matched by DOI) ---
def _bib_entries(path):
    txt = open(path, encoding="utf-8").read()
    out, i = [], 0
    while True:
        at = txt.find("@", i)
        if at < 0: break
        b = txt.find("{", at); depth, j = 0, b
        while j < len(txt):
            if txt[j] == "{": depth += 1
            elif txt[j] == "}":
                depth -= 1
                if depth == 0: break
            j += 1
        out.append(txt[at:j + 1]); i = j + 1
    return out

def _bf(raw, name):
    m = re.search(r"\b" + name + r"\s*=\s*\{(.*?)\}\s*[,}\n]", raw, re.I | re.S)
    return m.group(1).strip() if m else ""

def _initials(given):
    return " ".join("-".join(s[0].upper() + "." for s in part.split("-") if s)
                    for part in given.split())

_SURNAME, _GIVEN = "Jang", "Gukhwa"

def _is_me(sur, giv):
    """본인 저자 판별. 두 표기를 다 인정한다 —
    풀네임('Jang, Gukhwa' · 'Gukhwa Jang', Publications 시트)과
    축약형('Jang, G.', Under Review 시트). 축약형은 given 이 이니셜 한 글자일 때만
    인정한다 — 'Jang, Gilsu' 같은 동성 타인을 굵게 하지 않기 위함이다."""
    if sur != _SURNAME:
        return False
    g = giv.strip().rstrip(".")
    if len(g) == 1:
        return g.upper() == _GIVEN[0]
    return g.startswith(_GIVEN)

def _apa_author(name):
    if "," in name:
        sur, giv = [x.strip() for x in name.split(",", 1)]
    else:
        ps = name.split(); sur, giv = ps[-1], " ".join(ps[:-1])
    return f"{sur}, {_initials(giv)}".strip().rstrip(","), _is_me(sur, giv)

# --- [authors-parse] ---  (_tests/test_parse_authors.py 가 이 블록만 떼어내 테스트한다)
_INIT_PIECE = re.compile(r"(?:[A-Z]\.)+|[A-Z]")

def _is_initials(tok):
    """토큰이 '이니셜만'인가 — 'G.' 'G' 'J. S.' 'S.-H.' 'G.H.' 는 참,
    'Lee' 'LEE' 'Gukhwa' 는 거짓(약어 조각이 아니라 이름이므로)."""
    parts = [p for p in re.split(r"[\s-]+", tok.strip()) if p]
    return bool(parts) and all(_INIT_PIECE.fullmatch(p) for p in parts)

_AND_SEP = re.compile(r"\s*&\s*|\s+and\s+", re.I)

def _parse_authors(s):
    """저자 문자열을 저자 리스트로 분해한다. 두 표기를 다 받는다.

      APA 축약형  "Jang, G., Kim, S., & Lee, J."  ->  ["Jang, G.", "Kim, S.", "Lee, J."]
      풀네임 나열  "Gukhwa Jang, Saehoon Kim"      ->  ["Gukhwa Jang", "Saehoon Kim"]

    APA 에서 쉼표는 '성, 이니셜' 의 구분자이면서 저자 사이 구분자이기도 하다. 그래서
    쉼표로 쪼갠 뒤 '이니셜만인 토큰'을 바로 앞 토큰에 되붙인다. `&`·`and` 는 쉼표와
    같이 취급하고 빈 토큰은 버린다. 이름을 만들어 내지 않는다 — 입력 토큰의 재결합뿐이다."""
    toks = [t.strip() for t in _AND_SEP.sub(",", _txt(s)).split(",")]
    out = []
    for t in (t for t in toks if t):
        if out and _is_initials(t) and not _is_initials(out[-1]):
            out[-1] = f"{out[-1]}, {t}"
        else:
            out.append(t)
    return out
# --- [/authors-parse] ---

_site = json.load(open(os.path.join(HERE, "..", "src", "data", "site.json"), encoding="utf-8"))
IFQ = {(it.get("doi") or "").lower().strip():
           (it.get("if", ""), it.get("quartile", ""), it.get("index", ""))
       for it in _site["publications"]["items"] if it.get("doi")}

# KCI(국내) 논문은 CV 에 국문 서지로 표기한다 (2026-10-21 사용자 지시).
# DOI 로 site.json 항목을 찾고, DOI 없는 게재확정(in press) 건은 영문 제목으로 찾는다.
_PUB_BY_DOI = {(it.get("doi") or "").lower().strip(): it
               for it in _site["publications"]["items"] if it.get("doi")}
_PUB_BY_TITLE = {((it.get("title") or {}).get("en") or "").lower().strip(): it
                 for it in _site["publications"]["items"] if (it.get("title") or {}).get("en")}

_MO = {m: i for i, m in enumerate(
    ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"], 1)}
def _monthnum(raw):
    m = _bf(raw, "month").lower()[:3]
    return _MO.get(m, int(m) if m.isdigit() else 0)

_pubs = []
for raw in _bib_entries(os.path.join(HERE, "citations", "publications.bib")):
    doi = _bf(raw, "doi")
    iff, q, idx = IFQ.get(doi.lower(), ("", "", ""))
    yv = _bf(raw, "year"); yv = int(yv) if yv.isdigit() else 0
    _pg = re.search(r"\d+", _bf(raw, "pages")); sp = int(_pg.group()) if _pg else 0
    _it = _PUB_BY_DOI.get(doi.lower().strip()) or _PUB_BY_TITLE.get(_bf(raw, "title").lower().strip())
    _pubs.append({
        "ko": ({"authors": _it.get("authorsKo"), "title": (_it.get("title") or {}).get("ko"),
                "journal": _it.get("journalKo")} if _it else None),
        "apa": [_apa_author(a.strip()) for a in _bf(raw, "author").split(" and ")],
        "year": _bf(raw, "year"),
        # 연표의 월. bib 에 month 가 없는 게재확정(in press) 건은 0 -> 그 해 끝에 월 없이 놓인다.
        "mon": _monthnum(raw),
        "title": re.sub(r"\s*-\s*$", "", _bf(raw, "title")).replace(" - ", ": ").rstrip(". "),
        "journal": _bf(raw, "journal").replace("\\&", "&"),
        "vol": _bf(raw, "volume"), "num": _bf(raw, "number"),
        "pages": _bf(raw, "pages").replace("-", "–"), "doi": doi, "iff": iff, "q": q, "idx": idx,
        # 국제(SCIE/ESCI)→KCI, newest first, then start page.
        # 등재구분(index)이 있으면 그것으로, 없으면 IF/quartile 유무로 판정한다.
        # 'ESCI / Scopus' 같은 복합 표기도 국제로 (2026-11-01 JAU 오분류 수정)
        "sortk": (0 if any(idx.strip().upper().startswith(k) for k in ("SCIE", "ESCI"))
                  else (0 if (iff or q) else 1),
                  -yv, -_monthnum(raw), sp),
    })
_pubs.sort(key=lambda d: d["sortk"])

for n, d in enumerate(_pubs, 1):
    apa, year, title, journal = d["apa"], d["year"], d["title"], d["journal"]
    vol, num, pages, doi, iff, q = d["vol"], d["num"], d["pages"], d["doi"], d["iff"], d["q"]
    idx = d["idx"]
    # KCI(국내) 그룹이고 국문 서지가 있으면 저자·제목·학술지를 국문으로 쓴다 (2026-10-21)
    # 등재구분이 국제(ESCI 등)인 항목은 그룹과 무관하게 영문 유지 (2026-11-01)
    _ko = d["ko"]
    _idx_u = (idx or "").strip().upper()
    _use_ko = (d["sortk"][0] == 1 and (not _idx_u or _idx_u.startswith("KCI"))
               and bool(_ko and _ko["title"] and _ko["authors"]))
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.left_indent = Inches(0.28); p.paragraph_format.first_line_indent = Inches(-0.28)
    _set_run(p.add_run(f"{n}. "), 9, color=GREY)
    if _use_ko:
        for i, _nm in enumerate(_ko["authors"]):
            if i > 0:
                _set_run(p.add_run("·"), 9)
            _me_ko = _nm.strip() == "장국화"
            _set_run(p.add_run(_nm), 9, bold=_me_ko)
            if i == 0 and _me_ko:
                _set_run(p.add_run("†"), 9, bold=True, color=ACCENT)
        _set_run(p.add_run(f" ({year}). " if year else " "), 9)
        _set_run(p.add_run(_ko["title"] + ". "), 9)
        _set_run(p.add_run(_ko["journal"] or journal), 9, italic=True)
    else:
        nA = len(apa)
        for i, (disp, is_user) in enumerate(apa):
            if i > 0:
                _set_run(p.add_run(", "), 9)
            if i == nA - 1 and nA > 1:
                _set_run(p.add_run("& "), 9)
            _set_run(p.add_run(disp), 9, bold=is_user)
            if i == 0 and is_user:
                _set_run(p.add_run("†"), 9, bold=True, color=ACCENT)
        _set_run(p.add_run(f" ({year}). "), 9)
        _set_run(p.add_run(title + ". "), 9)
        _set_run(p.add_run(journal), 9, italic=True)
    if vol:
        _set_run(p.add_run(", "), 9)
        _set_run(p.add_run(vol), 9, italic=True)
        if num:
            _set_run(p.add_run(f"({num})"), 9)
        if pages:
            _set_run(p.add_run(f", {pages}"), 9)
    _set_run(p.add_run(". "), 9)
    if doi:
        add_hyperlink(p, f"https://doi.org/{doi}", f"https://doi.org/{doi}", color="444444", size=9)
    # 등재구분 태그(SCIE/ESCI/KCI) 먼저, 그 뒤에 기존 IF·quartile 태그.
    if idx:
        _set_run(p.add_run(f"  · {idx}"), 8.5, color=ACCENT)
    if iff or q:
        iv = iff.split("(")[0].strip()
        tag = "  · " + ", ".join(x for x in [f"IF {iv}" if iv else "", q] if x)
        _set_run(p.add_run(tag), 8.5, color=ACCENT)
    # 연표 — 본문과 같은 언어 선택(KCI 는 국문)을 그대로 따른다
    _tl(d["year"], d["mon"], "Publication",
        _ko["title"] if _use_ko else title,
        (_ko["journal"] or journal) if _use_ko else journal)

# ---------------------------------------------------------------- Manuscripts under review
section("Manuscripts Under Review")
# status = "stage|date"  (stage shown in brackets, date in grey)
# 심사중 원고: 저자는 한 문자열이므로 리스트로 되돌린다 — APA 축약형("Jang, G., Kim, S., & Lee, J.")과
# 풀네임 나열("Gukhwa Jang, Saehoon Kim") 둘 다 _parse_authors() 가 받는다.
under_review = [(_parse_authors(r["Authors(EN)"]),
                 _dot(r["Title(EN)"]), _txt(r["Journal(EN)"]),
                 f'{_txt(r["Status"])}|{_mon_yyyy(r["Status Date"])}')
                for r in sheet("Under Review", order=lambda r: r["No."])]
# 안내문의 건수는 'Under Review' 시트 행수에서 산출한다 (하드코딩 금지).
n_under_review = len(under_review)
p = doc.add_paragraph()
r = p.add_run("APA 7th.  Name in bold = author; † = first author.  "
              f"({n_under_review} manuscript{'' if n_under_review == 1 else 's'}; "
              "first-author listed first)")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)
for n, (auth, title, jour, status) in enumerate(under_review, 1):
    apa = [_apa_author(a) for a in auth]
    stage, _, dt = status.partition("|")
    # Status Date 가 비면 연도를 비워 둔다 — 없는 연도를 채워 넣지 않는다.
    yr = dt[:4] if dt[:4].isdigit() else ""
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.left_indent = Inches(0.28); p.paragraph_format.first_line_indent = Inches(-0.28)
    _set_run(p.add_run(f"{n}. "), 9, color=GREY)
    nA = len(apa)
    for i, (disp, is_user) in enumerate(apa):
        if i > 0:
            _set_run(p.add_run(", "), 9)
        if i == nA - 1 and nA > 1:
            _set_run(p.add_run("& "), 9)
        _set_run(p.add_run(disp), 9, bold=is_user)
        if i == 0 and is_user:
            _set_run(p.add_run("†"), 9, bold=True, color=ACCENT)
    _set_run(p.add_run(f" ({yr}). " if yr else " "), 9)
    _set_run(p.add_run(title.rstrip(". ")), 9, italic=True)       # APA: unpublished title in italics
    _set_run(p.add_run(f" [{stage}]. " if stage else " "), 9)
    _set_run(p.add_run(jour + "." if jour else ""), 9)
    if dt:
        _set_run(p.add_run(f"  · {dt}"), 8.5, color=GREY)

# ---------------------------------------------------------------- Books
section("Books")
# 저자 문자열 안에서 본인 이름만 굵게 — 이름 위치로 잘라 세 조각으로 출력한다
_ME = "Gukhwa Jang"
for _n, _b in enumerate(sheet("Books", order=lambda r: r["No."]), 1):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.left_indent = Inches(0.28); p.paragraph_format.first_line_indent = Inches(-0.28)
    _set_run(p.add_run(f"{_n}. "), 9, color=GREY)
    _auth = str(_b["Authors(EN)"] or "")
    _head, _sep, _tail = _auth.partition(_ME)
    _set_run(p.add_run(_head), 9)
    if _sep:
        _set_run(p.add_run(_sep), 9, bold=True)
    _yr_b = _txt(_b["Year"])
    _set_run(p.add_run(f'{_tail} ({_yr_b}). ' if _yr_b else f'{_tail} '), 9)
    _set_run(p.add_run(_txt(_b["Title(EN)"])), 9, italic=True)
    _pub_b, _isbn_b = _txt(_b["Publisher(EN)"]), _txt(_b["ISBN"])
    _set_run(p.add_run(". " + (f'{_pub_b}. ' if _pub_b else "")
                       + (f'ISBN {_isbn_b}.' if _isbn_b else "")), 9)
    # 단행본에는 월 정보가 없다 -> 월 미상(0)
    _tl(_yr_b, 0, "Book", _txt(_b["Title(EN)"]), _pub_b)

# ---------------------------------------------------------------- Fellowships & Funding
section("Fellowships & Funding")
p = doc.add_paragraph()
r = p.add_run("Competitive fellowships / scholarships (amounts converted at ₩1,400 = $1).")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)
funding = [(_txt(r["CV 표기(EN)"]), years_only(r["Period"]), _txt(r["Amount(USD)"]),
            _txt(r["Period"]))
           for r in sheet("Funding", order=lambda r: r["No."])]
for name, period, amount, _raw_period in funding:
    p = doc.add_paragraph(); add_right_tab(p); p.paragraph_format.space_before = Pt(4)
    _set_run(p.add_run(name), 9)
    _set_run(p.add_run(f"  ({amount})" if amount else ""), 9, color=ACCENT)
    _set_run(p.add_run("\t" + period), 9, color=GREY)
    # 연표는 수여/개시 시점 — 기간 문자열의 앞쪽(시작 연월)을 쓴다
    _y, _m = _ym2(_raw_period)
    _tl(_y, _m, "Funding", name)

# ---------------------------------------------------------------- Invited talks
section("Invited Talks")
# 초청강연의 정본은 'Invited Talks' 시트 하나다 (교과목 초청강의 · 학회 초청세션 모두 포함).
_INVITED_TYPE = {"guest lecture": "Guest lecture", "conference": "Invited talk",
                 "seminar": "Invited seminar", "colloquium": "Colloquium",
                 "keynote": "Keynote", "panel": "Panel"}


def _invited_order(r):
    """Date 내림차순. Date 가 비면 Year 내림차순 · No. 순 (연도만 아는 항목이 그 해 앞)."""
    d = str(r["Date"] or "").strip()
    return (-int(_yr(r["Year"]) or 0), 0 if not d else 1, _desc(d), r["No."])


def _invited_tail(r):
    """italic 꼬리 = 'TypeLabel, Host(EN), Venue(EN)'. 빈 값은 건너뛴다.
    guest lecture 는 Host 에 대학명이 이미 들어가므로 Venue 를 생략한다."""
    kind = str(r["Type"] or "").strip().lower()
    parts = [_INVITED_TYPE.get(kind, str(r["Type"] or "").strip()),
             str(r["Host(EN)"] or "").strip()]
    if kind != "guest lecture":
        parts.append(str(r["Venue(EN)"] or "").strip())
    return ", ".join(x for x in parts if x)


invited = [(_dot(r["Title(EN)"]) if _txt(r["Title(EN)"]) else "",
            _invited_tail(r), _txt(r["Year"]), _txt(r["Date"]))
           for r in sheet("Invited Talks", order=_invited_order)]
for title, venue, yr, _dt in invited:
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.left_indent = Inches(0.28); p.paragraph_format.first_line_indent = Inches(-0.28)
    b = p.add_run("• "); _set_run(b, 9, color=ACCENT)
    r = p.add_run("Gukhwa Jang"); _set_run(r, 9, bold=True)
    r = p.add_run(f" ({yr}). " if yr else " "); _set_run(r, 9)
    r = p.add_run(title + " " if title else ""); _set_run(r, 9)
    r = p.add_run(venue + "." if venue else ""); _set_run(r, 9, italic=True, color=GREY)
    _y, _m = _ym2(_dt)
    _tl(_y or yr, _m, "Invited Talk", title or venue, venue if title else "")

# ---------------------------------------------------------------- Conferences
section("Conference Presentations (International)")
p = doc.add_paragraph()
r = p.add_run("First author / presenter unless marked “team”.  "
              "10 domestic presentations (KOSERT, KSEE, KILA) available on request.")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)

# 국제학회 발표만 CV 에 싣는다 (국내 10건은 요약 문구로 대체).
# 제목은 실제 초록 제출 제목을 저장하고, 표기만 APA 문장형으로 변환한다.
confs = [([a for a in (x.strip() for x in _txt(r["Authors"]).split(",")) if a],
          _txt(r["Year"]),
          _txt(r["Type"]).capitalize() + (", team" if r["팀 발표"] == "Y" else ""),
          _dot(sentence_case(r["Title(EN)"])) if _txt(r["Title(EN)"]) else "",
          ", ".join(x for x in (_txt(r["CV 학회(EN)"]), _txt(r["Venue(EN)"])) if x),
          _txt(r["Date"]))
         for r in sheet("Conferences",
                        where=lambda r: r["Scope"] == "international",
                        order=lambda r: _desc(r["Date"]))]
for auth, yr, kind, title, venue, _dt in confs:
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.first_line_indent = Inches(-0.28)
    bullet = p.add_run("• "); _set_run(bullet, 9, color=ACCENT)
    authors_run(p, auth)
    yr_r = p.add_run(f" ({yr}). " if yr else " "); _set_run(yr_r, 9)
    t = p.add_run(title + " " if title else ""); _set_run(t, 9)
    v = p.add_run(((f"[{kind}] " if kind else "") + (venue + "." if venue else "")))
    _set_run(v, 9, italic=True, color=GREY)
    _y, _m = _ym2(_dt)
    _tl(_y or yr, _m, "Conference", title or venue, venue if title else "")

# ---------------------------------------------------------------- Grants / Projects
section("Research Projects")
p = doc.add_paragraph()
r = p.add_run("National R&D and municipal projects (newest first).  "
              "Budget = total project funding (national R&D).")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)

# 과제는 CV 표기(다듬은 과제명 · 기여 불릿 · 성과 요약)를 쓴다.
# xlsx 의 Title(EN)/기여(EN)/논문(EN)/특허(EN) 은 지원서 서식용 원문이라 별도로 둔다.
projects = [(_txt(r["CV 과제명(EN)"]), _txt(r["Institute(EN)"]),
             _period_years(r["기간"]), _txt(r["Role(EN)"]), _txt(r["PI(EN)"]),
             _txt(r["예산(USD)"]),
             str(r["CV 기여(EN)"]).splitlines() if r["CV 기여(EN)"] else None,
             r["CV 성과(EN)"], _txt(r["기간"]))
            for r in sheet("Projects", order=lambda r: r["No."])]
for title, funder, period, role, pi, budget, bullets, outputs, _raw_per in projects:
    # 제목이 매우 길어 우측탭 기간이 잘리던 문제 → 제목은 전체 폭, 기간은 상세줄 맨 앞으로 이동
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    _set_run(p.add_run(title), 9.5, bold=True)
    # 상세줄은 채워진 조각만 ' · ' 로 잇는다 (빈 칸이면 구분자·라벨째로 빠진다)
    ps = doc.add_paragraph()
    _set_run(ps.add_run(period), 9, bold=True, color=GREY)
    if period and (funder or role or pi):
        _set_run(ps.add_run(" · "), 9, color=GREY)
    if funder:
        _set_run(ps.add_run(funder + (" · " if (role or pi) else "")), 9, italic=True, color=GREY)
    if role:
        _set_run(ps.add_run(f"Role: {role}"), 9, italic=True, bold=True, color=ACCENT)
    if pi:
        _set_run(ps.add_run((" · " if role else "") + f"PI: {pi}"), 9, italic=True, color=GREY)
    if budget:
        _set_run(ps.add_run(f" · {budget} (total project)"), 9, italic=True, color=GREY)
    for b in (bullets or []):
        pb = doc.add_paragraph(style="List Bullet")
        pb.paragraph_format.left_indent = Inches(0.25); pb.paragraph_format.space_after = Pt(0)
        _set_run(pb.add_run(b), 9)
    if outputs:
        po = doc.add_paragraph(); po.paragraph_format.left_indent = Inches(0.25)
        po.paragraph_format.space_before = Pt(1)
        _set_run(po.add_run("→ Related outputs:  "), 8.5, bold=True, color=ACCENT)
        _set_run(po.add_run(outputs), 8.5, italic=True, color=GREY)
    # 연표는 과제 '시작' 시점. 본인 책임과제는 (PI) 를 붙인다 (Role(EN) 기준).
    _y, _m = _ym2(_raw_per)
    _tl(_y, _m, "Project",
        title + (" (PI)" if role == "Principal Investigator" else ""), funder)

# ---------------------------------------------------------------- Patents
section("Patents (Republic of Korea)")
p = doc.add_paragraph()
# xlsx Status 는 한글('등록'/'출원')일 수 있다 - _sync_site 와 같은 매핑 (2026-10-21)
def _pstat(v):
    s = str(v or "").strip().lower()
    if s.startswith("등록"):   # '등록', '등록 (디자인)' 모두
        return "registered"
    if s.startswith("출원"):
        return "application"
    return s
_n_reg = len(sheet("Patents", where=lambda r: _pstat(r["Status"]) == "registered"))
_n_pend = len(sheet("Patents", where=lambda r: _pstat(r["Status"]) == "application"))
r = p.add_run(f"{_n_reg} registered, {_n_pend} pending.")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)

def _patent_sub(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(5)
    r = p.add_run(text); _set_run(r, 9.5, bold=True, color=ACCENT)

def _patent_list(items):
    for title, meta in items:
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.left_indent = Inches(0.28)
        p.paragraph_format.first_line_indent = Inches(-0.28)
        bullet = p.add_run("• "); _set_run(bullet, 9, color=ACCENT)
        t = p.add_run(title + " " if title else ""); _set_run(t, 9)
        m = p.add_run(meta); _set_run(m, 8.5, italic=True, color=GREY)

def _patents(status, datecol, fmt):
    """등록/출원 목록. 상세줄은 상태 + 날짜 + 권리자 규칙으로 만든다.
    제목은 공식 등록명이 아니라 CV 용으로 다듬은 표기를 쓴다."""
    out = []
    for r in sheet("Patents", where=lambda r: _pstat(r["Status"]) == status,
                   order=lambda r: _desc(r[datecol])):
        raw = str(r[datecol] or "")
        when = _mon_yyyy(raw) if fmt == "Registered" else _yr(raw)   # 등록=시점(Mon YYYY), 출원중=연도
        # CV 전용 표기가 비어 있으면 영문 명칭 -> 국문 명칭 순으로 대체한다 (2026-10-21)
        _pt = _txt(r["CV 특허명(EN)"]) or _txt(r["Title(EN)"]) or _txt(r["특허명(KO)"])
        _po = _txt(r["CV 권리자(EN)"]) or _txt(r["Owner(EN)"]) or _txt(r["특허권자(KO)"])
        out.append((_pt,
                    " · ".join(x for x in (" ".join(y for y in (fmt, when) if y),
                                           _po) if x)))
        # 연표 날짜 = 등록건은 등록일, 출원중은 출원일 (datecol 이 그 구분을 담고 있다)
        _y, _m = _ym2(raw)
        _tl(_y, _m, "Patent", _pt, " · ".join(x for x in (fmt, _po) if x))
    return out


_patent_sub("Registered")
_patent_list(_patents("registered", "Reg Date", "Registered"))
_patent_sub("Pending")
_patent_list(_patents("application", "App Date", "Pending"))

# ---------------------------------------------------------------- Honors
section("Awards & Honors")

def _subhead(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    r = p.add_run(text); _set_run(r, 9.5, bold=True, color=ACCENT)

def _tabbed(rows):
    for h, d in rows:
        p = doc.add_paragraph(); add_right_tab(p); p.paragraph_format.space_before = Pt(4)
        r = p.add_run(h); _set_run(r, 9)
        r = p.add_run("\t" + d); _set_run(r, 9, color=GREY)

def _award_line(r):
    """'Award — Organization (details)'. 빈 칸은 구분자·괄호째로 빠진다."""
    head = " — ".join(x for x in (_txt(r["Award(EN)"]), _txt(r["Organization(EN)"])) if x)
    det = _lower_first(_txt(r["Details(EN)"]))
    return head + (f" ({det})" if det else "")


_subhead("Awards")
_awards = sheet("Awards", order=lambda r: _desc(r["Year"]))
_tabbed([(_award_line(r), _txt(r["Year"])) for r in _awards])
for _a in _awards:
    # 수상은 연도만 기록돼 있다 -> 월 미상(0)
    _tl(_yr(_a["Year"]), 0, "Award", _award_line(_a))

# Service & Membership 섹션은 2026-10-21 사용자 지시로 제거 (데이터도 해당 유형 없음)

# ---------------------------------------------------------------- Teaching
section("Teaching Experience")
_tw = _site["teaching"]["items"]
_n_univ   = sum(1 for _t in _tw if _t["type"] == "univ")
_n_online = sum(1 for _t in _tw if _t["type"] == "online")
_meta     = _site.get("meta", {})
_students = _meta.get("onlineStudents", "1,100+")
_sessions = _meta.get("invitedSessions", "20+")
_years    = _meta.get("invitedYears", "2022 — 2026")
p = doc.add_paragraph()
r = p.add_run(f"{_n_univ} university courses · {_sessions} invited lectures and workshops.")
_set_run(r, 8.5, italic=True, color=LIGHT); p.paragraph_format.space_after = Pt(3)

def _t_sub(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6)
    r = p.add_run(text); _set_run(r, 9.5, bold=True, color=ACCENT)

def _t_item(text, date, url=None):
    p = doc.add_paragraph(); add_right_tab(p); p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.first_line_indent = Inches(-0.28)
    b = p.add_run("• "); _set_run(b, 9, color=ACCENT)
    if url:
        add_hyperlink(p, url, text, color="1a1a1a", size=9)   # clickable title, normal colour
    else:
        r = p.add_run(text); _set_run(r, 9)
    r = p.add_run("\t" + date); _set_run(r, 9, color=GREY)

# CV 표기(EN) 가 채워진 행만 싣는다. 여러 회차를 한 줄로 묶은 경우 대표 행에만 표기가 있다.
def _teach(kind):
    return sheet("Teaching", where=lambda r: r["Type"] == kind and r["CV 표기(EN)"],
                 order=lambda r: _desc(r["Start"]))


_t_sub("University Courses")
for _u in _teach("univ"):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.left_indent = Inches(0.28); p.paragraph_format.first_line_indent = Inches(-0.28)
    _set_run(p.add_run("• "), 9, color=ACCENT)
    _set_run(p.add_run(_u["CV 표기(EN)"]), 9)
    _lines = str(_u["CV 서술(EN)"] or "").splitlines()
    for _i, _line in enumerate(_lines):
        p = doc.add_paragraph(); p.paragraph_format.left_indent = Inches(0.28)
        if _i == len(_lines) - 1:
            add_right_tab(p)
        _set_run(p.add_run(_line), 9, color=GREY)
        if _i == len(_lines) - 1:
            _set_run(p.add_run("	" + _txt(_u["CV 기간(EN)"])), 9, color=GREY)
    # 연표는 강의 한 과목을 1건으로, 학기 시작월(Start) 에 놓는다
    _y, _m = _ym2(_u["Start"])
    _tl(_y, _m, "Teaching", _txt(_u["CV 표기(EN)"]), _txt(_u["CV 기간(EN)"]))

_t_sub(f"Invited Lectures & Workshops  ({_sessions} sessions, {_years})")
for _v in _teach("special"):
    _t_item(_txt(_v["CV 표기(EN)"]), _txt(_v["CV 기간(EN)"]))
    _y, _m = _ym2(_v["Start"])
    _tl(_y, _m, "Teaching", _txt(_v["CV 표기(EN)"]))

# ---------------------------------------------------------------- Professional Experience
section("Professional Experience")
prof = _work("professional")
for lead, rest, date, sub, _start in prof:
    entry(lead, right=date, sub=sub, left_rest=rest)
    _y, _m = _ym2(_start)
    _tl(_y, _m, "Position", lead, rest)

# ---------------------------------------------------------------- References
section("References")
# 공개 CV에는 추천인 연락처를 노출하지 않음(사생활). 실제 3인 연락처는 비공개 레퍼런스 시트로 제출.
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
_set_run(p.add_run("Available upon request."), 9, italic=True, color=GREY)

# ---------------------------------------------------------------- Timeline (CV 맨 아래)
# 위 섹션들이 렌더되는 동안 _tl() 로 모인 TIMELINE 을 하나의 연표로 인쇄한다.
# 항목을 여기서 새로 수집하지 않으므로 본문과 어긋날 수 없다.
section(f"Timeline ({TL_FROM} — Present)")
p = doc.add_paragraph()
_set_run(p.add_run(f"All confirmed activity since January {TL_FROM} in a single chronology "
                   "(newest first). Manuscripts under review are not included."),
         8.5, italic=True, color=LIGHT)
p.paragraph_format.space_after = Pt(3)

_TL_LAB = Inches(0.34)    # 월 다음, 카테고리 라벨이 시작하는 자리
_TL_TXT = Inches(1.40)    # 서술이 시작하는 자리 = 줄바꿈 시 유지되는 들여쓰기
# 같은 달에 여러 건이 있을 때의 표시 순서 (본문 섹션 순서에 맞춘다)
_TL_CATS = ["Publication", "Book", "Conference", "Invited Talk", "Teaching",
            "Project", "Funding", "Award", "Patent", "Position"]
_TL_ORDER = {c: i for i, c in enumerate(_TL_CATS)}

for _ty in sorted({_i["y"] for _i in TIMELINE}, reverse=True):
    ph = doc.add_paragraph(); ph.paragraph_format.space_before = Pt(8)
    _set_run(ph.add_run(str(_ty)), 9.5, bold=True, color=ACCENT)
    # 월 내림차순. 월 미상(0)은 -(0 or -99) = 99 가 되어 그 연도 그룹의 끝으로 간다.
    for _i in sorted((x for x in TIMELINE if x["y"] == _ty),
                     key=lambda x: (-(x["m"] or -99), _TL_ORDER.get(x["cat"], 99), x["main"])):
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.left_indent = _TL_TXT
        p.paragraph_format.first_line_indent = Inches(-1.40)
        for _tab in (_TL_LAB, _TL_TXT):
            p.paragraph_format.tab_stops.add_tab_stop(_tab, WD_TAB_ALIGNMENT.LEFT)
        _set_run(p.add_run(_MONTHS.get(_i["m"], "")), 9, bold=True, color=GREY)
        _set_run(p.add_run("\t" + _i["cat"]), 8, color=LIGHT, caps=True, spacing=10)
        _set_run(p.add_run("\t" + _i["main"]), 9)
        if _i["meta"]:
            _set_run(p.add_run("  · " + _i["meta"]), 8.5, italic=True, color=GREY)

print("timeline items:", len(TIMELINE))
for _c in _TL_CATS:
    _n = sum(1 for _i in TIMELINE if _i["cat"] == _c)
    if _n:
        print(f"  {_c:<13} {_n}")
for _ty in sorted({_i['y'] for _i in TIMELINE}, reverse=True):
    print(f"  {_ty}: {sum(1 for _i in TIMELINE if _i['y'] == _ty)}"
          f"  (month unknown: {sum(1 for _i in TIMELINE if _i['y'] == _ty and not _i['m'])})")

# (Technical Skills section intentionally omitted)

# ---------------------------------------------------------------- last updated / (opt-in) signature
# 공개 CV 에는 서명을 넣지 않는다 — 영문 학술 CV 관행이 아니고(cv/refs 4편 모두 없음),
# 서명 이미지는 신원 자격증명이라 공개 PDF·저장소에 두면 위조 위험이 있다.
# 기본은 'Last updated: Mon YYYY' 한 줄. 서명본을 요구하는 지원서에만 CV_SIGNED=1 로 빌드하고,
# 서명 파일은 저장소 밖 %USERPROFILE%\.config\gugi516\signature.png 에서만 읽는다 (커밋 금지).
_sig_img = os.path.join(os.path.expanduser("~"), ".config", "gugi516", "signature.png")
if os.environ.get("CV_SIGNED") == "1" and os.path.exists(_sig_img):
    pim = doc.add_paragraph(); pim.paragraph_format.space_before = Pt(16)
    pim.add_run().add_picture(_sig_img, height=Inches(0.6))
    p = doc.add_paragraph(); add_right_tab(p); p.paragraph_format.space_before = Pt(1)
    _set_run(p.add_run("Gukhwa Jang"), 9.5, bold=True)
    _set_run(p.add_run("\tDate: " + _dtdate.today().strftime("%d %B %Y")), 9, color=GREY)
else:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_before = Pt(18)
    _set_run(p.add_run("Last updated: " + _dtdate.today().strftime("%b %Y")), 8.5, italic=True, color=LIGHT)

# ---------------------------------------------------------------- save
out = "Jang_Gukhwa_CV_v1.docx"
doc.save(out)
print("saved:", out)
