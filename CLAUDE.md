# gugi516.github.io — 작업 규칙

개인 연구 포트폴리오 사이트(Astro). **이 저장소는 공개(PUBLIC)다.**

가장 중요한 규칙 하나: **`src/data/site.json` 의 사실 데이터를 직접 고치지 말 것.**
그 값들의 정본은 `cv/application_info.xlsx` 이고, 직접 고치면 `cv/_sync_site.py` 가
다음 실행 때 xlsx 값으로 되돌려 버린다.

---

## 1. 정본이 셋이다 — 담당이 겹치지 않는다

| 정본 | 담당 | 위치 |
|---|---|---|
| **`.bib`** | 논문 **서지** — 권·호·페이지·DOI·발행월·저자 | `cv/citations/individual/*.bib` |
| **`application_info.xlsx`** | 그 외 **사실 전부** — 학력·경력·과제·특허·학회·강의·수상·장학·자격증·연구분야·저서, 그리고 논문의 IF·Quartile·저자순위·게재일자 | `cv/application_info.xlsx` |
| **`site.json`** | **표시용** — 로고·표지·PDF 경로·URL·id, hero·about.stats·nav·qrTargets, ko/en 웹 문구 | `src/data/site.json` |

```
citations/individual/*.bib ──merge_bib.py──▶ publications.bib
                                              ├─▶ build-citations.mjs ─▶ src/data/citations.json ─▶ 웹 인용문
                                              └─▶ build_cv.py 논문 섹션
application_info.xlsx ──┬─▶ build_cv.py (13시트) ─▶ Jang_Gukhwa_CV_v1.docx
                        ├─▶ _sync_site.py ─▶ site.json 사실 필드 ─▶ 웹
                        └─▶ _export_csv.py ─▶ cv/_csv/ (git diff 용 미러)
```

`npm run build` = `build-citations.mjs` → `optimize-images.mjs` → `astro build`.

---

## 2. 금지 사항

- **`site.json` 의 사실 필드 직접 수정 금지.** 논문 IF, 학회 제목·장소·저자, 학력, 과제,
  특허, 강의, 경력, 장학 기간, 자격증 — 전부 xlsx 에서 고치고 `_sync_site.py` 로 반영한다.
  (로고·표지·PDF 경로·hero·about·nav·QR 같은 **표시용 필드는 site.json 이 정본**이므로 직접 수정)
- **`build_cv.py` 에 데이터를 다시 하드코딩 금지.** `sheet("시트명", where=..., order=...)` 로 읽는다.
- **`cv/_csv/` 직접 편집 금지.** `_export_csv.py` 가 덮어쓰는 생성물이다.
- **`publications.bib` 직접 편집 금지.** `individual/*.bib` 를 고치고 `merge_bib.py` 를 돌린다.
- **`site.json` 을 `json.dump` 로 다시 쓰지 말 것.** 손으로 맞춘 빈 줄·정렬이 사라져
  약 2,950줄이 통째로 바뀌고 리뷰가 불가능해진다. `_sync_site.py` 는 문자열 리터럴만 국소 치환한다.
- **xlsx 의 시트명·컬럼명 변경 금지.** `build_cv.py` 가 이름으로 값을 찾는다.
  꼭 바꿔야 하면 `build_cv.py` 와 `_check_sync.py` 를 함께 고친다.
- **`application_info.xlsx` 를 커밋하지 말 것.** `.gitignore` 43~46행으로 제외되어 있다
  (개인정보 포함, 이 저장소는 공개). 규칙을 지우지 말 것.

---

## 3. 작업 절차

### 논문 추가
```
1. 출판사 BibTeX 를 cv/citations/original/ 에 원본 그대로 저장
2. 같은 파일을 individual/ 에 복사 → normalize_bib.py 로 양식 통일
3. cd cv/citations && python merge_bib.py && python bib_to_xlsx.py
4. xlsx Publications 시트에 행 추가
   - bib 이 갖는 것(권·호·페이지·DOI)은 참고용으로만, 정본은 bib
   - 여기서만 입력: Role(first/co) · Index(SCIE/ESCI/KCI) · IF · Quartile
     · 저자순위(본인) · 총저자수 · 게재일자
5. cd cv && python _export_csv.py && python _check_sync.py
6. site.json publications.items 에 항목 추가 (표지·PDF 경로는 사람이 결정)
7. python _sync_site.py        미리보기 → 이상 없으면 --write
8. python build_cv.py && python _cv_baseline.py --check
```
> `_sync_site.py` 는 **항목을 추가하지 않는다.** 기존 항목의 필드만 갱신한다.
> 새 논문·학회를 웹에 올릴 때는 site.json 에 항목 골격을 먼저 만들어야 한다.

### 초청 강연 추가
```
1. xlsx 'Invited Talks' 시트에 행 추가 (Date·Type·Scope·주최(KO)/Host(EN)·웹 표시=Y, No. 는 최신순 재부여)
   → 넣을 시트는 4절 "세 시트의 경계" 판정 한 줄로 고른다
2. python _export_csv.py && python _check_sync.py
3. site.json invitedTalks.items 에 항목 골격 추가 후 python _sync_site.py --write
4. python build_cv.py && python _cv_baseline.py --check
```

### 학회 발표 · 특허 · 강의 · 경력 · 수상 · 자격증 추가
```
1. xlsx 해당 시트에 행 추가 (+ 아래 4절의 제어 컬럼 지정)
2. python _export_csv.py && python _check_sync.py
3. 웹에도 올릴 항목이면 site.json 에 항목 추가 후 python _sync_site.py --write
4. python build_cv.py && python _cv_baseline.py --check
```

### 웹 문구·이미지만 수정
`site.json` 을 직접 고친다. 사실 필드가 아니면 xlsx 와 무관하다.
수정 후 `python cv/_sync_site.py` 로 "변경 없음" 이 나오는지만 확인한다.

### CV 재생성
```
cd cv && python build_cv.py            → Jang_Gukhwa_CV_v1.docx
python _cv_baseline.py --check         → 의도한 줄만 바뀌었는지 확인
python _cv_baseline.py --save          → 승인된 변경이면 기준선 갱신
```
PDF 는 **반드시 `powershell -File make_pdf.ps1 -Publish`** (LibreOffice headless) 로 만든다 — 최신 `Jang_Gukhwa_CV_v*.docx` 를 변환하고 `public/cv/Jang_Gukhwa_CV.pdf` 까지 복사한다.
**Word 내보내기·docx2pdf·Word COM 금지.** Word 는 Cambria 를 서브셋 여러 개로 쪼개고 일부를 임베드하지 않아 뷰어에서 자간이 깨진다
(2026-06-06 진단, 2026-09-15 재발 — 그때 `make_pdf.ps1` 을 old/ 로 치우고 docx2pdf 를 쓴 것이 원인). `make_pdf.ps1` 은 old/ 로 옮기지 말 것.

---

## 4. xlsx 구조 (18시트)

> 행 수는 본인 `application_info.xlsx` 실측 (2026-10-13, 데이터 입력 진행 중 — 채우면서 갱신).
> 시트명·컬럼 구조는 원주인 템플릿 그대로다.

| 시트 | 행 | 제어·CV 표기 컬럼 |
|---|---|---|
| Personal | 16 | — (지원서 전용, CV·웹 미사용) |
| Research Interests | 6 | — |
| Education | 3 | `CV 표시` `웹 표시` `CV 서술(EN)` `심사위원(EN)` |
| Publications | 7 | — (1행은 원주인 템플릿 잔재 의심 — 삭제 대기, bib·웹은 6편) |
| Under Review | 12 | — |
| Books | 2 | — |
| Conferences | 15 | `CV 학회(EN)` `팀 발표` |
| Invited Talks | 3 | `웹 표시` |
| Teaching | 14 | `CV 표기(EN)` `CV 서술(EN)` `CV 기간(EN)` |
| Funding | 9 | `CV 표기(EN)` |
| Projects | 13 | `CV 과제명(EN)` `CV 기여(EN)` `CV 성과(EN)` |
| Patents | 3 | `CV 특허명(EN)` `CV 권리자(EN)` |
| Awards | 7 | — |
| Service & Membership | 2 | — |
| Certifications | 1 | — |
| Language Tests | 0 | — |
| Technical Skills | 5 | — |
| Work & Internships | 6 | `웹 표시` `CV 구분` `CV 직위(EN)` `CV 소속(EN)` `CV 서술(EN)` |

**헤더가 황토색인 컬럼**은 나중에 덧붙인 것이다 — 국내 지원서용이거나 CV 표기용.

### 제어 플래그
| 컬럼 | 값 | 뜻 |
|---|---|---|
| `CV 표시` | Y / N | CV 수록 여부 (학력 중 CV 에 싣지 않을 단계는 N) |
| `웹 표시` | Y / N | 웹 노출 여부 (다른 시트와 중복되는 행은 N) |
| `CV 구분` | research / professional | Work 를 CV 의 Research Experience / Professional Experience 로 나눔 |
| `팀 발표` | Y / N | Y 면 CV 에 `[Oral, team]` 처럼 team 이 붙음 |

> CV 국제학회 10건 = `Conferences` 의 `Scope=international`
> CV 초청강연 3건 = `Invited Talks` 전 행 (2026-09-24 통합 — 학회 시트의 `CV 분류` 열은 삭제)
> 같은 사건은 한 행에만 존재한다. 두 시트에 중복 등록하지 말 것.

### 세 시트의 경계 — 어디에 넣는가
- **Teaching** (`특강`/special) = 본인이 **주강사(메인)인 교육 과정**. 반나절~수일, 수강생·시수를 관리한다 (K-water·환경공단·현대NGV 등).
- **Invited Talks** = **남의 과목·행사에 초청되어 1~2시간 안에 끝나는 강연·발표**. 본인이 메인이 아니다 (교과목 초청강의, 세미나·콜로퀴엄·기조·패널).
- **Conferences** = 학회 프로그램 안의 발표. **학회 초청 세션은 `Invited Talks`** 로 (`Type=conference`).

> 판정 한 줄: 내가 과정의 주강사인가(→Teaching), 남의 자리에 1~2시간 초청됐나(→Invited Talks), 학회 발표인가(→Conferences).

`Invited Talks` 의 `Type` 어휘(영문 소문자): `seminar` · `colloquium` · `keynote` · `panel` ·
`guest lecture`(교과목 초청강의) · `conference`(학회 프로그램 안 초청 발표).

### `CV ...(EN)` 컬럼의 의미
원자 필드로 CV 표기를 **규칙으로 만들 수 없을 때만** 둔다. 공식 명칭과 CV 표기가
다른 것은 정상이며, 둘 다 필요한 정보다.
```
Patents.Title(EN)      = 공식 등록 특허명 (지원서·산정표용, 직역투로 길다)
Patents.CV 특허명(EN)   = CV 표기 (읽히게 다듬음)

Conferences.Conference(EN) = 학회 정식명  "AGU (American Geophysical Union)"
Conferences.CV 학회(EN)    = 회차명       "AGU Fall Meeting"

Work.Institution Name(EN) = 정식 기관명  "Virginia Tech (Virginia Polytechnic Institute...)"
Work.CV 소속(EN)          = CV 표기      " — Virginia Tech (PI: Prof. Jaeyoung Ha)"
```

---

## 5. 표기 규칙

- **논문·발표 제목은 APA 문장형**(sentence case)으로 CV 에 인쇄된다.
  xlsx 에는 **실제 초록 제출 제목을 그대로** 저장하고, 변환은 `build_cv.py` 의
  `sentence_case()` 가 한다. 직접 소문자로 바꿔 저장하지 말 것.
- 약어·고유명사는 `_KEEP_WORDS` / `_KEEP_PHRASES` 로 보호한다
  (`LiDAR` `GEDI` `XGBoost` `UAV` `RGB` `Sentinel-2` `Gyeonggi-do Province` 등).
  새 약어가 등장하면 이 목록에 추가한다. 안 하면 `LiDAR` 가 `lidar` 가 된다.
- **제목 끝 마침표는 xlsx 에 넣지 않는다.** 렌더러(`_dot()`)가 붙인다.
- 기간의 이음표는 **엔대시(–)** 또는 **엠대시(—)** 를 원문대로 쓴다.
  ASCII 하이픈(-)으로 바꾸지 말 것. 과거에 이 전사 오류가 여러 번 있었다.
- 국문 제목이 없는 항목의 `...(KO)` 칸에는 **영문을 그대로 복사**해 둔다.
  한쪽만 고치면 `_sync_site.py` 가 매번 웹을 되돌리려 든다(비멱등).
- **영문 CV 기간은 연도만** — 학력·직위는 `2022 — 2026`, 같은 해면 `2024`, 진행 중이면 `2020 — Present`.
  수여 월 같은 **특정 시점만 `Mon YYYY`**(헤더 `Ph.D., Aug 2026`). `YYYY.MM`(한국식)은 CV 에 쓰지 않는다 —
  xlsx 는 `YYYY.MM` 그대로 두고 `build_cv.py` 의 `_year_range()`/`_mon_yyyy()` 가 변환한다(웹은 `YYYY.MM` 유지).
  근거: `cv/refs` 샘플 CV 4편 전부 학력·직위를 연도만 표기(2026-09-06 대조).

## 6. 정렬 규칙

- CV 는 **전 섹션 최신순**이다. `sheet(..., order=lambda r: _desc(r["날짜컬럼"]))`.
- **엑셀 행 순서에 의존하지 말 것.** 행을 위아래로 옮겨도 출력이 바뀌면 안 된다.
  날짜가 없는 시트는 `order=lambda r: r["No."]` 로 명시한다.

---

## 7. 도구

| 명령 | 하는 일 | 종료코드 |
|---|---|---|
| `python cv/_export_csv.py` | xlsx → `cv/_csv/*.csv` 미러. **xlsx 수정 후 반드시 실행** | — |
| `python cv/_check_sync.py` | xlsx ↔ site.json 대조, 미입력 컬럼, 리팩터 회귀 감시 | 0=정상 |
| `python cv/_sync_site.py` | 변경 미리보기 | 1=변경 있음 |
| `python cv/_sync_site.py --write` | site.json 반영 (.bak 생성 + JSON 파싱 검증) | — |
| `python cv/_cv_baseline.py --check` | CV 출력이 기준선과 같은지 | 0=동일 |
| `python cv/_cv_baseline.py --save` | 승인된 변경을 새 기준선으로 | — |

`_sync_site.py` 의 안전장치: 표시용 필드 미변경 · 빈 값으로 덮지 않음 ·
항목 추가삭제 없음 · 국소 치환 · 매칭 실패 보고.

---

## 8. 알려진 함정

- **`_sync_site.py` 의 매칭 실패는 조용히 넘어갈 뻔했다.** 지금은 보고하지만,
  "매칭 주의" 가 뜨면 xlsx 와 site.json 의 키(날짜·제목·DOI)가 어긋난 것이다.
- **`ko`/`en` 쌍 배열**(`patents[].inventors`)을 평평한 배열로 덮으면 한글이 사라진다.
  `put_pair_list()` 를 쓴다.
- **`build_cv.py` 안의 이름 충돌.** 섹션 루프가 `yr`, `p`, `r` 같은 짧은 이름을
  재할당한다. 새 헬퍼는 긴 이름(`years_only`, `sentence_case`)을 쓴다.
- **시트를 지우면 `_check_sync.py` 가 크래시하지 않고 보고한다.** 보고가 뜨면
  시트명이 바뀐 것인지 확인한다.
- **short CV 는 은퇴했다.** `build_cv_short.py` 와 산출물(`Lee_Seunghyeon_CV_short.docx/.pdf`)은
  `cv/old/` 로 옮겼고, `_cv_baseline.py` 의 검증 대상에서도 제외했다. 이제 CV 는 full 한 종류다.
- **`cv/Jang_Gukhwa_CV_v1.docx` 와 `public/cv/Jang_Gukhwa_CV.pdf`** 는
  같은 CV 의 다른 이름이다. CV 를 갱신하면 두 곳 다 바꿔야 사이트가 구버전을 안 보여준다.
- **Windows cp949 콘솔에서 `_sync_site.py`·`_check_sync.py` 출력이 `UnicodeEncodeError` 로
  죽을 수 있다** (em-dash 등 출력 시). `PYTHONIOENCODING=utf-8` 을 주고 실행한다. (2026-10-13 실측)
- **`make_pdf.ps1` 은 `powershell -File` 단독 호출이 ExecutionPolicy 에 막힌다** (`UnauthorizedAccess`).
  `powershell -NoProfile -ExecutionPolicy Bypass -File cv\make_pdf.ps1` 로 실행한다. (2026-10-13 실측)
- **플래그 컬럼 값이 코드의 기대 토큰과 다르면 에러 없이 섹션이 조용히 빈다.**
  정확한 토큰: Work `CV 구분`=`research`/`professional`, Teaching `Type`=`univ`/`online`/`special`,
  Service & Membership `Type`=`service`/`membership`, Conferences `Scope`=`international`.
  섹션이 통째로 비면 이 철자부터 확인한다.
- **Patents `Status` 는 한글로 써도 된다.** `_sync_site.py` 의 `patent_status()` 가
  `등록`·`등록 (디자인)` → `registered`, `출원` → `application` 으로 자동 매핑한다
  (영문 토큰을 그대로 써도 통과). 이 매핑이 없으면 `--write` 때 site.json 의
  `status` 가 한글로 덮여 `Home.astro` 의 `status === 'registered'` 필터에서
  등록 특허 카드가 통째로 사라진다. (2026-10-20 수정)

---

## 9. 커밋 전 체크리스트

```
cd cv
python _export_csv.py
python _check_sync.py      → "확인 필요 항목 없음"
python _sync_site.py       → "반영할 변경 없음"
python _cv_baseline.py --check   → "기준선과 동일" (또는 의도한 줄만)
cd .. && npm run build     → "[build] Complete!"
git status                 → application_info.xlsx / _csv/ 가 안 잡히는지 확인
```
