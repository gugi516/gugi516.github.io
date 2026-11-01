# citations/ — CV 인용(citation) 소스 관리

CV publication 목록(APA 7th)의 소스. **original(원본보호) → individual(정규화) → merge(합본) → xlsx** 구조.

## 폴더 구조
```
cv/citations/
├── original/          ← 원본 보호. 논문 페이지에서 받은 BibTeX 그대로. 파이프라인이 '읽기만/안 읽음', 수정 안 함.
├── individual/        ← 정규화된 작업본(통일 양식). merge 가 읽는 source of truth.
│   └── *.bib            (1편 1파일, 양식 통일: "Last, First" + 동일 필드순서)
├── original/, individual/ 는 같은 파일명을 쓰면 추적이 쉬움 (예: lee2021_rs.bib)
├── normalize_bib.py   ← individual/ 의 .bib 들을 통일 양식으로 정규화(in-place)
├── merge_bib.py       ← individual/ 를 합쳐 publications.bib 생성 (DOI 중복제거 + 출판 오름차순)
├── bib_to_xlsx.py     ← publications.bib → publications.xlsx (DOI 클릭링크, 빠른 확인용)
├── publications.bib   ← AUTO-GENERATED (편집 금지)
├── publications.xlsx  ← AUTO-GENERATED (바로 열어 확인)
└── README.md
```
> 주의: `merge_bib.py` 는 `individual/*.bib` 만 읽습니다(`original/` 은 안 읽음). 원본은 안전하게 보존됩니다.

## 새 논문 추가 워크플로우
1. 논문 페이지에서 BibTeX 받기 → **`original/` 에 저장** (원본, 손대지 않음)
2. 같은 파일을 `individual/` 에도 복사 → **Claude에게 "정규화해줘" 요청**
   (Claude가 `normalize_bib.py` 로 양식 통일 + 깨진 export 수리. 출판사가 month를 빼면 보강 필요 → 알려줌)
3. `python merge_bib.py && python bib_to_xlsx.py`
   → `publications.bib` + `publications.xlsx` 갱신 (출판 오름차순)
4. "publications.bib 반영해줘" → CV publication 섹션 APA 7th 갱신 (본인 굵게 + 1저자 † 유지)

## 정규화(normalize_bib.py)가 자동 처리하는 것
- 저자명 통일 → `Last, First` (붙은 이름 `LeeSeunghyeon`→`Lee, Seunghyeon`, 본인/지도교수 표기 통일)
- 깨진 export 수리(pages 뒤 콤마 누락 등), DOI 통일(URL 제거 / citekey에서 추출)
- 군더더기 필드(abstract/keywords/url/issn/publisher) 제거, 필드순서·citekey(=파일명) 통일
- ⚠ month 없으면 경고만 출력 → 출판사 BibTeX가 month를 빼는 경우 site.json 출판일 등으로 수동 보강

## 메모
- KOSERT(국내저널) export는 콤마 누락·저자명 붙음·doi 필드 없음 등 품질 이슈가 잦음 → normalize가 잡아줌
- ⚠ Jekal(Salix): 공식 게재본 저자순서 = Jekal, Yang, Lee, Song (이승현 3번째). site.json도 이에 맞춤. 연도=2025(28권 6호).
