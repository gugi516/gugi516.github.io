# 특허 자산 (Patents)

> 이 문서는 `public/patents/` 폴더 설명서다. public/ 아래 파일은 전부 URL 로 공개되므로
> 문서(README)는 여기 docs/public_assets/ 에 둔다 (2026-10-21 이동).

등록 특허·디자인의 공보 PDF(`pdfs/`)를 저장합니다.

> 2026-10-21: 특허 섹션을 **목록(행) 전용**으로 축소 — 카드 그리드(`CardPatent.astro`),
> 표지 이미지(`covers/`), 특허증 PDF(접미사 없는 `10-XXXXXXX.pdf`)를 모두 제거했습니다.
> 남은 것은 공보(`_B1`/`_S`) PDF와 `RowPatent.astro` 의 "공보 / Gazette" 버튼뿐입니다.

> ## ⚠️ `public/` 에 두는 순간 공개됩니다
>
> **`public/` 아래의 모든 파일은 빌드 시 `dist/` 로 그대로 복사되고, 사이트가 링크하지
> 않아도 URL 로 직접 접근할 수 있습니다.** (`https://gugi516.github.io/patents/pdfs/<파일명>`)
> `site.json` 의 `pdf`·`certificate` 필드에 연결하지 않았다는 것은 **버튼이 안 보인다는 뜻일 뿐,
> 비공개라는 뜻이 아닙니다.** 파일명을 아는 사람(또는 크롤러)은 그대로 내려받습니다.
>
> **따라서 이 폴더에는 공개해도 되는 파일만 둡니다.** 보관 전용 원본은 저장소 밖에 두세요.
>
> 특히 **특허증(certificate) PDF 는 주민등록번호 앞자리와 특허권자·발명자의 집주소를 담고 있습니다.**
> 여기에 두려면 **반드시 해당 영역을 제거(redaction)한 뒤** 두어야 합니다.
> 이미지 스캔 PDF 라 흰 사각형을 덧그리는 것만으로는 원본 픽셀이 남습니다 —
> pymupdf `apply_redactions(images=PDF_REDACT_IMAGE_PIXELS)` 로 픽셀을 지우고,
> 지운 뒤 페이지를 재래스터화해야 합니다. (2026-10-20 실시)
>
> 한국 특허 정보의 공식 조회는 [KIPRIS](https://www.kipris.or.kr) ·
> [Google Patents](https://patents.google.com) 에서 가능합니다.

## 폴더 구조

```
patents/
└── pdfs/
    ├── 10-XXXXXXX_B1.pdf   # 등록특허공보   (site.json pdf 필드 → "공보 / Gazette" 버튼)
    └── 30-XXXXXXX_S.pdf    # 등록디자인공보 (site.json pdf 필드 → "공보 / Gazette" 버튼)
```

파일 이름 규칙: `{등록번호}{접미사}.pdf` — `_B1` 등록특허공보, `_S` 등록디자인공보.
접미사 없는 파일명(`10-XXXXXXX.pdf`)은 특허증 — **여기 두지 않는다** (아래 매핑 참조).

### 현재 매핑 (site.json `patents.items` 기준)

| id | 등록번호 | 구분 | 등록일 | `pdf` |
|---|---|---|---|---|
| 3 | 10-2862472 | 등록특허 | 2025.09.17 | `pdfs/10-2862472_B1.pdf` |
| 2 | 10-2736050 | 등록특허 | 2024.11.26 | `pdfs/10-2736050_B1.pdf` |
| 1 | 30-1135512 | 등록디자인 | 2021.11.02 | `pdfs/30-1135512_S.pdf` |

공보 3건 모두 게시 전 발명자·특허권자 **집주소를 redaction** 했습니다
(로그: `my_research\02.projects\04.cv_website\_logs\261020_patent_pii_masking.csv`).

## 표시 방식

전 항목 `RowPatent.astro` 한 줄 행으로만 표시 (등록/출원은 행 안의 pill 로 구분).
`certificate` 필드는 더 이상 사용하지 않는다 — 빈 문자열로 둔다.

> `status` 값은 `cv/application_info.xlsx` 의 Patents 시트가 정본이고 `_sync_site.py` 가
> `등록`→`registered`, `출원`→`application` 으로 매핑합니다. site.json 을 직접 고치지 마세요.

## 새 특허 등록 시 작업 흐름

1. `cv/application_info.xlsx` Patents 시트에 행 추가 → `python cv/_export_csv.py && python cv/_check_sync.py`
2. 공보 PDF 에서 **집주소 등 PII redaction** → `pdfs/{등록번호}_B1.pdf` (디자인은 `_S`)
3. `site.json` `patents.items` 에 항목 골격 추가 후 `python cv/_sync_site.py --write`
4. `npm run build` → `dist/` 에 파일이 들어갔는지, 깨진 참조가 없는지 확인
