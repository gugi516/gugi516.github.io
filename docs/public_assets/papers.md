# 논문 자산 (Papers)

> 이 문서는 `public/papers/` 폴더 설명서다. public/ 아래 파일은 전부 URL 로 공개되므로
> 문서(README)는 여기 docs/public_assets/ 에 둔다 (2026-10-21 이동).

각 게재 논문의 표지 이미지와 원본 PDF를 여기에 저장합니다.

## 폴더 구조

```
papers/
├── covers/     # 논문 표지 이미지 (.jpg / .png / .webp)
└── pdfs/       # 논문 원본 PDF (다운로드용)
```

## 파일 이름 규칙

`{연도}_{저널약어}_{제목-시작-단어}.{확장자}` 형식.

- **연도**: 게재 연도 4자리 (`2026`, `2025`, ...)
- **저널약어**: 학계 통용 약어 (5자 이하면 풀이름, 길면 단축). 띄어쓰기·콤마 제거.
  - 예: `KOSERT`, `RSE`, `AEE`, `GECCO` (5자 이하 → 풀이름)
  - 예: `EcolIndic` (Ecological Indicators), `RemoteSens` (Remote Sensing MDPI), `JGR-Bio` (J. Geophys. Res. Biogeosciences)
- **제목-시작-단어**: 제목 첫 1~3개 단어, 하이픈 연결. 영문만, 띄어쓰기 X, 특수문자 X.

### 현재 논문 파일명 (참고 매핑)

게재 논문 7편 전부 **표지(cover) 있음**, PDF 는 3편만 게시.

| id | 연도 | 저널 | 파일명 | cover | pdf |
|---|---|---|---|---|---|
| 1 | 2019 | Journal of Architecture and Urbanism | `2019_JAU_Investigating-the-effect` | ✅ | ✅ (gold OA, CC BY — 출판본 게시 가능) |
| 2 | 2021 | Urban Climate | `2021_UC_Are-decline-oriented` | ✅ | — (Elsevier) |
| 3 | 2022 | Building and Environment | `2022_BAE_Planning-scenarios` | ✅ | — (Elsevier) |
| 4 | 2025 | Nature Cities | `2025_NatCities_Association-between` | ✅ | — (Springer Nature) |
| 5 | 2025 | 국토계획 (JKPA) | `2025_JKPA_Climate-Gentrification` | ✅ | ✅ (KCI, 게재본) |
| 6 | 2026 | Sustainable Cities and Society | `2026_SCS_What-health-risks` | ✅ | — (Elsevier) |
| 7 | 2026 | 한국환경복원기술학회지 (KOSERT) | `2026_KOSERT_Comparing-LST-Estimation` | ✅ | ✅ (KCI, 게재본) |

- 표지 7개는 **게재본 PDF 1쪽을 200dpi 렌더 → 가로 1000px → WebP 품질 88** 로 생성 (2026-10-20).
  원본 PDF 는 `D:\my_research\02.projects\04.cv_website\files\논문\` 에 보관. 사이트에는 표지 이미지만 올라간다.
- 표지는 논문 첫 페이지(제목·저자·초록)라 **본문 재배포가 아닌 서지 표시** 성격이다. PDF 전문 게시와는 별개.
- KCI 2편(id 5·7)의 **PDF 전문**은 사용자 결정으로 게시 — 두 학회 모두 자가보관(self-archiving) 정책을
  공개하지 않아 명시적 허가는 확인되지 않았다.

PDF 전문이 없는 4편(id 2·3·4·6)은 표지 + DOI 링크만 — 저작권 사유 (아래 "저작권 주의" 참조).
Elsevier 3편은 심사통과 원고(accepted manuscript) 확보 시 CC BY-NC-ND 표시로 추가 가능.

예시:
- `covers/2019_JAU_Investigating-the-effect.webp`
- `pdfs/2019_JAU_Investigating-the-effect.pdf`

## 표지 이미지 만드는 법

1. **저널 표지 캡처**: 출판 저널 홈페이지에서 해당 호(volume/issue)의 cover image 다운로드.
2. **논문 PDF 1페이지 캡처**: PDF 첫 페이지를 그래픽으로 캡처 (PowerPoint, Adobe, Preview 등에서 PDF → 이미지 export). A4 비율(210×297)을 유지하면 카드 안에 정확히 fit.
3. **Graphical abstract 활용**: 일부 저널은 graphical abstract를 제공 — 그것도 OK.

**권장 사양**:
- 비율: A4 세로 (210:297 = 1:1.414). 카드 컴포넌트가 `aspect-[210/297]` 박스 안에 `object-cover`로 표시합니다.
- **해상도: 가로 800~1200px (필수)** — 화면에 약 290px로 표시되지만 HiDPI/Retina 화면에서 2x 픽셀 필요. **600px 이하는 글자가 흐려 보입니다.**
- **형식: PNG 강력 권장** (논문 제목·저자 같은 텍스트가 많은 페이지). JPG는 텍스트 가장자리 압축 손실로 흐림 발생. 사진형 표지면 고품질 JPG(품질 90+) 가능.

**현재 쓰는 방법 (2026-10-20 기준)**: PyMuPDF(fitz)로 1쪽을 200dpi 렌더 → Pillow LANCZOS 로 가로
1000px 축소 → WebP 품질 88 저장. PDF 페이지 비율을 그대로 유지하므로 저널에 따라 1.33~1.41 로
갈리지만, 카드가 `aspect-[210/297]` + `object-cover` 라 좌우 여백만 미세하게 잘린다.

**생성물 경로**: 원본은 `covers/<이름>.webp`, 사이트가 실제로 참조하는 것은
`covers/web/<이름>.webp` (800px). `web/` 파생본은 `npm run build` 가 `scripts/optimize-images.mjs`
를 통해 자동 생성하므로 **손으로 만들지 않는다.** `site.json` 에는 원본 경로(`/papers/covers/...`)를
적고, `CardPaper.astro` 가 `/web/` 경로로 치환한다.

**PDF → 고해상도 이미지 변환 방법 (수동 대안)**:
- **Adobe Acrobat**: File → Export To → Image → PNG, **해상도 300 DPI**
- **macOS Preview**: PDF 페이지 우클릭 → Export As → Format: PNG, **Resolution 300 dpi**
- **온라인**: [pdf2png.com](https://pdf2png.com) (High Quality 옵션)
- **무료 PowerPoint 방법**: PDF 1페이지를 슬라이드에 삽입 → 우클릭 Save as Picture → PNG 선택

## site.json 연결

`src/data/site.json` → `publications.items[].cover` / `.pdf` 필드에 경로를 적어주세요.

```json
{
  "id": 7,
  "cover": "/papers/covers/2026_EcolIndic_Multi-scale-typologies.webp",
  "pdf":   "/papers/pdfs/2026_EcolIndic_Multi-scale-typologies.pdf",
  ...
}
```

- 경로는 항상 `/papers/...` 로 시작 (public 폴더 기준 절대 경로).
- 파일이 없으면 빈 문자열 `""` 로 두세요 — 카드 컴포넌트가 자동으로 placeholder(저널명·연도 그라데이션)를 표시하고, PDF/Web 버튼은 회색 비활성 상태가 됩니다.

## 새 논문 추가 시 작업 흐름

1. 표지 + PDF 파일을 위 규칙으로 명명
2. 각각 `covers/`와 `pdfs/`에 저장
3. `site.json`의 `publications.items` 맨 앞에 새 항목 추가:
   - `id` 는 기존 최대값 + 1
   - `cover`, `pdf` 경로 입력
   - 나머지 필드 (date, year, role, journal, doi, authors, title) 채움
4. `git add . && git commit -m "add: 논문 #N 추가" && git push`

## 저작권 주의

논문 PDF를 자신의 사이트에 호스팅하려면 출판사 라이선스 확인 필요:
- **Open access (CC-BY 등)**: 자유롭게 게시 가능
- **Subscription 논문**: 보통 저자 accepted manuscript(피어리뷰 통과 후 출판 전 버전, "preprint" 또는 "postprint")만 공개 가능. publisher version은 불가
- **확실하지 않으면**: PDF 대신 DOI 링크만 두는 게 안전 (site.json `doi` 필드는 이미 활용 중)
