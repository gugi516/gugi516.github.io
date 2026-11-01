# 기관 로고 (Logos)

> 이 문서는 `public/logos/` 폴더 설명서다. public/ 아래 파일은 전부 URL 로 공개되므로
> 문서(README)는 여기 docs/public_assets/ 에 둔다 (2026-10-21 이동).

학력·경력 기관의 로고를 SVG로 저장합니다.

## 폴더 구조

```
logos/
├── education/      # 학력 기관 로고
├── career/         # 경력 기관 로고
└── scholarships/   # 장학금 재단 로고
```

## 권장 형식

**SVG 우선** (벡터 → 모든 크기에서 깔끔). 없으면 PNG (배경 투명).

- 사이즈: 라이브 렌더 크기는 ~40×40px ~ 80×80px이므로, SVG는 원본 viewBox 그대로 두면 됩니다.
- 배경: **투명**. 다크 모드에서도 보여야 하므로 배경 흰색 PNG는 가급적 피하세요. 흰색 PNG만 있으면 다크 모드용 backdrop이 자동 추가됩니다.

## 파일 이름 규칙

기관의 일반적인 약칭 또는 영문 이름. 소문자, 단어 구분은 하이픈.

### Education (학력)
- `snu.svg` — Seoul National University
- `snu-gses.svg` — SNU Graduate School of Environmental Studies (별도 로고 있다면)
- `knu.svg` — Kyungpook National University

### Career (경력 · 인턴십)
- `virginia-tech.svg`
- `purdue.svg`
- `auri.svg` — 건축공간연구원 (Architecture & Urban Research Institute)
- `sunjin.svg` — (주)선진엔지니어링종합건축사사무소
- `atlas-china.svg` — ATLAS Landscape Architecture CHINA

### Scholarships (장학금 재단)
- `cmk-foundation.svg` — 현대차 정몽구 재단 (Chung Mong-Koo Foundation)
- `bk21.svg` — BK21
- `ilju.svg` — 일주학술문화재단
- `knu.svg` — 경북대학교 (Education 폴더와 공유해도 OK — 경로만 맞추면 됩니다)

## 로고 구하는 법

1. **기관 공식 홈페이지의 "Brand / Press / Media Kit"**: 공식 SVG/PNG 제공
2. **Wikimedia Commons**: 대학·정부기관 로고는 대부분 SVG로 등록되어 있음 (CC 라이선스 확인)
3. **로고가 없다면**: 빈 채로 두세요. 카드 컴포넌트가 자동으로 placeholder(이니셜 또는 기관명) 표시.

## site.json 연결

```json
{
  "period": "2022.03 — Present",
  "logo": "/logos/education/snu.svg",
  "institution": { "ko": "서울대학교", "en": "Seoul National University" },
  ...
}
```

- 경로는 `/logos/...` 로 시작.
- `logo` 필드가 빈 문자열이면 placeholder 표시.

## 저작권

기관 로고는 대부분 trademark입니다. 본인 학력·경력의 출처를 표시하는 일반적인 학자 포트폴리오 용도라면 통상 fair use이지만, 사용 가이드라인이 있는 기관이면 따르세요 (예: 일부 대학은 학외 사이트 사용 시 사전 허가 요구).
