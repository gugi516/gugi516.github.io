# 저서 자산 (Books)

> 이 문서는 `public/books/` 폴더 설명서다. public/ 아래 파일은 전부 URL 로 공개되므로
> 문서(README)는 여기 docs/public_assets/ 에 둔다 (2026-10-21 이동).

## 폴더 구조

```
books/
└── covers/   # 책 표지 이미지 (.jpg / .png)
```

## 파일 이름 규칙

`book_{N}_{short_name}.{ext}` 형식. 순서대로 1, 2, 3...

예시:
- `covers/book_01_new_normal_city.jpg` (뉴노멀시티)

## 표지 이미지 만드는 법

- yes24, 교보문고, 알라딘 등의 도서 상세 페이지에서 표지 이미지 다운로드 (보통 우클릭 → 이미지 저장)
- 또는 출판사 사이트의 도서 페이지 표지
- 권장 사양: 가로 600~900px, A4 비율과 비슷한 세로형이 카드에 자연스럽게 맞음

## site.json 연결

`src/data/site.json` → `honors.books.items[]` 에 `cover` 와 `url` 필드:

```json
{
  "title": { "ko": "뉴노멀시티", "en": "New Normal City" },
  "cover": "/books/covers/book_01_new_normal_city.jpg",
  "url":   "https://www.yes24.com/Product/Goods/...",
  ...
}
```

- `cover`: 없으면 빈 문자열 → 책 제목 표시한 그라데이션 placeholder
- `url`: 책 소개 페이지 (yes24, 교보문고, 출판사 페이지 등). 없으면 빈 문자열 → 외부 링크 버튼 비표시
