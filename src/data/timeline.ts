// ───────── Timeline (2023 — Present) ─────────
// 페이지 맨 아래 연표의 재료. 항목을 여기서 새로 입력하지 않고, 이미 각 섹션이 렌더하는
// site.json 데이터(논문·저서·학회 발표·초청 강연·강의·펀딩·과제·특허·경력)에서 파생한다.
// 따라서 본문 섹션과 연표는 언제나 같은 사실을 가리키고, 어긋날 수 없다.
//
// 수집 규칙은 PDF CV 판(커밋 2e30dd7 의 cv/build_cv.py `_tl()`)과 같다:
//   · 수록 기준 = 2023-01 이후에 '일어난' 확정 실적
//   · 심사 중 원고(underReview)는 제외 — 확정 실적이 아니다
//   · 기간이 있는 항목(과제·펀딩·경력)은 '시작' 시점에 놓는다
//   · 월을 알 수 없는 항목은 그 연도 그룹의 끝에 둔다 (month = 0)

import {
  publications, conferences, invitedTalks, projects, patents,
  teaching, career, scholarships, books,
  type Bi,
} from './index';

/** 연표 수록 기준 연도 (이 해 1월 이후 실적만) */
export const TIMELINE_FROM = 2023;

export type TimelineCat =
  | 'publication' | 'book' | 'conference' | 'invitedTalk'
  | 'teaching' | 'funding' | 'project' | 'patent' | 'position';

export type TimelineItem = {
  year: number;
  /** 1–12, 0 = 월 미상 */
  month: number;
  cat: TimelineCat;
  main: Bi;
  meta: Bi;
};

// ───────── primitives ─────────

/** 날짜·기간 문자열에서 (연, 월) 을 읽는다. site.json 에 섞여 있는 표기를 모두 받는다.
 *  '2026/10/28' · '2025.09.17' · '2025/06 — 2027/05' · '2024/09 — Present' · '2016'
 *  기간은 앞쪽(시작 연월)만 쓴다. 월을 읽을 수 없으면 month = 0. */
function ym(v: string | null | undefined): { y: number; m: number } {
  const s = String(v ?? '').trim();
  const full = s.match(/^(\d{4})\s*[./-]\s*(\d{1,2})\b/);
  if (full) return { y: Number(full[1]), m: Number(full[2]) };
  const yearOnly = s.match(/^(\d{4})/);
  return { y: yearOnly ? Number(yearOnly[1]) : 0, m: 0 };
}

/** ko 가 비면 en 으로, en 이 비면 ko 로 메운다 (한쪽만 있는 항목이 사라지지 않게). */
const bi = (ko: string | null | undefined, en: string | null | undefined): Bi => {
  const k = String(ko ?? '').trim();
  const e = String(en ?? '').trim();
  return { ko: k || e, en: e || k };
};

/** 여러 Bi 를 구분자로 잇는다. 빈 조각은 구분자째로 빠진다. */
const joinBi = (parts: (Bi | null | undefined)[], sep = ' · '): Bi => ({
  ko: parts.map((p) => (p?.ko || p?.en || '').trim()).filter(Boolean).join(sep),
  en: parts.map((p) => (p?.en || p?.ko || '').trim()).filter(Boolean).join(sep),
});

// ───────── collection ─────────

const items: TimelineItem[] = [];

/** 연표에 한 줄 추가. TIMELINE_FROM 이전 항목과 연도 미상 항목은 조용히 빠진다. */
function push(
  dateLike: string | null | undefined,
  cat: TimelineCat,
  main: Bi,
  meta: Bi,
  fallbackYear = 0,
) {
  const { y, m } = ym(dateLike);
  const year = y || Number(fallbackYear) || 0;
  if (!year || year < TIMELINE_FROM) return;
  if (!main.ko && !main.en) return;
  items.push({ year, month: m, cat, main, meta });
}

// 논문 — 국문 페이지는 KCI 국문 저널명(journalKo)을 쓴다 (본문 카드와 같은 선택).
for (const p of publications) {
  push(p.date, 'publication',
    bi(p.title.ko, p.title.en),
    bi(p.journalKo || p.journal, p.journal),
    p.year);
}

// 저서 — site.json 에 월이 없는 항목이 있다 → 월 미상으로 그 해 끝에 간다.
for (const b of books) {
  push(b.date, 'book', bi(b.title.ko, b.title.en), bi(b.publisher.ko, b.publisher.en));
}

// 초청 강연 — date 가 비고 year 만 있는 항목이 있다.
for (const t of invitedTalks) {
  push(t.date, 'invitedTalk', bi(t.title.ko, t.title.en), bi(t.host.ko, t.host.en), t.year);
}

// 학회 발표 — 웹 본문은 국제·국내를 모두 싣는다. 연표도 같게 한다.
for (const c of conferences) {
  const conf: Bi = typeof c.conference === 'string'
    ? bi(c.conference, c.conference)
    : bi(c.conference.ko, c.conference.en);
  push(c.date, 'conference',
    bi(c.title.ko, c.title.en),
    joinBi([conf, bi(c.venue.ko, c.venue.en)]),
    c.year);
}

// 강의 — 한 과목(또는 한 특강)을 1건으로, 시작일에 놓는다.
for (const t of teaching) {
  push(t.date, 'teaching', bi(t.title.ko, t.title.en), bi(t.host.ko, t.host.en));
}

// 펀딩 — 수여·개시 시점(기간 앞쪽).
for (const s of scholarships) {
  push(s.period, 'funding', bi(s.name.ko, s.name.en), bi(s.foundation.ko, s.foundation.en));
}

// 연구 과제 — 과제 시작 시점. 본인 책임과제는 표기를 덧붙인다 (Home.astro 의 PI 판별과 같은 규칙).
for (const p of projects) {
  const isPi = (p.role?.en || '').includes('Principal') || (p.role?.ko || '').includes('책임');
  push(p.period, 'project',
    bi(p.title.ko + (isPi ? ' (연구 책임)' : ''), p.title.en + (isPi ? ' (PI)' : '')),
    bi(p.institute.ko, p.institute.en));
}

// 특허 — 등록건은 등록일, 출원중은 출원일.
for (const pt of patents) {
  const registered = pt.status === 'registered';
  const stage: Bi = registered ? { ko: '등록', en: 'Registered' } : { ko: '출원', en: 'Pending' };
  push(registered ? pt.registrationDate : pt.applicationDate, 'patent',
    bi(pt.title.ko, pt.title.en),
    joinBi([stage, bi(pt.owner.ko, pt.owner.en)]));
}

// 직위 — 부임 시점. main 은 직위명, meta 는 소속 (PDF 판과 같은 배치).
for (const w of career) {
  push(w.period, 'position', bi(w.type.ko, w.type.en), bi(w.company.ko, w.company.en));
}

// ───────── ordering ─────────

// 같은 달에 여러 건이 있을 때의 표시 순서 = 웹 페이지의 섹션 순서
// (경력 03 → 논문·저서 04 → 초청강연·학회발표 05 → 강의 06 → 펀딩 07 → 과제 08 → 특허 09)
const CAT_ORDER: TimelineCat[] = [
  'position', 'publication', 'book', 'invitedTalk', 'conference',
  'teaching', 'funding', 'project', 'patent',
];
const catRank = (c: TimelineCat) => {
  const i = CAT_ORDER.indexOf(c);
  return i < 0 ? CAT_ORDER.length : i;
};

/** 전체 연표 항목 (정렬 전 원본) */
export const timeline: TimelineItem[] = items;

/** 연도 역순 그룹. 연도 안에서는 월 내림차순, 월 미상(0)은 그 해의 끝. */
export const timelineByYear: { year: number; items: TimelineItem[] }[] =
  Array.from(new Set(items.map((i) => i.year)))
    .sort((a, b) => b - a)
    .map((year) => ({
      year,
      items: items
        .filter((i) => i.year === year)
        .sort((a, b) =>
          // 월 미상(0)은 -99 로 바꿔 내림차순에서 맨 뒤로 보낸다
          (b.month || -99) - (a.month || -99)
          || catRank(a.cat) - catRank(b.cat)
          || a.main.en.localeCompare(b.main.en)),
    }));

/** 카테고리별 건수 (라벨 옆 표기·점검용) */
export const timelineCounts: Record<string, number> = CAT_ORDER.reduce(
  (acc, c) => { acc[c] = items.filter((i) => i.cat === c).length; return acc; },
  {} as Record<string, number>,
);

const MONTHS_KO = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
const MONTHS_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** 월 표기. 월 미상은 빈 문자열 (칸만 비워 둔다 — 값을 지어내지 않는다). */
export const monthLabel = (m: number): Bi =>
  m >= 1 && m <= 12 ? { ko: MONTHS_KO[m - 1], en: MONTHS_EN[m - 1] } : { ko: '', en: '' };
