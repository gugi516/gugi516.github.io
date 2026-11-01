// 사이트 메타데이터. 실제 텍스트는 src/data/site.json 에서 관리됩니다.
import site from './data/site.json';

export const SITE_TITLE_KO = site.site.ko.title;
export const SITE_TITLE_EN = site.site.en.title;
export const SITE_DESCRIPTION_KO = site.site.ko.description;
export const SITE_DESCRIPTION_EN = site.site.en.description;

// 기본값 (서버사이드 렌더링용). 클라이언트 JS가 paint 직전에 html lang 속성을 갱신.
export const SITE_TITLE = SITE_TITLE_EN;
export const SITE_DESCRIPTION = SITE_DESCRIPTION_EN;

export const TRANSITION_API = false;
