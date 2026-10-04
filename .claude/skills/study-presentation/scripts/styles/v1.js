// 스타일 버전 1 (v1) — "토스 블루": 흰 배경 + Signal Blue, Pretendard, Darcula 코드
// 2026-10-04 고정. 스프링 입문 발표(피드백 3회 반영)로 확정한 규격이다.
// 색·폰트·크기·좌표는 전부 여기서만 정의한다. 값을 바꾸면 references/styles/v1.md 변경 이력도 함께 갱신.
// 크게 다른 스타일이 필요하면 이 파일을 고치지 말고 styles/v2.js 를 새로 만든다.

const COLOR = {
  // 기본 팔레트 (Signal Blue / Porcelain 조합, 흰 배경)
  bg: "FFFFFF",
  primary: "0057FF",     // Signal Blue — 유일한 강조색
  primaryTint: "EBF2FF", // Signal Blue 10% 틴트 — 태그 칩 배경
  ink: "23262F",         // Graphite — 본문/제목 텍스트
  sub: "6B7280",         // 보조 텍스트 (흰 배경 대비 4.8:1)
  card: "F8F7F4",        // Porcelain — 카드 면
  onPrimary: "FFFFFF",   // 파란 카드 위 텍스트
  warn: "FF4696",        // Dragonfruit — 경고/함정. 18pt 이상 굵은 글씨에만, 덱당 1~2회
  codeBg: "23262F",      // Graphite — 코드 블록 배경
  codeLabel: "6B7280",
};

// IntelliJ Darcula 토큰 색
const CODE = {
  plain: "A9B7C6",
  keyword: "CC7832",
  number: "6897BB",
  string: "6A8759",
  comment: "808080",
  method: "FFC66D",
  annotation: "BBB529",
  tag: "E8BF6A",
  attr: "BABABA",
  attrValue: "A5C261",
};

const FONT = {
  head: "Pretendard",
  body: "Pretendard",
  code: "Consolas",
};

// pt 단위. 최소 크기: 본문 14 / 캡션 12 / 코드 14(온라인 화면공유 기준)
const SIZE = {
  coverTitle: 40,
  title: 32,
  statement: 40,
  principle: 36,
  eyebrow: 14,
  sectionHead: 20,
  cardHead: 17,
  body: 16,
  bodySmall: 14,
  caption: 12,
  keyline: 18,
  bigNumber: 40,
  quiz: 34,            // 퀴즈 문제(2줄까지)
  quizReason: 18,      // 퀴즈 이유
  quizCodeMax: 20,     // 퀴즈 코드는 크게(최대)
  lead: 22,            // 슬라이드 상단 질문 띠
  codeMax: 16,
  codeMin: 14,         // 온라인(디스코드) 화면공유 기준. 오프라인 발표여도 14 유지
};

// 인치 단위. 16:9 와이드 13.333 x 7.5
const GRID = {
  w: 13.333,
  h: 7.5,
  mx: 0.75,                 // 좌우 여백
  cw: 13.333 - 1.5,         // 콘텐츠 폭 11.833
  eyebrowY: 0.5,
  titleY: 0.85,
  titleH: 0.85,
  top: 1.95,                // 본문 시작
  bottom: 6.75,             // 본문 끝 (하단 여백 0.75)
  gap: 0.3,                 // 블록 간격
  radius: 0.12,             // 라운드 반경
  codePad: 0.2,             // 코드 블록 안쪽 여백
};

module.exports = { COLOR, CODE, FONT, SIZE, GRID };
