#!/usr/bin/env node
// 스터디 발표자료 생성기
//   node build.js <deck.json>
// 하나의 덱 스펙(JSON)에서 발표자료(.pptx, 노트에 대사 전문)와 발표대본(.md)을 함께 만든다.
// 스펙 형식은 ../references/deck-spec.md 참고.

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");

// ───────────────────────── 입력 ─────────────────────────
const specArg = process.argv[2];
if (!specArg) {
  console.error("usage: node build.js <deck.json>");
  process.exit(1);
}
const specPath = path.resolve(specArg);
const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
const meta = spec.meta || {};

// 스타일 버전(meta.style, 기본 v1)을 고른 뒤 토큰을 읽는다 — highlight.js 도 같은 버전을 쓴다
process.env.STUDY_STYLE = meta.style || "v1";
const { VERSION: STYLE_VERSION, COLOR, FONT, SIZE, GRID } = require("./style");
const { highlight } = require("./highlight");
const { makeShare } = require("./share");
const out = spec.output || {};
const outDir = path.resolve(path.dirname(specPath), out.dir || ".");
const pptxName = `${out.name || "발표자료"}.pptx`;
const pptxPath = path.join(outDir, pptxName);
const scriptPath = path.join(outDir, out.script || "발표대본.md");
const warnings = [];
let currentSlide = 0;
const warn = (msg) => warnings.push(typeof currentSlide === "number" ? `[${currentSlide}번] ${msg}` : `[${currentSlide}] ${msg}`);

// ───────────────────────── 측정 ─────────────────────────
// 줄바꿈·넘침 판단용 근사 폭(em). Pretendard 기준, 5% 여유를 둔다.
function units(str) {
  let u = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    if ((c >= 0x1100 && c <= 0x11ff) || (c >= 0x2e80 && c <= 0xd7af) || (c >= 0xf900 && c <= 0xfaff) || (c >= 0xff00 && c <= 0xffef)) u += 0.95;
    else if (ch === " ") u += 0.28;
    else if (/[A-Z0-9]/.test(ch)) u += 0.62;
    else if (/[a-z]/.test(ch)) u += 0.53;
    else u += 0.4;
  }
  return u;
}
function codeUnits(str) {
  let u = 0;
  for (const ch of str) u += ch.codePointAt(0) > 0x2e80 ? 1.0 : 0.55;
  return u;
}
const plain = (s) => String(s ?? "").replace(/\*\*/g, "").replace(/`/g, "");
// 어절(공백) 단위 줄바꿈 시뮬레이션 — 긴 인라인 코드가 통째로 다음 줄로 넘어가는 것까지 반영
function lineCount(text, pt, wIn) {
  const perLine = ((wIn * 72) / pt) * 0.97; // Pretendard 실측(한글 0.95em) 기준 3% 여유
  const SPACE = 0.28;
  let total = 0;
  for (const para of String(text ?? "").split("\n")) {
    const words = [];
    for (const part of para.split(/(`[^`]+`)/)) {
      if (!part) continue;
      const isCode = /^`[\s\S]*`$/.test(part);
      const body = isCode ? part.slice(1, -1) : part.replace(/\*\*/g, "");
      for (const w of body.split(" ")) if (w) words.push(isCode ? codeUnits(w) : units(w));
    }
    let lines = 1, cur = 0;
    for (const u of words) {
      if (cur > 0 && cur + SPACE + u > perLine) { lines += 1; cur = 0; }
      cur += (cur > 0 ? SPACE : 0) + u;
      while (cur > perLine) { lines += 1; cur -= perLine; }
    }
    total += lines;
  }
  return total;
}
const textH = (text, pt, wIn, lh = 1.35) => (lineCount(text, pt, wIn) * pt * lh) / 72;

// ───────────────────────── 텍스트 ─────────────────────────
// **강조** → 강조 스타일, `코드` → 코드 폰트, \n → 줄바꿈
function rich(str, base, emph) {
  const runs = [];
  const lines = String(str ?? "").split("\n");
  lines.forEach((ln, li) => {
    const parts = ln.split(/(\*\*[^*]+\*\*|`[^`]+`)/).filter((p) => p !== "");
    if (!parts.length) parts.push("");
    parts.forEach((p, pi) => {
      let text = p;
      const o = { ...base };
      if (/^\*\*[\s\S]*\*\*$/.test(p)) { text = p.slice(2, -2); Object.assign(o, emph); }
      else if (/^`[\s\S]*`$/.test(p)) { text = p.slice(1, -1); o.fontFace = FONT.code; }
      if (pi === parts.length - 1 && li < lines.length - 1) o.breakLine = true;
      runs.push({ text, options: o });
    });
  });
  return runs;
}
const LANG = "ko-KR"; // 한글 줄바꿈 규칙·동아시아 글꼴 적용
const EMPH = { color: COLOR.primary, bold: true };
const EMPH_ON_PRIMARY = { color: COLOR.onPrimary, bold: true, underline: { style: "sng" } };

function text(slide, str, o) {
  const { color = COLOR.ink, size = SIZE.body, bold = false, emph = EMPH, fontFace = FONT.body, ...rest } = o;
  slide.addText(rich(str, { color, fontSize: size, bold, fontFace, lang: LANG }, emph), {
    isTextBox: true, margin: 0, valign: "top", fontFace, fontSize: size, color, lang: LANG, lineSpacingMultiple: 1.1, ...rest,
  });
}

function bullets(slide, items, { x, y, w, size = SIZE.body, color = COLOR.ink }) {
  const runs = [];
  items.forEach((item, i) => {
    const r = rich(item, { color, fontSize: size, fontFace: FONT.body, lang: LANG }, EMPH);
    r[0].options.bullet = { indent: 16 };
    const last = r[r.length - 1].options;
    if (i < items.length - 1) last.breakLine = true;
    runs.push(...r);
  });
  const h = bulletsH(items, size, w);
  slide.addText(runs, { x, y, w, h, isTextBox: true, margin: 0, valign: "top", fontFace: FONT.body, fontSize: size, color, lang: LANG, paraSpaceAfter: 8, lineSpacingMultiple: 1.1 });
  return h;
}
const bulletsH = (items, size, w) => items.reduce((h, it) => h + textH(it, size, w - 0.25) + 8 / 72, 0);

function eyebrow(slide, label, accent) {
  if (!label) return;
  slide.addText(label, {
    x: GRID.mx, y: GRID.eyebrowY, w: GRID.cw, h: 0.32, isTextBox: true, margin: 0, valign: "middle", lang: LANG,
    fontFace: FONT.body, fontSize: SIZE.eyebrow, bold: true, color: accent ? COLOR.primary : COLOR.sub,
  });
}

function title(slide, str, { size = SIZE.title, width = GRID.cw, maxLines = 1 } = {}) {
  let pt = size;
  while (lineCount(str, pt, width) > maxLines && pt > size - 4) pt -= 2;
  if (lineCount(str, pt, width) > maxLines) warn(`제목이 ${maxLines}줄을 넘습니다 — 줄이세요: "${plain(str)}"`);
  slide.addText(rich(str, { color: COLOR.ink, fontSize: pt, bold: true, lang: LANG, align: "left" }, { color: COLOR.primary }), { placeholder: "title", align: "left" });
}

function card(slide, x, y, w, h, fill = COLOR.card) {
  slide.addShape("roundRect", { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: GRID.radius });
}

// 결론 박스: style "card"(Porcelain) | "accent"(Signal Blue) | "warn"(Dragonfruit 글씨)
function callout(slide, c, x, y, w) {
  const o = typeof c === "string" ? { text: c } : c;
  const style = o.style || "card";
  const padX = 0.4, padY = 0.28;
  const labelH = o.label ? 0.36 : 0;
  const bodyH = textH(o.text, SIZE.keyline, w - 2 * padX, 1.45);
  const h = labelH + bodyH + 2 * padY;
  const onPrimary = style === "accent";
  card(slide, x, y, w, h, onPrimary ? COLOR.primary : COLOR.card);
  if (o.label) text(slide, o.label, { x: x + padX, y: y + padY, w: w - 2 * padX, h: 0.3, size: SIZE.bodySmall, bold: true, color: onPrimary ? COLOR.onPrimary : COLOR.sub });
  text(slide, o.text, {
    x: x + padX, y: y + padY + labelH, w: w - 2 * padX, h: bodyH, size: SIZE.keyline, bold: true,
    color: onPrimary ? COLOR.onPrimary : style === "warn" ? COLOR.warn : COLOR.ink,
    emph: onPrimary ? EMPH_ON_PRIMARY : EMPH,
  });
  return h;
}
const calloutH = (c, w) => {
  const o = typeof c === "string" ? { text: c } : c;
  return (o.label ? 0.36 : 0) + textH(o.text, SIZE.keyline, w - 0.8, 1.45) + 0.56;
};

// ───────────────────────── 코드 블록 ─────────────────────────
const CODE_LH = 1.4; // 줄 높이(em) — lineSpacingMultiple 1.2 기준, 실측 약 1.36
const LABEL_H = 0.34;
const MAX_CODE_LINES = 12; // 온라인(디스코드) 화면공유 압축 대비
const normCode = (b) => (b == null ? null : typeof b === "string" ? { code: b, lang: "java" } : { lang: "java", ...b });
// code(문자열·객체·배열) + output(실행 결과)을 위에서 아래로 쌓을 블록 목록으로
const codeBlocks = (s) => [
  ...[].concat(s.code ?? []).map(normCode),
  s.output ? { lang: "output", ...(typeof s.output === "string" ? { code: s.output } : s.output) } : null,
].filter(Boolean);
const codeLines = (b) => highlight(b.code, b.lang);
const codeBoxH = (lines, pt) => (lines.length * pt * CODE_LH) / 72 + 2 * GRID.codePad;
const codeWidthOk = (lines, pt, w) => Math.max(...lines.map((l) => codeUnits(l.map((t) => t.text).join("")))) * pt / 72 <= w - 2 * GRID.codePad;

// 여러 코드 블록을 같은 글자 크기로 세로로 쌓을 수 있는 최대 pt
function fitStack(blocks, w, H, maxPt = 16) {
  blocks.forEach((b) => { const n = codeLines(b).length; if (n > MAX_CODE_LINES) warn(`코드 블록이 ${n}줄입니다 — 온라인 발표는 블록당 ${MAX_CODE_LINES}줄 이하(핵심 줄만 남기세요)`); });
  for (let pt = maxPt; pt >= SIZE.codeMin; pt -= 1) {
    const total = blocks.reduce((s, b) => s + (b.label ? LABEL_H : 0) + codeBoxH(codeLines(b), pt), 0) + GRID.gap * (blocks.length - 1);
    if (total <= H + 0.05 && blocks.every((b) => codeWidthOk(codeLines(b), pt, w))) return pt; // 0.05in 허용(실측 여유)
  }
  warn(`코드가 영역(${w.toFixed(1)}×${H.toFixed(1)}in)에 ${SIZE.codeMin}pt로도 들어가지 않습니다 — 코드를 줄이거나 슬라이드를 나누세요`);
  return SIZE.codeMin;
}

function drawCode(slide, b, x, y, w, pt) {
  let top = y;
  if (b.label) {
    text(slide, b.label, { x, y: top, w, h: 0.3, size: SIZE.caption, bold: true, color: COLOR.codeLabel });
    top += LABEL_H;
  }
  const lines = codeLines(b);
  const h = codeBoxH(lines, pt);
  card(slide, x, top, w, h, COLOR.codeBg);
  const runs = [];
  lines.forEach((line, li) => {
    const toks = line.length ? line : [{ text: "", color: "A9B7C6" }];
    toks.forEach((t, ti) => {
      const o = { color: t.color, fontFace: FONT.code, fontSize: pt };
      if (ti === toks.length - 1 && li < lines.length - 1) o.breakLine = true;
      runs.push({ text: t.text, options: o });
    });
  });
  slide.addText(runs, {
    x: x + GRID.codePad, y: top + GRID.codePad, w: w - 2 * GRID.codePad, h: h - 2 * GRID.codePad,
    isTextBox: true, margin: 0, valign: "top", fontFace: FONT.code, fontSize: pt, lineSpacingMultiple: 1.2,
  });
  return top + h - y;
}

function drawStack(slide, blocks, x, y, w, H, maxPt) {
  const pt = fitStack(blocks, w, H, maxPt);
  let cy = y;
  blocks.forEach((b, i) => { cy += drawCode(slide, b, x, cy, w, pt) + (i < blocks.length - 1 ? GRID.gap : 0); });
  return cy - y;
}

// ───────────────────────── 슬라이드 유형 ─────────────────────────
const { mx, cw, top: TOP, bottom: BOTTOM, gap: GAP } = GRID;
const RIGHT = mx + cw;

const RENDER = {
  // 표지 겸 목차
  cover(pres, s) {
    const slide = pres.addSlide({ masterName: "STUDY_COVER" });
    eyebrow(slide, s.eyebrow || meta.course);
    title(slide, s.title, { size: SIZE.coverTitle, maxLines: 1 });
    if (s.subtitle) text(slide, s.subtitle, { x: mx, y: 2.15, w: cw, h: 0.45, size: 18, color: COLOR.sub });
    const items = s.agenda || [];
    if (items.length) {
      const n = items.length;
      const w = (cw - GAP * (n - 1)) / n;
      const y = s.subtitle ? 3.0 : 2.45, h = 2.75; // 부제가 없으면 카드를 올린다
      items.forEach((it, i) => {
        const x = mx + i * (w + GAP);
        card(slide, x, y, w, h);
        text(slide, it.no, { x: x + 0.35, y: y + 0.3, w: w - 0.7, h: 0.7, size: SIZE.bigNumber, bold: true, color: COLOR.primary, fontFace: FONT.head });
        text(slide, it.title, { x: x + 0.35, y: y + 1.15, w: w - 0.7, h: 0.45, size: SIZE.sectionHead, bold: true });
        if (it.desc) text(slide, it.desc, { x: x + 0.35, y: y + 1.65, w: w - 0.7, h: 0.95, size: SIZE.bodySmall, color: COLOR.sub });
        if (textH(it.desc || "", SIZE.bodySmall, w - 0.7) > 0.95) warn(`목차 설명이 깁니다: "${it.desc}"`);
      });
    }
    if (s.extra) text(slide, s.extra, { x: mx, y: 6.0, w: cw, h: 0.4, size: SIZE.body, bold: true, color: COLOR.sub });
    const footer = [meta.presenter, meta.date].filter(Boolean).join("  ·  ");
    if (footer) text(slide, footer, { x: mx, y: 6.7, w: cw, h: 0.3, size: SIZE.caption, color: COLOR.sub });
    return slide;
  },

  // 큰 한 문장 (질문·훅·전환)
  statement(pres, s) {
    const slide = pres.addSlide({ masterName: "STUDY_STATEMENT" });
    eyebrow(slide, s.eyebrow, s.accent);
    title(slide, s.text || s.title, { size: SIZE.statement, maxLines: 2 });
    if (s.sub) text(slide, s.sub, { x: mx, y: 4.55, w: cw, h: 0.9, size: 18, color: COLOR.sub });
    return slide;
  },

  // before / after 코드 비교
  "code-compare"(pres, s) {
    const slide = content(pres, s);
    const TOP = bodyTop(s);
    const colW = (cw - 0.4) / 2;
    const left = normCode(s.left), right = normCode(s.right);
    let bottom = BOTTOM;
    if (s.callout) { const h = calloutH(s.callout, cw); callout(slide, s.callout, mx, BOTTOM - h, cw); bottom -= h + GAP; }
    const pts = s.points || [];
    const ptsH = pts.length ? Math.max(...pts.map((p) => 0.4 + textH(p.body || "", SIZE.bodySmall, colW))) : 0;
    const codeArea = bottom - TOP - (ptsH ? ptsH + GAP : 0);
    const pt = Math.min(fitStack([left], colW, codeArea), fitStack([right], colW, codeArea));
    const used = Math.max(drawCode(slide, left, mx, TOP, colW, pt), drawCode(slide, right, mx + colW + 0.4, TOP, colW, pt));
    pts.forEach((p, i) => {
      const x = mx + i * (colW + 0.4), y = TOP + used + GAP;
      text(slide, p.head, { x, y, w: colW, h: 0.36, size: SIZE.cardHead, bold: true });
      if (p.body) text(slide, p.body, { x, y: y + 0.4, w: colW, h: ptsH - 0.4, size: SIZE.bodySmall, color: COLOR.sub });
    });
    return slide;
  },

  // 코드(+실행 결과) 왼쪽, 설명 오른쪽
  "code-explain"(pres, s) {
    const slide = content(pres, s);
    const TOP = bodyTop(s);
    let bottom = BOTTOM;
    if (s.callout) { const h = calloutH(s.callout, cw); callout(slide, s.callout, mx, BOTTOM - h, cw); bottom -= h + GAP; }
    const leftW = 6.9, rx = mx + leftW + 0.45, rw = RIGHT - rx;
    const blocks = codeBlocks(s);
    const used = drawStack(slide, blocks, mx, TOP, leftW, bottom - TOP);
    rightColumn(slide, s, rx, TOP + 0.05, rw, bottom, used - 0.05);
    return slide;
  },

  // 실무 사례: 설명 왼쪽, 코드 오른쪽, 아래 질문 박스
  case(pres, s) {
    const slide = content(pres, { eyebrow: "실무", accent: true, ...s });
    const leftW = 6.0, rx = mx + leftW + 0.45, rw = RIGHT - rx;
    let leftBottom = BOTTOM;
    if (s.question) {
      const q = { label: s.question.q, text: s.question.a, style: "card" };
      const h = calloutH(q, leftW); callout(slide, q, mx, BOTTOM - h, leftW); leftBottom -= h + GAP;
    }
    rightColumn(slide, s, mx, TOP + 0.05, leftW, leftBottom);
    const blocks = codeBlocks(s);
    if (blocks.length) drawStack(slide, blocks, rx, TOP, rw, BOTTOM - TOP);
    return slide;
  },

  // (선택) 코드 위, 개념 카드 2~4개 아래
  cards(pres, s) {
    const slide = content(pres, s);
    const TOP = bodyTop(s);
    const items = s.cards || [];
    const n = Math.max(items.length, 1);
    const w = (cw - GAP * (n - 1)) / n;
    const bodySize = n <= 3 ? SIZE.body : SIZE.bodySmall; // 카드가 3장 이하면 본문을 키운다
    const cardH = Math.max(1.3, ...items.map((c) => 0.95 + textH(c.body || "", bodySize, w - 0.6)));
    const keyH = s.keyline ? textH(s.keyline, SIZE.keyline, cw) + 0.1 : 0;
    const cardsY = BOTTOM - keyH - (keyH ? GAP : 0) - cardH;
    if (s.code) drawStack(slide, [normCode(s.code)], mx, TOP, cw, cardsY - GAP - TOP);
    const y = s.code ? cardsY : TOP;
    items.forEach((c, i) => {
      const x = mx + i * (w + GAP);
      const fg = c.accent ? COLOR.onPrimary : COLOR.ink; // accent: 권장·정답처럼 강조할 카드를 파란 카드로
      card(slide, x, y, w, cardH, c.accent ? COLOR.primary : COLOR.card);
      text(slide, c.head, { x: x + 0.3, y: y + 0.28, w: w - 0.6, h: 0.4, size: SIZE.cardHead, bold: true, color: fg, emph: c.accent ? EMPH_ON_PRIMARY : EMPH });
      if (c.body) text(slide, c.body, { x: x + 0.3, y: y + 0.78, w: w - 0.6, h: cardH - 0.95, size: bodySize, color: c.accent ? COLOR.onPrimary : COLOR.sub, emph: c.accent ? EMPH_ON_PRIMARY : EMPH });
    });
    if (s.keyline) text(slide, s.keyline, { x: mx, y: (s.code ? cardsY : TOP) + cardH + GAP, w: cw, h: keyH, size: SIZE.keyline, bold: true });
    return slide;
  },

  // 둘을 대비: 왼쪽 Porcelain 카드, 오른쪽 Signal Blue 카드
  contrast(pres, s) {
    const slide = content(pres, s);
    const w = (cw - 0.35) / 2;
    const codeB = normCode(s.code);
    const codeH = codeB ? Math.min(2.0, codeBoxH(codeLines(codeB), 14) + (codeB.label ? LABEL_H : 0)) : 0;
    const h = BOTTOM - TOP - (codeB ? codeH + GAP : 0);
    [s.left, s.right].forEach((side, i) => {
      if (!side) return;
      const x = mx + i * (w + 0.35);
      const accent = i === 1;
      const fg = accent ? COLOR.onPrimary : COLOR.ink;
      const fg2 = accent ? COLOR.onPrimary : COLOR.sub;
      card(slide, x, TOP, w, h, accent ? COLOR.primary : COLOR.card);
      let y = TOP + 0.35;
      text(slide, side.head, { x: x + 0.4, y, w: w - 0.8, h: 0.45, size: SIZE.sectionHead, bold: true, color: fg, emph: accent ? EMPH_ON_PRIMARY : EMPH });
      y += 0.6;
      if (side.sub) { text(slide, side.sub, { x: x + 0.4, y, w: w - 0.8, h: 0.4, size: 15, color: fg, fontFace: FONT.code }); y += 0.6; }
      const body = side.body || [];
      if (body.length) text(slide, body.join("\n"), { x: x + 0.4, y, w: w - 0.8, h: TOP + h - y - 0.3, size: SIZE.body, color: fg2, paraSpaceAfter: 6, emph: accent ? EMPH_ON_PRIMARY : EMPH });
    });
    if (codeB) drawStack(slide, [codeB], mx, BOTTOM - codeH, cw, codeH, 15);
    return slide;
  },

  // 핵심 원칙: 큰 문장 왼쪽, 코드·결과 오른쪽, 아래 파란 결론 카드
  principle(pres, s) {
    const slide = pres.addSlide({ masterName: "STUDY_PRINCIPLE" });
    eyebrow(slide, s.eyebrow);
    title(slide, s.title, { size: SIZE.principle, width: 5.9, maxLines: 3 });
    let bottom = BOTTOM;
    if (s.keyline) { const c = { text: s.keyline, style: "accent" }; const h = calloutH(c, cw); callout(slide, c, mx, BOTTOM - h, cw); bottom -= h + GAP; }
    if (s.quote) text(slide, s.quote, { x: mx, y: 3.75, w: 5.9, h: 0.5, size: 18, color: COLOR.sub });
    const rx = mx + 6.3, rw = RIGHT - rx;
    const blocks = codeBlocks(s);
    if (blocks.length) drawStack(slide, blocks, rx, 1.2, rw, bottom - 1.2);
    return slide;
  },

  // 그림(강의자료에서 잘라낸 다이어그램 등) + 설명
  image(pres, s) {
    const slide = content(pres, s);
    const TOP = bodyTop(s);
    const img = s.image || {};
    const imgPath = path.resolve(path.dirname(specPath), img.path || "");
    if (!img.path || !fs.existsSync(imgPath)) { warn(`이미지를 찾을 수 없습니다: ${img.path}`); return slide; }
    let bottom = BOTTOM;
    if (s.callout) { const h = calloutH(s.callout, cw); callout(slide, s.callout, mx, BOTTOM - h, cw); bottom -= h + GAP; }
    const full = !(s.bullets && s.bullets.length);
    const w = full ? cw : 7.4;
    const capH = img.caption ? 0.4 : 0;
    const h = bottom - TOP - capH;
    card(slide, mx, TOP, w, h);
    const box = { x: mx + 0.2, y: TOP + 0.2, w: w - 0.4, h: h - 0.4 };
    const px = imageSize(imgPath);
    const scale = px ? Math.min(box.w / px.w, box.h / px.h) : null;
    const iw = scale ? px.w * scale : box.w, ih = scale ? px.h * scale : box.h;
    if (!px) warn(`이미지 크기를 읽지 못해 영역에 맞춰 늘렸습니다(PNG/JPG 권장): ${img.path}`);
    slide.addImage({ path: imgPath, x: box.x + (box.w - iw) / 2, y: box.y + (box.h - ih) / 2, w: iw, h: ih, altText: img.alt || plain(s.title) });
    if (img.caption) text(slide, img.caption, { x: mx, y: TOP + h + 0.08, w, h: 0.3, size: SIZE.caption, color: COLOR.sub });
    if (!full) {
      const rx = mx + w + 0.45, rw = RIGHT - rx;
      // 설명이 불릿뿐이면 그림 카드 높이에 고르게 배치, keyline·note 가 있으면 위에서부터 쌓는다
      if (!s.keyline && !s.note) spreadBullets(slide, s.bullets, rx, TOP, rw, h);
      else rightColumn(slide, s, rx, TOP + 0.05, rw, bottom, h - 0.05);
    }
    return slide;
  },

  // 정리: 단원별 한 문장
  summary(pres, s) {
    const slide = content(pres, { title: "정리", ...s });
    const items = s.items || [];
    const closeH = s.closing ? 0.6 : 0;
    const gap = 0.2;
    const rowH = Math.min(1.45, (BOTTOM - TOP - closeH - gap * (items.length - 1)) / Math.max(items.length, 1));
    items.forEach((it, i) => {
      const y = TOP + i * (rowH + gap);
      card(slide, mx, y, cw, rowH);
      text(slide, it.no, { x: mx + 0.4, y, w: 1.3, h: rowH, size: 40, bold: true, color: COLOR.primary, valign: "middle", fontFace: FONT.head });
      const hasSub = !!it.sub;
      text(slide, it.head, { x: mx + 1.9, y: y + (hasSub ? 0.22 : 0), w: cw - 2.2, h: hasSub ? 0.55 : rowH, size: 24, bold: true, valign: hasSub ? "top" : "middle" });
      if (hasSub) text(slide, it.sub, { x: mx + 1.9, y: y + 0.82, w: cw - 2.2, h: rowH - 0.9, size: 17, color: COLOR.sub });
      if (lineCount(it.head, 24, cw - 2.2) > 1) warn(`정리 문장이 1줄을 넘습니다: "${plain(it.head)}"`);
    });
    if (s.closing) text(slide, s.closing, { x: mx, y: BOTTOM - 0.45, w: cw, h: 0.45, size: SIZE.cardHead, bold: true, color: COLOR.sub });
    return slide;
  },

  // 퀴즈 문제 (quiz-item 이 펼쳐진 앞장) — 문제와 코드·선택지만 보여준다
  "quiz-q"(pres, s) {
    const slide = quizFrame(pres, s);
    quizBody(slide, s, false);
    return slide;
  },

  // 퀴즈 정답 — 같은 배치에 정답과 이유를 드러낸다 (넘기면 공개되는 효과)
  "quiz-a"(pres, s) {
    const slide = quizFrame(pres, s);
    const body = quizBody(slide, s, true);
    let ry = body.y;
    if (!body.answered) {
      const ansH = 1.1;
      card(slide, body.x, body.y, body.w, ansH, COLOR.primary);
      text(slide, "정답", { x: body.x + 0.4, y: body.y, w: 1.1, h: ansH, size: SIZE.cardHead, bold: true, color: COLOR.onPrimary, valign: "middle" });
      text(slide, s.answer, { x: body.x + 1.5, y: body.y, w: body.w - 1.9, h: ansH, size: 30, bold: true, color: COLOR.onPrimary, valign: "middle", emph: EMPH_ON_PRIMARY });
      ry += ansH + GAP;
    }
    const rs = SIZE.quizReason;
    const rh = Math.min(BOTTOM - ry, 0.75 + textH(s.reason, rs, body.w - 0.8) + 0.35); // 내용 높이에 맞춤
    card(slide, body.x, ry, body.w, rh);
    text(slide, "이유", { x: body.x + 0.4, y: ry + 0.28, w: body.w - 0.8, h: 0.3, size: SIZE.bodySmall, bold: true, color: COLOR.sub });
    text(slide, s.reason, { x: body.x + 0.4, y: ry + 0.72, w: body.w - 0.8, h: rh - 0.9, size: rs });
    if (0.75 + textH(s.reason, rs, body.w - 0.8) > rh - 0.15) warn("퀴즈 이유 설명이 영역을 넘칩니다 — 문장을 줄이세요(자세한 설명은 대사로)");
    return slide;
  },
};

// 퀴즈 공통 틀: 소제목 "퀴즈 n / N · 유형", 제목 = 문제 문장
function quizFrame(pres, s) {
  const slide = pres.addSlide({ masterName: "STUDY_QUIZ" });
  eyebrow(slide, `퀴즈 ${s.qno} / ${s.qtotal}   ·   ${s.kind || "단답"}`, true);
  title(slide, s.q, { size: SIZE.quiz, maxLines: 2 });
  return slide;
}

// 퀴즈 본문. 문제·정답 슬라이드가 같은 배치가 되도록 같은 함수로 그린다.
//  - 코드가 있으면: 왼쪽에 코드(최대 20pt), 오른쪽이 답 영역
//  - O/X 이고 코드가 없으면: O·X 큰 카드 두 장. 정답 슬라이드에서는 정답 카드가 파랗게 바뀐다
//  - 선택지가 있으면: 왼쪽에 선택지, 오른쪽이 답 영역 / 그 외: 전체 폭이 답 영역
function quizBody(slide, s, reveal) {
  const y = 2.45;
  const leftW = 6.9, rx = mx + leftW + 0.45, rw = RIGHT - rx;
  const blocks = codeBlocks(s);
  if (blocks.length) { drawStack(slide, blocks, mx, y, leftW, BOTTOM - y, SIZE.quizCodeMax); return { x: rx, y, w: rw }; }
  if (/^O\s*\/\s*X$/i.test(s.kind || "")) {
    const w = (cw - GAP) / 2, h = 2.3;
    ["O", "X"].forEach((v, i) => {
      const hit = reveal && String(s.answer).trim().toUpperCase() === v;
      const x = mx + i * (w + GAP);
      card(slide, x, y, w, h, hit ? COLOR.primary : COLOR.card);
      text(slide, v, { x, y, w, h, size: 80, bold: true, align: "center", valign: "middle", color: hit ? COLOR.onPrimary : reveal ? COLOR.sub : COLOR.ink, fontFace: FONT.head });
    });
    return { x: mx, y: y + h + GAP, w: cw, answered: true };
  }
  if ((s.choices || []).length) {
    text(slide, s.choices.join("\n"), { x: mx, y: y + 0.1, w: leftW, h: BOTTOM - y, size: 22, paraSpaceAfter: 10 });
    return { x: rx, y, w: rw };
  }
  return { x: mx, y, w: cw };
}

// quiz-item 하나를 [문제 슬라이드, 정답 슬라이드] 두 장으로 펼친다
function expandSlides(list) {
  const total = list.filter((s) => s.type === "quiz-item").length;
  let n = 0;
  return list.flatMap((s) => {
    if (s.type !== "quiz-item") return [s];
    n += 1;
    const base = { ...s, qno: n, qtotal: total };
    return [
      { ...base, type: "quiz-q", label: s.label || `퀴즈 ${n} — 문제`, time: s.time || "0:30", script: s.script || [], cues: s.cues || [] },
      { ...base, type: "quiz-a", label: `퀴즈 ${n} — 정답·이유`, tag: s.answerTag, time: s.answerTime || "0:40", script: s.answerScript || [], cues: s.answerCues || [] },
    ];
  });
}

// PNG/JPEG 헤더에서 픽셀 크기를 읽는다 (비율 유지용)
function imageSize(file) {
  const b = fs.readFileSync(file);
  if (b.length > 24 && b.readUInt32BE(0) === 0x89504e47) return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i += 1; continue; }
      const marker = b[i + 1];
      const len = b.readUInt16BE(i + 2);
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) return { w: b.readUInt16BE(i + 7), h: b.readUInt16BE(i + 5) };
      i += 2 + len;
    }
  }
  return null;
}

function content(pres, s) {
  const slide = pres.addSlide({ masterName: "STUDY_CONTENT" });
  eyebrow(slide, s.eyebrow, s.accent);
  title(slide, s.title);
  if (s.lead) {
    const h = leadH(s);
    card(slide, mx, TOP, cw, h);
    text(slide, s.lead, { x: mx + 0.4, y: TOP, w: cw - 0.8, h, size: SIZE.lead, bold: true, valign: "middle" });
  }
  return slide;
}

// 상단 질문 띠(lead) 높이와, 그만큼 내려간 본문 시작 위치
const leadH = (s) => textH(s.lead, SIZE.lead, cw - 0.8) + 0.4;
const bodyTop = (s) => TOP + (s.lead ? leadH(s) + GAP : 0);

// 오른쪽(또는 왼쪽) 설명 열: bullets → keyline → note
// 공간이 남으면 글씨를 한 단계 키운다(18/20/16), 모자라면 기본(16/18/14)
// fillH: 왼쪽 코드(그림) 높이. 주면 항목 사이를 벌려(최대 SPREAD_MAX) 오른쪽 아래가 비지 않게 한다
const SPREAD_MAX = 0.7;
function rightColumn(slide, s, x, y, w, bottom, fillH) {
  const steps = [{ b: 18, k: 20, n: 16 }, { b: SIZE.body, k: SIZE.keyline, n: SIZE.bodySmall }];
  const need = (z) => (s.bullets?.length ? bulletsH(s.bullets, z.b, w) + 0.2 : 0) + (s.keyline ? textH(s.keyline, z.k, w) + 0.2 : 0) + (s.note ? textH(s.note, z.n, w) : 0);
  const z = steps.find((st) => y + need(st) <= bottom) || steps[steps.length - 1];
  const items = s.bullets || [];
  const gaps = items.length - 1 + (s.keyline ? 1 : 0) + (s.note ? 1 : 0);
  const target = fillH ? Math.min(fillH, bottom - y) : 0;
  const extra = gaps > 0 && target ? Math.max(0, Math.min(SPREAD_MAX, (target - need(z)) / gaps)) : 0;
  let cy = y;
  items.forEach((it, i) => { cy += bullets(slide, [it], { x, y: cy, w, size: z.b }) + (i < items.length - 1 ? extra : 0); });
  if (items.length) cy += 0.2;
  if (s.keyline) { if (items.length) cy += extra; const h = textH(s.keyline, z.k, w); text(slide, s.keyline, { x, y: cy, w, h, size: z.k, bold: true }); cy += h + 0.2; }
  if (s.note) { if (items.length || s.keyline) cy += extra; const h = textH(s.note, z.n, w); text(slide, s.note, { x, y: cy, w, h, size: z.n, color: COLOR.sub }); cy += h; }
  if (cy > bottom + 0.1) warn(`설명 글이 영역을 넘칩니다(${(cy - bottom).toFixed(2)}in) — 문장을 줄이세요`);
}

// 그림 옆 설명: 항목을 그림 영역 높이에 고르게 배치해 그림과 평행을 맞춘다
function spreadBullets(slide, items, x, y, w, h) {
  const size = 18;
  const slot = h / items.length;
  items.forEach((it, i) => {
    const th = textH(it, size, w - 0.25);
    const r = rich(it, { color: COLOR.ink, fontSize: size, fontFace: FONT.body, lang: LANG }, EMPH);
    r[0].options.bullet = { indent: 16 };
    slide.addText(r, { x, y: y + i * slot + (slot - th) / 2, w, h: th, isTextBox: true, margin: 0, valign: "middle", fontFace: FONT.body, fontSize: size, lang: LANG, lineSpacingMultiple: 1.1 });
    if (th > slot) warn("그림 옆 설명이 칸보다 깁니다 — 항목 문장을 줄이세요");
  });
}

// ───────────────────────── 시간·노트·대본 ─────────────────────────
const toSec = (t) => { if (!t) return 0; const [m, s] = String(t).split(":").map(Number); return (m || 0) * 60 + (s || 0); };
const fmt = (sec) => `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, "0")}`;
const label = (s) => s.label || plain(s.title || s.text || { cover: "목차", summary: "정리" }[s.type] || s.type);

function notesFor(s, i, total, cum) {
  const head = `[${i + 1}/${total} · ${s.time || "-"} · 누적 ${fmt(cum)}]${s.tag ? ` 【${s.tag}】` : ""}`;
  const lines = (s.script || []).map(plain);
  const cues = (s.cues || []).map((c) => `· ${plain(c)}`);
  return [head, "", ...lines, ...(cues.length ? ["", "── 연출 ──", ...cues] : [])].join("\n");
}

function buildScript(slides, times) {
  const total = times.reduce((a, b) => a + b, 0);
  const md = [];
  md.push(`# 발표 대본 — ${meta.title || ""}`.trim(), "");
  md.push(`> 대상 파일: \`${pptxName}\` (**${slides.length}장**)`);
  md.push(`> 목표 러닝타임: **${meta.runtime || fmt(total)}** (합계 ${fmt(total)}, 상한 15분, 질문 여유 포함)`);
  if (meta.scope) md.push(`> 모든 내용은 강의자료 PDF(${meta.scope}) 범위 안에서만 구성했습니다.`);
  md.push("", "## 시간 배분", "", "| # | 슬라이드 | 시간 | 누적 |", "|---|---|---|---|");
  let cum = 0;
  slides.forEach((s, i) => {
    cum += times[i];
    const name = label(s).replace(/\|/g, "\\|");
    md.push(`| ${i + 1} | ${s.tag && /핵심|실무|데모/.test(s.tag) ? `**${name}**` : name} | ${s.time || "-"} | ${fmt(cum)} |`);
  });
  if (meta.compress) md.push("", `> 밀리면 ${meta.compress}`);
  slides.forEach((s, i) => {
    md.push("", "---", "", `## ${i + 1}. ${label(s)} · ${s.time || "-"}${s.tag ? ` 【${s.tag}】` : ""}`, "");
    (s.script || []).forEach((l) => md.push(`> ${l}`));
    const cues = s.cues || [];
    if (cues.length === 1) md.push("", `**강조** — ${cues[0]}`);
    else if (cues.length > 1) md.push("", "**강조**", ...cues.map((c) => `- ${c}`));
  });
  if ((spec.rehearsal || []).length) md.push("", "---", "", "## 리허설 체크리스트", "", ...spec.rehearsal.map((r) => `- [ ] ${r}`));
  if ((spec.qna || []).length) {
    md.push("", "## 예상 질문과 답");
    spec.qna.forEach((q) => md.push("", `**Q. ${q.q}**`, `> ${q.a}`));
  }
  return md.join("\n") + "\n";
}

// ───────────────────────── 후처리 ─────────────────────────
// 1) pptxgenjs가 색이 섞인 문단에 <a:pPr>를 여러 번 쓰는 문제 → 문단 첫 pPr만 남긴다
// 2) 코드 폰트(Consolas)의 한글은 Pretendard로 표시
// 3) 테마 폰트의 동아시아(ea) 글꼴도 Pretendard로 지정 → PowerPoint에서 새로 입력하는 한글도 Pretendard
async function postProcess(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/(slides|notesSlides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) {
      let xml = await zip.file(name).async("string");
      xml = xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
        let idx = 0;
        const fixed = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (pp, off) => (idx++ === 0 && off === 0 ? pp : ""));
        return `<a:p>${fixed}</a:p>`;
      });
      xml = xml.replace(/<a:ea typeface="Consolas"/g, `<a:ea typeface="${FONT.body}"`);
      zip.file(name, xml);
    }
    if (/^ppt\/theme\/theme\d+\.xml$/.test(name)) {
      const xml = (await zip.file(name).async("string")).replace(/<a:ea typeface=""\s*\/>/g, `<a:ea typeface="${FONT.body}"/>`);
      zip.file(name, xml);
    }
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ───────────────────────── 실행 ─────────────────────────
async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: FONT.head, bodyFontFace: FONT.body };
  pres.title = meta.title || "";
  pres.author = meta.presenter || "";
  pres.subject = meta.course || "";

  const titlePh = (o) => ({ placeholder: { options: { name: "title", type: "title", margin: 0, bold: true, align: "left", color: COLOR.ink, fontFace: FONT.head, ...o }, text: "" } });
  const num = { x: RIGHT - 0.8, y: 6.95, w: 0.8, h: 0.3, fontFace: FONT.body, fontSize: SIZE.caption, color: COLOR.sub, align: "right" };
  pres.defineSlideMaster({ title: "STUDY_COVER", background: { color: COLOR.bg }, objects: [titlePh({ x: mx, y: 1.0, w: cw, h: 1.0, fontSize: SIZE.coverTitle, valign: "middle" })] });
  pres.defineSlideMaster({ title: "STUDY_CONTENT", background: { color: COLOR.bg }, objects: [titlePh({ x: mx, y: GRID.titleY, w: cw, h: GRID.titleH, fontSize: SIZE.title, valign: "middle" })], slideNumber: num });
  pres.defineSlideMaster({ title: "STUDY_STATEMENT", background: { color: COLOR.bg }, objects: [titlePh({ x: mx, y: 1.9, w: cw, h: 2.4, fontSize: SIZE.statement, valign: "bottom" })], slideNumber: num });
  pres.defineSlideMaster({ title: "STUDY_QUIZ", background: { color: COLOR.bg }, objects: [titlePh({ x: mx, y: 0.95, w: cw, h: 1.35, fontSize: SIZE.quiz, valign: "top" })], slideNumber: num });
  pres.defineSlideMaster({ title: "STUDY_PRINCIPLE", background: { color: COLOR.bg }, objects: [titlePh({ x: mx, y: 1.2, w: 5.9, h: 2.4, fontSize: SIZE.principle, valign: "top" })], slideNumber: num });

  const slides = expandSlides(spec.slides || []);
  const times = slides.map((s) => toSec(s.time));
  let cum = 0;
  slides.forEach((s, i) => {
    currentSlide = i + 1;
    const render = RENDER[s.type];
    if (!render) { warn(`알 수 없는 type "${s.type}" — 건너뜀`); return; }
    const slide = render(pres, s);
    cum += times[i];
    slide.addNotes(notesFor(s, i, slides.length, cum));
    if (!s.time) warn("time(발표 시간)이 없습니다");
    if (!(s.script || []).length) warn("script(대사)가 없습니다");
  });

  const total = times.reduce((a, b) => a + b, 0);
  currentSlide = "전체";
  if (total > 900) warn(`총 발표 시간 ${fmt(total)} — 상한 15분을 넘습니다. 줄일 슬라이드를 정하세요(사용자 핵심 주제는 유지)`);
  else if (total && total < 600) warn(`총 발표 시간 ${fmt(total)} — 10분보다 짧습니다`);
  const count = (spec.slides || []).length; // 퀴즈 한 쌍(문제·정답)은 1장으로 센다
  if (count > 16) warn(`슬라이드 ${count}장(퀴즈는 문제당 1장으로 계산) — 15분 발표에는 많습니다(권장 12~15장)`);
  const quizN = (spec.slides || []).filter((s) => s.type === "quiz-item").length;
  if (quizN && quizN !== 2) warn(`퀴즈 ${quizN}문제 — 규격은 2문제(핵심 주제 대상)입니다`);

  fs.mkdirSync(outDir, { recursive: true });
  await pres.writeFile({ fileName: pptxPath });
  await postProcess(pptxPath);
  const sharePath = await makeShare(pptxPath); // 노트(대본)를 뺀 공유본 — 항상 같이 만든다
  fs.writeFileSync(scriptPath, buildScript(slides, times), "utf8");

  console.log(`PPTX  : ${pptxPath}`);
  console.log(`공유본: ${sharePath}`);
  console.log(`대본  : ${scriptPath}`);
  console.log(`스타일: ${STYLE_VERSION}`);
  console.log(`구성  : ${slides.length}장 · 총 ${fmt(total)}`);
  if (warnings.length) { console.log(`\n경고 ${warnings.length}건`); warnings.forEach((w) => console.log(`  - ${w}`)); }
  else console.log("경고 없음");
}

main().catch((e) => { console.error(e); process.exit(1); });
