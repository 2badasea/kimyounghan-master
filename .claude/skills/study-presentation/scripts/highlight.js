// 코드 문자열 → 줄 단위 토큰 배열 [[{text, color}], ...]
// 지원 lang: java(기본) · gradle · html · sql · properties · yaml · output(text)
const { CODE } = require("./style");

const JAVA_KEYWORDS = new Set((
  "abstract assert boolean break byte case catch char class const continue default do double else enum " +
  "extends final finally float for goto if implements import instanceof int interface long native new " +
  "package private protected public return short static strictfp super switch synchronized this throw " +
  "throws transient try void volatile while var record yield sealed permits true false null def"
).split(" "));

const SQL_KEYWORDS = new Set((
  "select from where insert into values update set delete create table drop alter add primary key foreign " +
  "references not null default auto_increment unique bigint int integer varchar char text datetime timestamp " +
  "date boolean order by desc asc and or engine charset collate join on left right inner outer group having " +
  "limit offset as is in like between distinct count sum avg min max now current_timestamp if exists index"
).split(" "));

const RE = {
  dqString: /^"(?:[^"\\]|\\.)*"?/,
  sqString: /^'(?:[^'\\]|\\.)*'?/,
  annotation: /^@[A-Za-z_]\w*/,
  number: /^(0[xX][0-9a-fA-F_]+|\d[\d_]*(\.\d+)?([eE][+-]?\d+)?)[lLfFdD]?\b/,
  ident: /^[A-Za-z_$][\w$]*/,
  callParen: /^\s*\(/,
};

function push(line, text, color) {
  if (!text) return;
  const last = line[line.length - 1];
  if (last && last.color === color) last.text += text;
  else line.push({ text, color });
}

// 자바/그루비 계열 (블록 주석 상태를 줄 사이에 유지)
function tokenizeJava(src) {
  const lines = [];
  let inBlock = false;
  for (const raw of src.split("\n")) {
    const line = [];
    let i = 0;
    let prevWord = "";
    while (i < raw.length) {
      const rest = raw.slice(i);
      let m;
      if (inBlock) {
        const end = rest.indexOf("*/");
        if (end < 0) { push(line, rest, CODE.comment); break; }
        push(line, rest.slice(0, end + 2), CODE.comment); i += end + 2; inBlock = false; continue;
      }
      if (rest.startsWith("//")) { push(line, rest, CODE.comment); break; }
      if (rest.startsWith("/*")) { inBlock = true; continue; }
      if ((m = rest.match(RE.dqString)) || (m = rest.match(RE.sqString))) {
        push(line, m[0], CODE.string); i += m[0].length; prevWord = ""; continue;
      }
      if ((m = rest.match(RE.annotation))) { push(line, m[0], CODE.annotation); i += m[0].length; continue; }
      if ((m = rest.match(RE.number))) { push(line, m[0], CODE.number); i += m[0].length; continue; }
      if ((m = rest.match(RE.ident))) {
        const w = m[0];
        let color = CODE.plain;
        if (JAVA_KEYWORDS.has(w)) color = CODE.keyword;
        else if (RE.callParen.test(raw.slice(i + w.length)) && prevWord !== "new") color = CODE.method;
        push(line, w, color); i += w.length; prevWord = w; continue;
      }
      push(line, raw[i], CODE.plain);
      if (!/\s/.test(raw[i])) prevWord = "";
      i += 1;
    }
    lines.push(line);
  }
  return lines;
}

function tokenizeHtml(src) {
  const lines = [];
  let inComment = false;
  let inTag = false;
  for (const raw of src.split("\n")) {
    const line = [];
    let i = 0;
    while (i < raw.length) {
      const rest = raw.slice(i);
      let m;
      if (inComment) {
        const end = rest.indexOf("-->");
        if (end < 0) { push(line, rest, CODE.comment); break; }
        push(line, rest.slice(0, end + 3), CODE.comment); i += end + 3; inComment = false; continue;
      }
      if (!inTag && rest.startsWith("<!--")) { inComment = true; continue; }
      if (!inTag && (m = rest.match(/^<\/?[A-Za-z!][\w:-]*/))) { push(line, m[0], CODE.tag); i += m[0].length; inTag = true; continue; }
      if (inTag) {
        if ((m = rest.match(/^\/?>/))) { push(line, m[0], CODE.tag); i += m[0].length; inTag = false; continue; }
        if ((m = rest.match(/^"[^"]*"?/)) || (m = rest.match(/^'[^']*'?/))) { push(line, m[0], CODE.attrValue); i += m[0].length; continue; }
        if ((m = rest.match(/^[A-Za-z_:][\w:.-]*/))) { push(line, m[0], CODE.attr); i += m[0].length; continue; }
        push(line, raw[i], CODE.plain); i += 1; continue;
      }
      const next = rest.indexOf("<");
      const chunk = next === 0 ? raw[i] : next < 0 ? rest : rest.slice(0, next);
      push(line, chunk, CODE.plain); i += chunk.length;
    }
    lines.push(line);
  }
  return lines;
}

function tokenizeSql(src) {
  return src.split("\n").map((raw) => {
    const line = [];
    let i = 0;
    while (i < raw.length) {
      const rest = raw.slice(i);
      let m;
      if (rest.startsWith("--")) { push(line, rest, CODE.comment); break; }
      if ((m = rest.match(/^'(?:[^']|'')*'?/))) { push(line, m[0], CODE.string); i += m[0].length; continue; }
      if ((m = rest.match(/^\d+(\.\d+)?\b/))) { push(line, m[0], CODE.number); i += m[0].length; continue; }
      if ((m = rest.match(/^[A-Za-z_]\w*/))) {
        push(line, m[0], SQL_KEYWORDS.has(m[0].toLowerCase()) ? CODE.keyword : CODE.plain); i += m[0].length; continue;
      }
      push(line, raw[i], CODE.plain); i += 1;
    }
    return line;
  });
}

function tokenizeProps(src) {
  return src.split("\n").map((raw) => {
    const line = [];
    if (/^\s*#/.test(raw)) { push(line, raw, CODE.comment); return line; }
    const m = raw.match(/^(\s*[^=:\s][^=:]*?)(\s*[=:]\s*)(.*)$/);
    if (!m) { push(line, raw, CODE.plain); return line; }
    push(line, m[1], CODE.keyword); push(line, m[2], CODE.plain); push(line, m[3], CODE.string);
    return line;
  });
}

function tokenizeOutput(src) {
  return src.split("\n").map((raw) => {
    const line = [];
    for (const part of raw.split(/(-?\b\d[\d,.]*\b)/)) {
      if (!part) continue;
      push(line, part, /^-?\d/.test(part) ? CODE.number : CODE.plain);
    }
    return line;
  });
}

function highlight(src, lang = "java") {
  const code = String(src).replace(/\t/g, "    ").replace(/\s+$/, "");
  switch ((lang || "java").toLowerCase()) {
    case "html": case "thymeleaf": case "xml": return tokenizeHtml(code);
    case "sql": return tokenizeSql(code);
    case "properties": case "yaml": case "yml": return tokenizeProps(code);
    case "output": case "text": case "console": return tokenizeOutput(code);
    default: return tokenizeJava(code); // java, gradle, groovy, kotlin(근사)
  }
}

module.exports = { highlight };
