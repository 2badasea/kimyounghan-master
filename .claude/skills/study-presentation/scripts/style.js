// 스타일 버전 선택기 — 실제 디자인 토큰은 styles/<버전>.js 에 있다.
// 버전은 build.js 가 덱 스펙의 meta.style 을 읽어 STUDY_STYLE 로 넘긴다(없으면 v1).
const fs = require("fs");
const path = require("path");

const VERSION = process.env.STUDY_STYLE || "v1";
const file = path.join(__dirname, "styles", `${VERSION}.js`);
if (!fs.existsSync(file)) {
  const have = fs.readdirSync(path.join(__dirname, "styles")).map((f) => f.replace(/\.js$/, ""));
  throw new Error(`스타일 버전 "${VERSION}" 이 없습니다. 사용 가능: ${have.join(", ")}`);
}

module.exports = { VERSION, ...require(file) };
