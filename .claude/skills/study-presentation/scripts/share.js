#!/usr/bin/env node
// 공유용 pptx 만들기 — 슬라이드 노트(발표자 대본)를 통째로 뺀 사본을 만든다.
//   node share.js <발표.pptx> [출력.pptx]      (출력 생략 시 <이름>_공유.pptx)
// build.js 가 빌드 끝에 자동으로 호출한다. 생성기로 만들지 않은 pptx(예: 예전 발표)에도 쓸 수 있다.

const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const shareName = (file) => path.join(path.dirname(file), `${path.basename(file, ".pptx")}_공유.pptx`);

async function makeShare(src, dest = shareName(src)) {
  const zip = await JSZip.loadAsync(fs.readFileSync(src));

  // 1) 노트 슬라이드 파일과 그 관계 파일 삭제
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/notesSlides\//.test(name)) zip.remove(name);
  }
  // 2) 각 슬라이드 관계에서 노트 슬라이드 참조 제거
  for (const name of Object.keys(zip.files)) {
    if (!/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(name)) continue;
    const xml = await zip.file(name).async("string");
    zip.file(name, xml.replace(/<Relationship\b[^>]*relationships\/notesSlide"[^>]*\/>/g, ""));
  }
  // 3) 콘텐츠 형식 목록에서 노트 슬라이드 항목 제거 (노트 마스터는 남겨도 무방)
  const ct = "[Content_Types].xml";
  const ctXml = await zip.file(ct).async("string");
  zip.file(ct, ctXml.replace(/<Override\b[^>]*PartName="\/ppt\/notesSlides\/[^"]*"[^>]*\/>/g, ""));

  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(dest, buf);
  return dest;
}

module.exports = { makeShare, shareName };

if (require.main === module) {
  const src = process.argv[2];
  if (!src) { console.error("usage: node share.js <발표.pptx> [출력.pptx]"); process.exit(1); }
  makeShare(path.resolve(src), process.argv[3] && path.resolve(process.argv[3]))
    .then((out) => console.log(`공유본: ${out}`))
    .catch((e) => { console.error(e); process.exit(1); });
}
