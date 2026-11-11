// 사전 검사 명령. GitHub 에 올리기 전에, 화면에서 저장한 institutions.js 를 이 명령으로 먼저 검사한다.
// 실행: node tools/check-dictionary.js [사전파일]      (기본: dictionary/institutions.js)
//       node tools/check-dictionary.js --migrate        (dictionary/institutions.js 를 현재 구조로 바꿔 저장 — 먼저 백업할 것)
// 종료 코드: 0 = 사용 가능(경고는 있을 수 있음), 1 = 오류가 있어 쓸 수 없음
'use strict';
const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
require(path.join(root, 'engine', 'engine.js'));
const DICT = require(path.join(root, 'engine', 'dictionary.js'));

const args = process.argv.slice(2);
const migrateMode = args.includes('--migrate');
const file = path.resolve(args.find(a => !a.startsWith('--')) || path.join(root, 'dictionary', 'institutions.js'));
const text = fs.readFileSync(file, 'utf8');

let dict;
try {
  dict = DICT.parse(text); // 구조 변환 + 검사(오류면 예외)
} catch (e) {
  console.log(`오류: ${e.message}`);
  (e.errors || []).forEach(x => console.log(`  - ${x}`));
  process.exit(1);
}

const info = dict.loadInfo || { notes: [], warnings: [] };
const count = m => dict.securities[m].length;
const profiles = m => dict.securities[m].reduce((n, b) => n + (b.profiles || []).length, 0);
const shared = m => dict.securities[m].reduce((n, b) => n + ['depositTypes', 'withdrawalTypes', 'removalTypes'].reduce((k, x) => k + ((b.shared || {})[x] || []).length, 0), 0);
console.log(`사전: ${path.relative(root, file) || file}`);
console.log(`  구조 버전 ${dict.schemaVersion} · 내용 버전 ${dict.version || 1}${dict.updatedAt ? ` (${dict.updatedAt})` : ''}`);
console.log(`  국내 증권사 ${count('domestic')}곳 · 형식 ${profiles('domestic')}개 · 공통 분류 ${shared('domestic')}개`);
console.log(`  해외 증권사 ${count('foreign')}곳 · 형식 ${profiles('foreign')}개 · 공통 분류 ${shared('foreign')}개`);
console.log(`  은행 형식 ${dict.bank.profiles.length}개 · 파일명 약칭 ${dict.abbreviations.length}개`);
if (info.notes.length) { console.log('구조 변환(불러올 때 자동 적용):'); info.notes.forEach(n => console.log(`  - ${n}`)); }
if (info.warnings.length) { console.log(`경고 ${info.warnings.length}건(사용 가능):`); info.warnings.forEach(w => console.log(`  - ${w}`)); }
console.log('오류 없음 — 사용할 수 있습니다.');

if (migrateMode) {
  if (!info.notes.length) { console.log('이미 현재 구조입니다. 저장하지 않았습니다.'); process.exit(0); }
  const out = DICT.serialize(dict)
    .replace(/^\/\/ 통합 통장내역 편집기 기관사전\. 화면의 "사전 파일 저장"으로 만든 파일입니다\.\n/, '// 통합 통장내역 편집기 기관사전 (tools/check-dictionary.js --migrate 로 구조 변환한 파일).\n');
  fs.writeFileSync(file, out, 'utf8');
  // 다시 읽어 같은지 확인
  const back = DICT.parse(fs.readFileSync(file, 'utf8'));
  const same = JSON.stringify({ ...back }) === JSON.stringify((() => { const o = JSON.parse(JSON.stringify(dict)); delete o.accounts; return o; })());
  console.log(same ? `구조 변환 저장 완료: ${file}` : '저장 후 다시 읽은 내용이 다릅니다 — 백업에서 되돌리세요.');
  process.exit(same ? 0 : 1);
}
