// GitHub 에 올리기 전 고객 정보 검사.
// 로컬 샘플 엑셀에서 민감 문자열(계좌번호, 은행 거래상대방, 제목 줄의 계좌명·고객명)을 모아,
// .gitignore 로 제외되지 않는(=올라갈) 파일에 들어 있는지 찾는다.
// 실행: node tools/check-publish.js   (문제 있으면 종료 코드 1)
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
XLSX.set_cptable(require('xlsx/dist/cpexcel.js'));
const root = path.resolve(__dirname, '..');
require(path.join(root, 'dictionary', 'institutions.js'));
const E = require(path.join(root, 'engine', 'engine.js'));

// .gitignore 의 단순 규칙(폴더/, *.확장자, 경로)만 해석
const rules = fs.readFileSync(path.join(root, '.gitignore'), 'utf8').split(/\r?\n/).map(s => s.trim()).filter(s => s && !s.startsWith('#'));
const ignored = rel => rules.some(r => {
  if (r.endsWith('/')) return rel === r.slice(0, -1) || rel.startsWith(r);
  if (r.startsWith('*.')) return rel.toLowerCase().endsWith(r.slice(1).toLowerCase());
  return rel === r;
});
const walk = dir => fs.readdirSync(path.join(root, dir), { withFileTypes: true }).flatMap(d => {
  const rel = dir ? `${dir}/${d.name}` : d.name;
  if (d.name === '.git' || ignored(d.isDirectory() ? rel + '/' : rel)) return [];
  return d.isDirectory() ? walk(rel) : [rel];
});
const publish = walk('');

// 민감 문자열 수집
const sensitive = new Set();
const samples = fs.readdirSync(root).filter(f => /\.(xlsx|xls|csv)$/i.test(f));
samples.forEach(f => {
  const acc = (f.match(/\(([^)]+)\)/) || [])[1];
  if (acc) sensitive.add(acc);
  E.analyzeFile(XLSX, new Uint8Array(fs.readFileSync(path.join(root, f))), f, globalThis.BANKBOOK_DICTIONARY).forEach(r => {
    const L = r.layout && r.layout.ok ? r.layout : null;
    (r.rows || []).slice(0, L ? L.dataStart : 10).forEach(row => (row || []).forEach((v, i) => {
      if (/^(계좌명|고객명|계좌번호|계좌정보)$/.test(String(v || '').trim()) && row[i + 1]) sensitive.add(String(row[i + 1]).trim());
    }));
    if (L && r.kind === 'bank') {
      const cp = L.header.indexOf('보낸분/받는분');
      if (cp >= 0) L.records.forEach(rec => { const v = String(rec[cp] || '').trim(); if (v.length >= 4) sensitive.add(v); });
    }
  });
});

// 사전에 원래 있는 기관명(증권사·은행명, 별칭, 약칭)은 공개 정보이므로 제외
const D = globalThis.BANKBOOK_DICTIONARY;
const publicNames = new Set([
  ...D.securities.domestic.flatMap(b => [b.name, ...(b.aliases || [])]),
  ...D.bank.profiles.flatMap(p => [p.name, ...(p.aliases || [])]),
  ...D.abbreviations.flatMap(a => [...a.keys, a.bank, a.securities])
].filter(Boolean).map(s => String(s).replace(/\s+/g, '')));
[...sensitive].forEach(s => { if (publicNames.has(s.replace(/\s+/g, ''))) sensitive.delete(s); });

const hits = [];
publish.filter(f => !/\.(png|jpg|gif|ico)$/i.test(f) && f !== 'lib/xlsx.full.min.js').forEach(f => {
  const text = fs.readFileSync(path.join(root, f), 'utf8');
  sensitive.forEach(s => { if (text.includes(s)) hits.push(`${f}: "${s}"`); });
});

console.log(`올라갈 파일 ${publish.length}개, 검사한 민감 문자열 ${sensitive.size}개 (샘플 ${samples.length}개 기준)`);
if (hits.length) { console.log('고객 정보가 들어 있는 파일:'); hits.forEach(h => console.log('  ' + h)); process.exit(1); }
console.log('고객 정보 없음');
if (process.argv.includes('--list')) publish.forEach(f => console.log('  ' + f));
