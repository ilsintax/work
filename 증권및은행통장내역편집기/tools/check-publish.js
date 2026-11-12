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
// 샘플 위치: samples/ 폴더(없으면 프로젝트 루트). 샘플이 0개면 민감 문자열을 모을 수 없어 검사가 무의미하므로 실패로 끝낸다.
const SAMPLES = fs.existsSync(path.join(root, 'samples')) ? path.join(root, 'samples') : root;
const samples = fs.readdirSync(SAMPLES).filter(f => /\.(xlsx|xls|csv)$/i.test(f));
if (!samples.length) { console.log('샘플 파일을 찾지 못했습니다(samples/ 또는 루트) — 고객 정보 검사를 할 수 없습니다.'); process.exit(1); }
samples.forEach(f => {
  const acc = (f.match(/\(([^)]+)\)/) || [])[1];
  if (acc) sensitive.add(acc);
  // 파일명의 고객명·계좌번호 조각(기관 이름·일반 단어 제외)도 민감 문자열로
  const inst = new Set();
  const D0 = globalThis.BANKBOOK_DICTIONARY;
  [...D0.securities.domestic, ...D0.securities.foreign].forEach(b => [b.name, ...(b.aliases || [])].forEach(x => inst.add(x)));
  ((D0.bank && D0.bank.names) || []).forEach(x => inst.add(x));
  D0.abbreviations.forEach(a => a.keys.forEach(k => inst.add(k)));
  f.replace(/\.[^.]+$/, '').split(/[_()\s]+/).forEach(t => {
    if (t.length >= 2 && !inst.has(t) && !/^(거래내역|잔고증명|\d{1,2}월|\d{2}년상반기|\d{2}년하반기|KB\d|키움해외|아이엠증권|한국투자증권)$/.test(t) && !/^\d{1,2}$/.test(t)) sensitive.add(t);
  });
  E.analyzeFile(XLSX, new Uint8Array(fs.readFileSync(path.join(SAMPLES, f))), f, globalThis.BANKBOOK_DICTIONARY).forEach(r => {
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

// 사전 안의 거래처 이름: 샘플에 없는 거래처도 잡도록, 회사 표시(주식회사·㈜·(주) 등)와 거래처가 붙은 거래종류를 따로 검사
{
  const dictText = fs.readFileSync(path.join(root, 'dictionary', 'institutions.js'), 'utf8');
  const company = dictText.match(/"[^"\n]*(주식회사|㈜|\(주\)|유한회사|합자회사)[^"\n]*"/g) || [];
  company.forEach(s => hits.push(`dictionary/institutions.js: 회사 이름으로 보이는 글자 ${s.slice(0, 60)}`));
  const DICT = require(path.join(root, 'engine', 'dictionary.js'));
  try {
    const dd = DICT.parse(dictText);
    (dd.loadInfo ? dd.loadInfo.warnings : []).filter(w => /거래처 이름/.test(w)).forEach(w => hits.push(`dictionary/institutions.js: ${w}`));
  } catch (e) { hits.push(`dictionary/institutions.js: 사전 검사 오류 ${e.message}`); }
}

console.log(`올라갈 파일 ${publish.length}개, 검사한 민감 문자열 ${sensitive.size}개 (샘플 ${samples.length}개 기준) + 사전의 회사·거래처 이름 검사`);
if (hits.length) { console.log('고객 정보가 들어 있는 파일:'); hits.forEach(h => console.log('  ' + h)); process.exit(1); }
console.log('고객 정보 없음');
if (process.argv.includes('--list')) publish.forEach(f => console.log('  ' + f));
