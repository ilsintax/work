// 예외 계좌 분석 도구 — 사전에 기록하기 전에 파일을 진단한다. 파일·사전은 읽기만 한다.
// 실행: node tools/analyze-account.js <파일경로> [파일경로...]
// 보고: ① 파일·구조 ② 기관 판별 근거(파일명, 증권사 형식별 열 일치율) ③ 통화 구성
//       ④ 현재 사전으로 변환한 결과·잔액 검증 ⑤ 통화별(국내/해외 사전) 변환 결과 ⑥ 사전에 없는 거래종류·열 이름 차이 ⑦ 기록 후보
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
XLSX.set_cptable(require('xlsx/dist/cpexcel.js'));
const root = path.resolve(__dirname, '..');
require(path.join(root, 'dictionary', 'institutions.js'));
const E = require(path.join(root, 'engine', 'engine.js'));
require(path.join(root, 'engine', 'securities-legacy.js'));
const C = require(path.join(root, 'engine', 'convert.js'));
const D = globalThis.BANKBOOK_DICTIONARY;
// 레거시 엔진의 옛 알림은 분석 결과에 섞이지 않게 숨김
const warn = console.warn; console.warn = (...a) => { if (!/legacy notice/.test(String(a[0]))) warn(...a); };

const REQ = ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'];
const CURRENCY_HEADERS = ['통화', '통화코드', '통화구분'];
const DOMESTIC_CURRENCY = ['', 'KRW', '원화', '원', '0'];
const n = v => (Number.isFinite(v) ? Math.round(v * 100) / 100 : v);
const out = [];
const p = (...a) => out.push(a.join(' '));

function profileMatches(header, presets) {
  const rows = [];
  presets.forEach(b => (b.profiles || []).forEach((pr, i) => {
    const keys = Object.entries(pr.columns || {});
    const hit = keys.filter(([, c]) => E.findHeader(header, c) >= 0);
    const miss = keys.filter(([, c]) => E.findHeader(header, c) < 0).map(([k, c]) => `${k}=${c.join('|')}`);
    rows.push({ broker: b.name, index: i, hit: hit.length, total: keys.length, req: hit.filter(([k]) => REQ.includes(k)).length, miss });
  }));
  return rows.sort((a, b) => b.hit / b.total - a.hit / a.total || b.hit - a.hit);
}

// 파일명 어디에 있든 약칭 찾기 (현재 엔진은 "번호_약칭(" 형식만 읽음)
function abbreviationsInName(fileName) {
  const tokens = fileName.replace(/\.[^.]+$/, '').split(/[_\s()\-]+/).filter(Boolean);
  const hits = [];
  tokens.forEach(t => (D.abbreviations || []).forEach(e => e.keys.forEach(k => {
    if (k.length >= 2 && t.toUpperCase().includes(k.toUpperCase())) hits.push({ token: t, key: k, bank: e.bank, securities: e.securities });
  })));
  return hits;
}

function convertSubset(account, keepIdx, dict, label) {
  const a = C.createAccount({ ...account, layout: { ...account.layout, records: keepIdx.map(i => account.layout.records[i]), sourceRows: keepIdx.map(i => account.layout.sourceRows[i]) } }, 1);
  C.convertAccount(a, dict);
  const o = a.output || {};
  if (o.error) { p(`  [${label}] 변환 불가: ${o.error}`); return a; }
  const v = o.verification;
  p(`  [${label}] ${keepIdx.length}건 → 결과 ${o.sheet.length - 1}행 · 판별 ${a.institution || '-'} · 상태 ${a.status} · 기초 ${n(v.baseBalance)} → 계산 ${n(v.calculated)} / 원본 ${n(v.original)} (차이 ${n(v.diff)}, 불일치 ${v.mismatchCount}곳)`);
  p(`     분류 입금 ${JSON.stringify(a.config.classification.depositTypes)} 출금 ${JSON.stringify(a.config.classification.withdrawalTypes)} 제거 ${JSON.stringify(a.config.classification.valuesToRemove)} 수수료 ${a.config.classification.inferredFeeMode}`);
  (o.issues || []).forEach(m => p(`     · ${m}`));
  (v.mismatchRows || []).slice(0, 6).forEach(r => p(`     불일치 행: ${JSON.stringify(o.sheet[r])}`));
  return a;
}

function analyze(file) {
  const fileName = path.basename(file);
  const results = E.analyzeFile(XLSX, new Uint8Array(fs.readFileSync(file)), fileName, D);
  p(`\n==================== ${fileName}`);
  const nameHits = abbreviationsInName(fileName);
  p(`파일명 해석(현재 엔진): ${JSON.stringify(E.parseFileName(fileName))}`);
  p(`파일명 안의 약칭 후보: ${nameHits.length ? nameHits.map(h => `${h.token}→${[h.securities, h.bank].filter(Boolean).join('/')}`).join(', ') : '없음'}${/해외|외화|global/i.test(fileName) ? ' · "해외/외화" 표시 있음' : ''}`);
  results.forEach(r => {
    p(`\n--- 시트: ${r.sheetName} (형식 ${r.format})`);
    if (!r.layout || !r.layout.ok) { p(`구조 감지 실패: ${(r.messages || []).join(' / ')}`); return; }
    const L = r.layout;
    p(`구조: 제목 ${L.titleRows.length}줄 · 헤더 ${L.headerRows.length}줄(줄바꿈 ${L.continuationRows.length}) · 거래당 ${L.stride}줄 · 거래 ${L.records.length}건 · 묶이지 않은 행 ${L.skippedRows.length}`);
    p(`헤더: ${L.header.filter(Boolean).join(' | ')}`);
    p(`현재 판별: ${r.kind} / ${r.institution || '-'} / ${r.status} — ${(r.messages || []).join(' / ') || '메시지 없음'}`);

    const dom = profileMatches(L.header, D.securities.domestic).slice(0, 4);
    const forn = profileMatches(L.header, D.securities.foreign).filter(x => x.total).slice(0, 2);
    p('국내 사전 형식 일치 상위:'); dom.forEach(m => p(`  ${m.broker} 형식${m.index + 1}: ${m.hit}/${m.total} (필수 ${m.req}/4)${m.miss.length ? ' 없음: ' + m.miss.join(', ') : ' — 완전 일치'}`));
    if (forn.length) { p('해외 사전 형식 일치:'); forn.forEach(m => p(`  ${m.broker} 형식${m.index + 1}: ${m.hit}/${m.total}${m.miss.length ? ' 없음: ' + m.miss.join(', ') : ' — 완전 일치'}`)); }
    const full = dom.filter(m => m.hit === m.total);
    p(`보수적 판정: ${full.length === 1 ? `완전 일치 1곳(${full[0].broker})` : full.length > 1 ? `완전 일치 ${full.length}곳 → 판별 보류` : '완전 일치 없음 → 자동 확정 금지'}${nameHits.length ? ` / 파일명 후보 ${[...new Set(nameHits.map(h => h.securities || h.bank))].join(',')}` : ''}`);

    // 통화 구성
    const ci = L.header.findIndex(h => CURRENCY_HEADERS.includes(E.norm(h)));
    const accounts = C.buildAccounts([r], D);
    const acc = accounts[0];
    p('\n현재 사전 그대로 변환:');
    if (acc.config) convertSubset(acc, L.records.map((_, i) => i), D, '전체');
    if (ci >= 0) {
      const groups = {};
      L.records.forEach((rec, i) => { const c = E.text(rec[ci]).trim().toUpperCase(); (groups[c] = groups[c] || []).push(i); });
      p(`\n통화 열 "${L.header[ci]}": ${Object.entries(groups).map(([c, ix]) => `${c || '(빈칸)'} ${ix.length}건`).join(', ')}`);
      const foreignCodes = Object.keys(groups).filter(c => !DOMESTIC_CURRENCY.includes(c));
      if (foreignCodes.length && acc.config) {
        p('통화별 분리 변환(국내=국내 사전, 외화=해외 사전):');
        const domIdx = Object.keys(groups).filter(c => DOMESTIC_CURRENCY.includes(c)).flatMap(c => groups[c]).sort((a, b) => a - b);
        if (domIdx.length) convertSubset(acc, domIdx, D, '원화');
        const Df = JSON.parse(JSON.stringify(D)); Df.securities.domestic = Df.securities.foreign;
        foreignCodes.forEach(c => convertSubset(acc, groups[c], Df, c));
        // 한 행에 원화·외화가 함께 움직이는 행(환전 등) 찾기
        const amt = L.header.indexOf('거래금액'); const famt = L.header.findIndex(h => /\(외\)$/.test(h) && /거래금액/.test(h));
        if (amt >= 0 && famt >= 0) {
          const both = L.records.map((rec, i) => i).filter(i => (E.parseAmount(L.records[i][amt]) || 0) !== 0 && (E.parseAmount(L.records[i][famt]) || 0) !== 0);
          p(`원화·외화가 한 행에서 함께 움직이는 행(환전 등): ${both.length ? both.map(i => `${L.sourceRows[i] + 1}행`).join(', ') : '없음'}`);
        }
      }
    } else p('\n통화 열 없음');
  });
}

const files = process.argv.slice(2);
if (!files.length) { console.log('사용: node tools/analyze-account.js <파일경로> [...]'); process.exit(2); }
files.forEach(f => { try { analyze(path.resolve(f)); } catch (e) { p(`\n${f}: 분석 실패 — ${e.message}`); } });
console.log(out.join('\n'));
