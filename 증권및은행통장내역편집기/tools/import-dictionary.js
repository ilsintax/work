// 사전 가져오기: 화면에서 "사전 파일 저장"으로 내려받은 institutions.js 를 프로젝트 사전(dictionary/institutions.js)에 반영한다.
// 순서: ① 파일 찾기 ② 사전 검사(오류면 중단) ③ 지금 사전과 비교해 바뀐 점 보여 주기 ④ 확인 ⑤ 백업 ⑥ 교체
//
// 실행: node tools/import-dictionary.js [파일] [--yes] [--dry]
//   파일 생략 → 다운로드 폴더에서 가장 최근의 institutions*.js
//   --dry  → 비교만 하고 바꾸지 않음
//   --yes  → 묻지 않고 반영(지금 사전에만 있는 내용이 사라지거나 더 오래된 사전이면 --yes 가 있어도 --force 필요)
//   --force → 사라지는 내용·오래된 버전이 있어도 반영
// 종료 코드: 0 = 반영(또는 미리보기) 완료, 1 = 오류·취소
'use strict';
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const readline = require('readline');

const root = path.resolve(__dirname, '..');
require(path.join(root, 'engine', 'engine.js'));
const DICT = require(path.join(root, 'engine', 'dictionary.js'));
const args = process.argv.slice(2);
// --target=경로 : 시험용(기본은 프로젝트 사전). 백업은 대상 파일의 프로젝트 version/ 대신 대상 옆 version/ 에 만든다.
const targetArg = (args.find(a => a.startsWith('--target=')) || '').slice('--target='.length);
const target = targetArg ? path.resolve(targetArg) : path.join(root, 'dictionary', 'institutions.js');
const backupRoot = targetArg ? path.join(path.dirname(target), 'version') : path.join(root, 'version');
const has = f => args.includes(f);
const pause = has('--pause'); // 더블클릭 실행 시 창이 바로 닫히지 않게
const done = code => {
  if (!pause) process.exit(code);
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  rl.question('\n엔터를 누르면 창이 닫힙니다.', () => { rl.close(); process.exit(code); });
};

// ---- ① 파일 찾기
function findDownloaded() {
  const dirs = [path.join(os.homedir(), 'Downloads'), path.join(os.homedir(), '다운로드')].filter(d => fs.existsSync(d));
  const found = dirs.flatMap(d => fs.readdirSync(d).filter(f => /^institutions.*\.js$/i.test(f)).map(f => path.join(d, f)));
  return found.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0] || null;
}
const fileArg = args.find(a => !a.startsWith('--'));
const source = fileArg ? path.resolve(fileArg) : findDownloaded();
if (!source || !fs.existsSync(source)) {
  console.log('가져올 사전 파일을 찾지 못했습니다. 화면에서 "사전 파일 저장"을 누른 뒤 다시 실행하거나, 파일 경로를 지정하세요.');
  done(1); return;
}
if (path.resolve(source) === path.resolve(target)) { console.log('프로젝트 사전 자체를 가리킵니다. 다운로드한 파일을 지정하세요.'); done(1); return; }

// ---- ② 검사
let incoming, current;
try { incoming = DICT.parse(fs.readFileSync(source, 'utf8')); }
catch (e) {
  console.log(`가져올 사전에 오류가 있어 반영하지 않습니다: ${e.message}`);
  (e.errors || []).forEach(x => console.log(`  - ${x}`));
  done(1); return;
}
current = DICT.parse(fs.readFileSync(target, 'utf8'));

// ---- ③ 비교(기관 단위 지식만: 형식·분류·공통 분류·예외 규칙·특수 규칙·은행 형식·은행 이름·약칭)
const added = [], removed = [], changed = [];
const MK = { domestic: '국내', foreign: '해외' };
const B = ['depositTypes', 'withdrawalTypes', 'removalTypes'];
const BL = { depositTypes: '입금', withdrawalTypes: '출금', removalTypes: '제거' };
const profKey = p => JSON.stringify(Object.keys(p.columns || {}).sort().map(k => [k, (p.columns[k] || [])[0]]));
const profName = (b, p, i) => `${b.name} 형식${i + 1}${p.label ? ` "${p.label}"` : ''}`;
const listDiff = (where, a, b) => {
  const sa = new Set(a || []), sb = new Set(b || []);
  const plus = [...sb].filter(x => !sa.has(x)), minus = [...sa].filter(x => !sb.has(x));
  if (plus.length) added.push(`${where}: ${plus.join(', ')}`);
  if (minus.length) removed.push(`${where}: ${minus.join(', ')}`);
};
['domestic', 'foreign'].forEach(m => {
  const cur = current.securities[m] || [], inc = incoming.securities[m] || [];
  inc.filter(b => !cur.some(c => c.name === b.name)).forEach(b => added.push(`${MK[m]} 증권사 "${b.name}"`));
  cur.filter(c => !inc.some(b => b.name === c.name)).forEach(c => removed.push(`${MK[m]} 증권사 "${c.name}"(형식 ${(c.profiles || []).length}개)`));
  cur.forEach(c => {
    const b = inc.find(x => x.name === c.name);
    if (!b) return;
    const cp = c.profiles || [], bp = b.profiles || [];
    // 형식은 1:1로 짝지음(같은 열 구성의 형식이 여러 개일 수 있어 개수까지 비교). 같은 번호를 먼저 짝짓는다.
    const used = new Set();
    bp.forEach((p, i) => {
      let j = cp[i] && !used.has(i) && profKey(cp[i]) === profKey(p) ? i : cp.findIndex((q, k) => !used.has(k) && profKey(q) === profKey(p));
      const old = j >= 0 ? cp[j] : null;
      if (old) used.add(j);
      if (!old) { added.push(`${MK[m]} ${profName(b, p, i)}(열 ${Object.keys(p.columns || {}).length}개)`); return; }
      B.forEach(k => listDiff(`${MK[m]} ${profName(b, p, i)} ${BL[k]}`, old[k], p[k]));
      if ((old.label || '') !== (p.label || '')) changed.push(`${MK[m]} ${b.name} 형식${i + 1} 이름: "${old.label || ''}" → "${p.label || ''}"`);
      if ((old.inferredFeeMode || 'separate') !== (p.inferredFeeMode || 'separate')) changed.push(`${MK[m]} ${profName(b, p, i)} 수수료 처리: ${old.inferredFeeMode || 'separate'} → ${p.inferredFeeMode || 'separate'}`);
      if (JSON.stringify(old.memo || null) !== JSON.stringify(p.memo || null)) changed.push(`${MK[m]} ${profName(b, p, i)} 추가 적요 설정 변경`);
      if (JSON.stringify(old.specialRules || null) !== JSON.stringify(p.specialRules || null)) changed.push(`${MK[m]} ${profName(b, p, i)} 특수 규칙 변경`);
    });
    cp.forEach((q, i) => { if (!used.has(i)) removed.push(`${MK[m]} ${profName(c, q, i)}`); });
    B.forEach(k => listDiff(`${MK[m]} ${b.name} 공통 분류 ${BL[k]}`, (c.shared || {})[k], (b.shared || {})[k]));
    listDiff(`${MK[m]} ${b.name} 예외 규칙`, (c.exceptions || []).map(x => x.id), (b.exceptions || []).map(x => x.id));
    if (JSON.stringify(c.specialRules || null) !== JSON.stringify(b.specialRules || null)) changed.push(`${MK[m]} ${b.name} 증권사 특수 규칙 변경`);
  });
});
listDiff('은행 형식', current.bank.profiles.map(p => `${p.name}(${(p.columns.memoCols || []).join('+')})`), incoming.bank.profiles.map(p => `${p.name}(${(p.columns.memoCols || []).join('+')})`));
listDiff('은행 이름 목록', current.bank.names, incoming.bank.names);
listDiff('파일명 약칭', current.abbreviations.flatMap(a => a.keys.map(k => `${k}→${a.securities || a.bank}`)), incoming.abbreviations.flatMap(a => a.keys.map(k => `${k}→${a.securities || a.bank}`)));
if (JSON.stringify(current.securities.keywordRules || null) !== JSON.stringify(incoming.securities.keywordRules || null)) changed.push('공통 키워드 규칙 변경');

const curV = current.version || 1, incV = incoming.version || 1;
console.log(`가져올 파일: ${source}`);
console.log(`  (저장 시각 ${incoming.updatedAt || '-'}, 사전 버전 ${incV})`);
console.log(`프로젝트 사전: 버전 ${curV}${current.updatedAt ? ` (${current.updatedAt})` : ''}`);
const show = (title, list) => { if (list.length) { console.log(`\n${title} ${list.length}건`); list.forEach(x => console.log(`  - ${x}`)); } };
show('[추가]', added);
show('[변경]', changed);
show('[사라짐 — 지금 사전에만 있음]', removed);
const newLog = (incoming.changelog || []).filter(c => !(current.changelog || []).some(o => o.version === c.version && o.at === c.at));
if (newLog.length) { console.log('\n가져올 사전의 변경 기록(새 항목)'); newLog.forEach(c => c.changes.forEach(x => console.log(`  · v${c.version} ${x}`))); }
const warnings = (incoming.loadInfo && incoming.loadInfo.warnings) || [];
if (warnings.length) console.log(`\n사전 경고 ${warnings.length}건(사용 가능)`);

if (!added.length && !changed.length && !removed.length) { console.log('\n바뀐 내용이 없습니다. 반영하지 않습니다.'); done(0); return; }
const risky = [];
if (removed.length) risky.push(`지금 사전에만 있는 내용 ${removed.length}건이 사라집니다(다른 사람이 반영한 내용일 수 있음)`);
if (incV < curV) risky.push(`가져올 사전(버전 ${incV})이 지금 사전(버전 ${curV})보다 오래됐습니다`);
if (risky.length) { console.log('\n주의:'); risky.forEach(x => console.log(`  ! ${x}`)); }
if (has('--dry')) { console.log('\n미리보기만 했습니다(--dry). 바꾸지 않았습니다.'); done(0); return; }

// ---- ④ 확인 ⑤ 백업 ⑥ 교체
function apply() {
  const stamp = new Date(Date.now() + 9 * 3600 * 1000).toISOString().replace(/[-:]/g, '').replace('T', '_').slice(0, 15); // KST
  const dir = path.join(backupRoot, `${stamp}_사전가져오기전`, 'files', 'dictionary');
  fs.mkdirSync(dir, { recursive: true });
  const bak = path.join(dir, 'institutions.js');
  fs.copyFileSync(target, bak);
  const sha = f => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
  if (sha(bak) !== sha(target)) { console.log('백업 확인 실패 — 반영하지 않습니다.'); done(1); return; }
  fs.writeFileSync(path.join(dir, '..', '..', 'manifest.json'), JSON.stringify({ task: '사전 가져오기', source, backup: 'files/dictionary/institutions.js', sha256: sha(bak), at: new Date().toISOString() }, null, 2));
  fs.writeFileSync(path.join(dir, '..', '..', 'RESTORE.md'), `되돌리기: files/dictionary/institutions.js 를 프로젝트의 dictionary/institutions.js 로 복사\n`);
  const out = DICT.serialize(incoming).replace(/^\/\/ .*\n/, '// 통합 통장내역 편집기 기관사전.\n');
  fs.writeFileSync(target, out, 'utf8');
  const back = DICT.parse(fs.readFileSync(target, 'utf8'));
  console.log(`\n반영했습니다: dictionary/institutions.js (버전 ${back.version || 1})`);
  console.log(`이전 사전 백업: ${path.relative(root, path.join(dir, '..', '..'))}`);
  console.log('GitHub 에는 올리지 않았습니다. 팀에 적용하려면 따로 업로드하세요.');
  done(0);
}
if (has('--yes') && (!risky.length || has('--force'))) { apply(); return; }
if (has('--yes') && risky.length) { console.log('\n위 주의 사항이 있어 --yes 만으로는 반영하지 않습니다. 확인 후 --force 를 함께 쓰세요.'); done(1); return; }
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
rl.question(`\n${risky.length ? '주의 사항을 확인했고 ' : ''}프로젝트 사전에 반영할까요? (예: y / 아니오: 엔터) `, ans => {
  rl.close();
  if (/^(y|yes|예|ㅇ)$/i.test(String(ans).trim())) apply();
  else { console.log('취소했습니다. 바꾸지 않았습니다.'); done(1); }
});
