// 기존 증권 도구 복사본(reference/stockbalance/index.html)에서 "실제로 실행되는 최종본" 변환 코드를
// 손대지 않고 잘라내 engine/securities-legacy.js 로 만든다. (화면 코드 UIRenderer/AppState 는 제외)
//  - const DataTransformer = { ... }            (기본 정의)
//  - DataTransformer.xxx = ...                  (뒤에서 다시 정의한 최종본, 원래 순서 그대로)
//  - const BrokerAutoMapper = { ... }            (증권사 사전·판별·전처리)
//  - AppController._buildPhase5AutoClassifyConfig (자동 분류 전 사전값 합치기) → 독립 함수로
// 실행: node tools/extract-securities-engine.js
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const root = path.resolve(__dirname, '..');
const srcPath = path.join(root, 'reference', 'stockbalance', 'index.html');
const outPath = path.join(root, 'engine', 'securities-legacy.js');
const html = fs.readFileSync(srcPath, 'utf8');
const sha = crypto.createHash('sha256').update(fs.readFileSync(srcPath)).digest('hex');

const scriptStart = html.lastIndexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const lines = html.slice(scriptStart + '<script>'.length, scriptEnd).split(/\r?\n/);

// 4칸 들여쓰기 + 식별자로 시작하는 줄 = 최상위 문장의 시작
const isTop = l => /^ {4}[A-Za-z_$]/.test(l);
const statements = [];
for (let i = 0; i < lines.length; i++) {
  if (!isTop(lines[i])) continue;
  let j = i + 1;
  while (j < lines.length && !isTop(lines[j])) j++;
  let end = j - 1;
  while (end > i && /^\s*(\/\/.*)?$/.test(lines[end])) end--;
  statements.push({ start: i, text: lines.slice(i, end + 1).join('\n') });
}

const pick = re => statements.filter(s => re.test(s.text.split('\n')[0]));
const dt = pick(/^ {4}const DataTransformer = \{/);
const bam = pick(/^ {4}const BrokerAutoMapper = \{/);
const dtAssign = pick(/^ {4}DataTransformer\.[A-Za-z_$][\w$]* = /);
const autoCfg = pick(/^ {4}AppController\._buildPhase5AutoClassifyConfig = /);
if (dt.length !== 1 || bam.length !== 1 || autoCfg.length !== 1 || dtAssign.length < 10) {
  throw new Error(`추출 대상 개수가 예상과 다릅니다: DataTransformer ${dt.length}, BrokerAutoMapper ${bam.length}, 재정의 ${dtAssign.length}, autoCfg ${autoCfg.length}`);
}

const body = [dt[0], ...dtAssign].map(s => s.text).join('\n\n') + '\n\n' + bam[0].text + '\n\n' +
  autoCfg[0].text.replace(/^ {4}AppController\._buildPhase5AutoClassifyConfig = /, '    const buildAutoClassifyConfig = ');

const forbidden = body.match(/\b(AppState|AppController)\b|\bdocument\./g);
if (forbidden) throw new Error(`화면 코드 참조가 섞였습니다: ${[...new Set(forbidden)].join(', ')}`);
// 옛 버전 함수(뒤에서 재정의되어 실제로는 쓰이지 않음) 안에 UIRenderer.showModal 호출이 남아 있다.
// 혹시 호출되면 멈추지 않고 경고만 남기도록 대체 객체를 둔다.
const uiStub = `    const UIRenderer = {
        showModal(title, message) { (root.LegacySecuritiesNotices = root.LegacySecuritiesNotices || []).push({ title, message }); console.warn('[legacy notice]', title, message); },
        hideModal() {}
    };`;

const out = `// 자동 생성 파일 — 직접 고치지 마세요. 원본: reference/stockbalance/index.html (SHA256 ${sha})
// 생성: node tools/extract-securities-engine.js  (${new Date().toISOString()})
// 포함: DataTransformer 기본 정의 + 재정의 최종본 ${dtAssign.length}개, BrokerAutoMapper, buildAutoClassifyConfig
(function (root) {
${uiStub}

${body}

    const api = { DataTransformer, BrokerAutoMapper, buildAutoClassifyConfig, sourceSha256: '${sha}' };
    root.LegacySecurities = api;
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
`;
fs.writeFileSync(outPath, out, 'utf8');
console.log(`생성: ${outPath}`);
console.log(`DataTransformer 재정의 ${dtAssign.length}개: ${dtAssign.map(s => s.text.match(/DataTransformer\.([\w$]+)/)[1]).join(', ')}`);
