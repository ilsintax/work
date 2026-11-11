// 기존 증권 도구 복사본(reference/stockbalance/index.html)의 BrokerAutoMapper 사전을 읽어
// dictionary/institutions.js 를 처음 한 번 생성한다.
// 이미 사전 파일이 있으면 덮어쓰지 않는다(사전은 이후 프로그램에서 계속 갱신되므로). 강제 재생성: --force
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const sourcePath = path.join(root, 'reference', 'stockbalance', 'index.html');
const outPath = path.join(root, 'dictionary', 'institutions.js');
const force = process.argv.includes('--force');

if (fs.existsSync(outPath) && !force) {
  console.error(`이미 사전 파일이 있습니다: ${outPath}\n덮어쓰려면 --force 를 붙이세요(먼저 version/ 백업).`);
  process.exit(1);
}

const html = fs.readFileSync(sourcePath, 'utf8');
const start = html.indexOf('presets: [', html.indexOf('const BrokerAutoMapper'));
const end = html.indexOf('normalizeText(value)', start);
if (start < 0 || end < 0) throw new Error('BrokerAutoMapper 사전 위치를 찾지 못했습니다.');
const literal = html.slice(start, end).trim().replace(/,\s*$/, '');
// eslint-disable-next-line no-new-func
const extracted = new Function(`return ({ ${literal} });`)();
if (!Array.isArray(extracted.presets) || !Array.isArray(extracted.foreignPresets)) {
  throw new Error('presets / foreignPresets 형식이 예상과 다릅니다.');
}

// 파일명 약칭 → 기관명. 같은 약칭이 은행·증권에 모두 있으면 열 구조(은행형/증권형)로 최종 결정한다.
const abbreviations = [
  { keys: ['KB', '국민', 'KB국민', '국민은행', 'KB국민은행'], bank: 'KB국민은행', securities: 'KB증권' },
  { keys: ['신한', '신한은행'], bank: '신한은행', securities: '신한투자증권' },
  { keys: ['NH', '농협', 'NH농협', '농협은행'], bank: 'NH농협은행', securities: 'NH투자증권' },
  { keys: ['IBK', '기업', '기업은행'], bank: 'IBK기업은행', securities: 'IBK투자증권' },
  { keys: ['IM', 'iM', '하이', '대구', 'iM뱅크'], bank: 'iM뱅크', securities: 'iM증권' },
  { keys: ['하나', '하나은행'], bank: '하나은행', securities: '하나증권' },
  { keys: ['우리', '우리은행'], bank: '우리은행', securities: '우리투자증권' },
  { keys: ['SC', '제일', 'SC제일'], bank: 'SC제일은행', securities: '한국스탠다드차타드증권' },
  { keys: ['토스'], bank: '토스뱅크', securities: '토스증권' },
  { keys: ['카카오'], bank: '카카오뱅크', securities: '카카오페이증권' },
  { keys: ['미래', '미래에셋'], securities: '미래에셋증권' },
  { keys: ['한투', '한국투자'], securities: '한국투자증권' },
  { keys: ['키움'], securities: '키움증권' },
  { keys: ['대신'], securities: '대신증권' },
  { keys: ['교보'], securities: '교보증권' },
  { keys: ['삼성'], securities: '삼성증권' },
  { keys: ['DB', '디비'], securities: 'DB금융투자' },
  { keys: ['한화'], securities: '한화투자증권' },
  { keys: ['현대차', '현대'], securities: '현대차증권' },
  { keys: ['메리츠'], securities: '메리츠증권' },
  { keys: ['유진'], securities: '유진투자증권' },
  { keys: ['SK'], securities: 'SK증권' },
  { keys: ['LS', '이베스트'], securities: 'LS증권' },
  { keys: ['유안타'], securities: '유안타증권' },
  { keys: ['신영'], securities: '신영증권' },
  { keys: ['부국'], securities: '부국증권' },
  { keys: ['다올'], securities: '다올투자증권' },
  { keys: ['BNK'], securities: 'BNK투자증권' },
  { keys: ['한양'], securities: '한양증권' },
  { keys: ['케이뱅크'], bank: '케이뱅크' },
  { keys: ['수협'], bank: 'Sh수협은행' },
  { keys: ['부산'], bank: '부산은행' },
  { keys: ['경남'], bank: '경남은행' },
  { keys: ['광주'], bank: '광주은행' },
  { keys: ['전북'], bank: '전북은행' },
  { keys: ['제주'], bank: '제주은행' },
  { keys: ['우체국'], bank: '우체국' },
  { keys: ['새마을'], bank: '새마을금고' },
  { keys: ['신협'], bank: '신협' },
  { keys: ['씨티'], bank: '한국씨티은행' }
];

// 은행 열 이름 사전. 은행은 대부분 입금/출금/잔액 열이 분리되어 있어 공통 규칙으로 찾고,
// 은행별로 다른 이름만 여기에 추가한다.
const bank = {
  columns: {
    dateCol: ['거래일시', '거래일자', '거래일', '일자', '날짜', '거래 일시'],
    depositCol: ['입금액(원)', '입금액', '입금', '맡기신금액', '받으신금액', '입금금액'],
    withdrawalCol: ['출금액(원)', '출금액', '출금', '찾으신금액', '지급금액', '출금금액'],
    balanceCol: ['잔액(원)', '잔액', '거래후잔액', '거래 후 잔액', '잔고'],
    signedAmountCol: ['거래금액', '거래 금액'],
    memoCols: ['보낸분/받는분', '의뢰인/수취인', '기재내용', '거래내용', '받는분/보낸분', '내용', '적요', '내 통장 표시', '메모']
  },
  profiles: [
    {
      id: 'kbstar', name: 'KB국민은행', aliases: ['KB국민은행', '국민은행'],
      columns: { dateCol: ['거래일시'], depositCol: ['입금액(원)'], withdrawalCol: ['출금액(원)'], balanceCol: ['잔액(원)'], memoCols: ['보낸분/받는분'] }
    },
    {
      id: 'toss', name: '토스뱅크', aliases: ['토스'],
      columns: { dateCol: ['거래 일시'], signedAmountCol: ['거래 금액'], balanceCol: ['거래 후 잔액'], memoCols: ['적요'] }
    }
  ]
};

const dictionary = {
  version: 1,
  updatedAt: new Date().toISOString(),
  source: 'reference/stockbalance/index.html BrokerAutoMapper (2026-11-10 추출)',
  abbreviations,
  bank,
  securities: { domestic: extracted.presets, foreign: extracted.foreignPresets },
  accounts: {}
};

const body = JSON.stringify(dictionary, null, 2);
const content =
  '// 통합 통장내역 편집기 기관사전. 프로그램의 "사전 반영" 기능이 이 파일을 다시 씁니다.\n' +
  '// 아래 { } 안은 순수 JSON 입니다.\n' +
  `(function (root) { root.BANKBOOK_DICTIONARY = ${body}; })(typeof window !== 'undefined' ? window : globalThis);\n`;

fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, content, 'utf8');
console.log(`생성: ${outPath}`);
console.log(`국내 증권사 ${extracted.presets.length}곳 (프로필 ${extracted.presets.reduce((n, p) => n + (p.profiles || []).length, 0)}개), 해외 ${extracted.foreignPresets.length}곳, 약칭 ${abbreviations.length}개`);
