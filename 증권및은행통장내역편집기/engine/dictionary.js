// 통합 통장내역 편집기 — 기관사전 반영·저장·불러오기
// 원칙: 사전에는 "기관 단위 지식"만 저장한다(열 이름, 거래종류 분류, 적요 열, 수수료 처리 방식, 파일명 약칭).
//       계좌번호·고객명·금액 같은 계좌별 정보는 저장하지 않는다(GitHub 공유 대비).
(function (root) {
  'use strict';
  const E = root.BankbookEngine || (typeof require !== 'undefined' ? require('./engine.js') : null);
  const SEC_KEYS = ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol', 'feeCol', 'taxCol', 'extraCostCol1', 'extraCostCol2', 'extraCostCol3',
    'stockNameCol', 'stockQtyCol', 'unitPriceCol', 'stockBalanceCol', 'kbDepositAmountCol', 'kbWithdrawalAmountCol', 'settlementCol'];
  const REQUIRED_SEC = ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'];
  const BANK_KEYS = ['dateCol', 'depositCol', 'withdrawalCol', 'balanceCol', 'signedAmountCol'];

  // 사전 구조 버전. version(내용 개정 번호)과 별개로, 사전의 "모양"이 바뀔 때만 올린다.
  //  1: 기존 도구에서 추출한 형태(형식별 분류만)
  //  2: 증권사 단위 처리 방식(shared 분류·specialRules), 공통 키워드 규칙(securities.keywordRules), 국내/해외 분리 명시
  const SCHEMA_VERSION = 2;
  const BUCKETS = ['depositTypes', 'withdrawalTypes', 'removalTypes'];
  const BUCKET_LABEL = { depositTypes: '입금', withdrawalTypes: '출금', removalTypes: '제거' };
  const MARKETS = ['domestic', 'foreign'];
  const MARKET_LABEL = { domestic: '국내', foreign: '해외' };
  // 기존 증권 도구(engine/securities-legacy.js)에 고정돼 있던 키워드 목록과 같은 값. 사전에서 고칠 수 있게 옮겨 둔다.
  const DEFAULT_KEYWORD_RULES = {
    directionHints: { deposit: ['매도', '입금', '배당', '환급', '입고', '이입'], withdrawal: ['매수', '출금', '출고', '수수료', '세금', '제출', '이체'] },
    removalHints: ['수수료', '세금', '제세금', '기타수수료', '기타제세금', '이체', '대체', '정산', '환전', '출고', '연체료'],
    blockDeposit: ['수수료', '세금', '제세금', '기타수수료', '기타제세금', '출금', '출고', '연체료', '대출이자'],
    mirrorPatterns: [
      { name: 'sell_summary_to_detail', detailKeywords: ['주식매도출고', '매도출고'], summaryKeywords: ['주식매도입금', '매도입금'], detailCategory: 'deposit', summaryCategory: 'remove' },
      { name: 'buy_summary_to_detail', detailKeywords: ['주식매수입고', '매수입고'], summaryKeywords: ['주식매수출금', '매수출금'], detailCategory: 'withdrawal', summaryCategory: 'remove' }
    ]
  };
  const SPECIAL_RULE_KEYS = ['ignoreTradeTypeContains', 'amountOverrideByTradeType', 'forceSeparateCosts', 'deriveAmountFromTradeTypeText',
    'classificationKeywords', 'unmatchedClassification', 'notice'];

  const clone = o => JSON.parse(JSON.stringify(o));
  const typeKey = v => E.norm(v).replace(/\s+/g, '');
  const uniq = arr => [...new Set(arr.filter(v => v !== undefined && v !== null && String(v).trim() !== ''))];
  const slug = s => 'custom-' + Array.from(String(s)).map(ch => ch.charCodeAt(0).toString(36)).join('').slice(0, 24);
  const sameColumns = (a, b) => {
    const ka = Object.keys(a || {}).sort(), kb = Object.keys(b || {}).sort();
    return ka.join('|') === kb.join('|') && ka.every(k => E.norm((a[k] || [])[0]) === E.norm((b[k] || [])[0]));
  };

  // 파일명 약칭 → 기관 연결 (이미 같은 결과가 나오면 바꾸지 않음)
  function applyAbbreviation(dict, account, changes) {
    const abbr = account.file && account.file.abbr;
    const inst = (account.institution || '').trim();
    if (!abbr || !inst) return;
    const field = account.kind === 'bank' ? 'bank' : 'securities';
    const hit = E.resolveAbbreviation(abbr, dict);
    if (hit && hit[field] === inst) return;
    const entry = (dict.abbreviations || []).find(e => e.keys.some(k => E.norm(k).toUpperCase() === E.norm(abbr).toUpperCase()));
    if (entry) {
      if (entry[field] && entry[field] !== inst) changes.push(`약칭 "${abbr}"의 ${field === 'bank' ? '은행' : '증권사'}: ${entry[field]} → ${inst}`);
      else changes.push(`약칭 "${abbr}" → ${inst}`);
      entry[field] = inst;
    } else {
      dict.abbreviations.push({ keys: [abbr], [field]: inst });
      changes.push(`새 약칭 "${abbr}" → ${inst}`);
    }
  }

  // ---------------------------------------------------------------- 구조 변환(1→2)
  // 증권사의 모든 형식에서 같은 칸으로 분류된 거래종류만 증권사 공통(shared)으로 올린다. 형식끼리 엇갈리는 값은 올리지 않는다.
  // 간편 형식(classificationKeywords, 포함 검색)은 정확히 같은 이름 목록이 아니라서 제외한다.
  function seedShared(broker) {
    const seen = {};
    (broker.profiles || []).filter(p => !(p.specialRules && p.specialRules.classificationKeywords)).forEach(p => BUCKETS.forEach(b => (p[b] || []).forEach(t => {
      const k = typeKey(t); if (!k) return;
      (seen[k] = seen[k] || { value: String(t).trim(), buckets: new Set() }).buckets.add(b);
    })));
    const shared = { depositTypes: [], withdrawalTypes: [], removalTypes: [] };
    Object.values(seen).forEach(s => { if (s.buckets.size === 1) shared[[...s.buckets][0]].push(s.value); });
    BUCKETS.forEach(b => shared[b].sort());
    return shared;
  }

  // 오래된 구조의 사전을 현재 구조로 바꾼다(제자리 변경). 반환: 바꾼 내용 설명 목록
  function migrate(dict) {
    const notes = [];
    const sv = Number(dict.schemaVersion || 1);
    if (sv > SCHEMA_VERSION) throw new Error(`이 사전은 더 새로운 구조(${sv})로 만들어졌습니다. 프로그램을 최신으로 받은 뒤 불러오세요(현재 프로그램은 ${SCHEMA_VERSION}까지 지원).`);
    if (sv < 2) {
      const sec = dict.securities = dict.securities || {};
      MARKETS.forEach(m => { sec[m] = Array.isArray(sec[m]) ? sec[m] : []; });
      if (!sec.keywordRules) { sec.keywordRules = clone(DEFAULT_KEYWORD_RULES); notes.push('공통 키워드 규칙(securities.keywordRules) 추가'); }
      let sharedCount = 0, hoisted = [];
      MARKETS.forEach(m => sec[m].forEach(b => {
        b.profiles = Array.isArray(b.profiles) ? b.profiles : [];
        if (!b.shared) { b.shared = seedShared(b); sharedCount += BUCKETS.reduce((n, k) => n + b.shared[k].length, 0); }
        // 모든 형식에 똑같은 특수 규칙이 있으면 증권사 단위 규칙으로 올린다(새 형식도 그대로 이어받게)
        const rules = b.profiles.map(p => JSON.stringify(p.specialRules || null));
        if (!b.specialRules && b.profiles.length && rules[0] !== 'null' && rules.every(r => r === rules[0])) {
          b.specialRules = JSON.parse(rules[0]);
          b.profiles.forEach(p => { delete p.specialRules; });
          hoisted.push(`${MARKET_LABEL[m]} ${b.name}`);
        }
      }));
      notes.push(`증권사 공통 분류(shared) ${sharedCount}개 생성`);
      if (hoisted.length) notes.push(`증권사 단위 특수 규칙으로 올림: ${hoisted.join(', ')}`);
      dict.schemaVersion = 2;
    }
    return notes;
  }

  // ---------------------------------------------------------------- 검사
  // 반환: { errors: [...], warnings: [...] }. errors 가 있으면 그 사전은 쓰지 않는다.
  function validate(dict) {
    const errors = [], warnings = [];
    const isStrList = v => Array.isArray(v) && v.every(x => typeof x === 'string');
    if (!dict || typeof dict !== 'object') return { errors: ['사전이 비어 있습니다.'], warnings };
    if (Number(dict.schemaVersion || 1) !== SCHEMA_VERSION) errors.push(`사전 구조 버전이 ${dict.schemaVersion || 1}입니다(필요: ${SCHEMA_VERSION}).`);
    if (!Array.isArray(dict.abbreviations)) errors.push('abbreviations(파일명 약칭) 목록이 없습니다.');
    if (!dict.bank || !Array.isArray(dict.bank.profiles)) errors.push('bank.profiles(은행 형식) 목록이 없습니다.');
    if (dict.bank && dict.bank.names !== undefined && !(Array.isArray(dict.bank.names) && dict.bank.names.every(x => typeof x === 'string'))) errors.push('bank.names(은행 이름 목록)는 글자 목록이어야 합니다.');
    const sec = dict.securities;
    if (!sec || !MARKETS.every(m => Array.isArray(sec[m]))) { errors.push('securities.domestic / securities.foreign 목록이 없습니다.'); return { errors, warnings }; }

    const kr = sec.keywordRules;
    if (!kr) warnings.push('공통 키워드 규칙(securities.keywordRules)이 없어 기존 도구의 기본값을 씁니다.');
    const checkKeywordRules = (r, where) => {
      if (!r) return;
      if (r.directionHints && !(isStrList(r.directionHints.deposit || []) && isStrList(r.directionHints.withdrawal || []))) errors.push(`${where} directionHints 는 글자 목록이어야 합니다.`);
      ['removalHints', 'blockDeposit'].forEach(k => { if (r[k] !== undefined && !isStrList(r[k])) errors.push(`${where} ${k} 는 글자 목록이어야 합니다.`); });
      if (r.mirrorPatterns !== undefined && !(Array.isArray(r.mirrorPatterns) && r.mirrorPatterns.every(p => p && isStrList(p.detailKeywords) && isStrList(p.summaryKeywords)))) errors.push(`${where} mirrorPatterns 형식이 잘못됐습니다.`);
    };
    checkKeywordRules(kr, '공통 키워드 규칙');

    // 반환: 거래종류 → 실제 적용되는 칸. 기존 엔진 규칙(buildAutoClassifyConfig)은 제거 > 입금 > 출금 순으로 앞선 쪽이 이긴다.
    // 동작은 정해져 있으므로 여러 칸에 있으면 경고만 한다.
    const PRIORITY = ['removalTypes', 'depositTypes', 'withdrawalTypes'];
    const checkBuckets = (obj, where) => {
      const owner = {};
      PRIORITY.forEach(b => {
        if (obj[b] === undefined) return;
        if (!isStrList(obj[b])) { errors.push(`${where} ${b} 는 글자 목록이어야 합니다.`); return; }
        obj[b].forEach(t => {
          const k = typeKey(t);
          if (owner[k] && owner[k] !== b) warnings.push(`${where}: "${t}"이(가) ${BUCKET_LABEL[owner[k]]}·${BUCKET_LABEL[b]} 두 칸에 모두 있습니다 — ${BUCKET_LABEL[owner[k]]}(으)로 적용됩니다(우선순위: 제거 > 입금 > 출금).`);
          if (!owner[k]) owner[k] = b;
        });
      });
      return owner;
    };
    // 통화 분리(currencySplit): 통화 열, 원화 쪽 금액·잔액 열, 분류 목록
    const checkSplit = (sp, where) => {
      if (sp === undefined) return;
      const w = `${where} 통화 분리`;
      if (!sp || typeof sp !== 'object') { errors.push(`${w} 형식이 잘못됐습니다.`); return; }
      if (!isStrList(sp.currencyCol || []) || !(sp.currencyCol || []).length) errors.push(`${w}: currencyCol(통화 열 이름 목록)이 없습니다.`);
      const bc = (sp.base && sp.base.columns) || {};
      if (!isStrList(bc.amountCol || []) || !(bc.amountCol || []).length || !isStrList(bc.balanceCol || []) || !(bc.balanceCol || []).length) errors.push(`${w}: base.columns 에 원화 쪽 amountCol·balanceCol 이 있어야 합니다.`);
      Object.entries({ ...bc, ...((sp.foreign && sp.foreign.columns) || {}) }).forEach(([k, v]) => { if (!isStrList(v) || !v.length) errors.push(`${w}: 열 ${k} 의 이름 목록이 비었거나 잘못됐습니다.`); });
      if (sp.base) checkBuckets(sp.base, `${w} 원화 쪽`);
    };
    const checkRules = (r, where) => {
      if (!r) return;
      if (typeof r !== 'object') { errors.push(`${where} specialRules 형식이 잘못됐습니다.`); return; }
      Object.keys(r).filter(k => !SPECIAL_RULE_KEYS.includes(k)).forEach(k => warnings.push(`${where}: 알 수 없는 특수 규칙 "${k}"(무시됨)`));
      if (r.ignoreTradeTypeContains !== undefined && !isStrList(r.ignoreTradeTypeContains)) errors.push(`${where} ignoreTradeTypeContains 는 글자 목록이어야 합니다.`);
      if (r.amountOverrideByTradeType !== undefined && !(Array.isArray(r.amountOverrideByTradeType) && r.amountOverrideByTradeType.every(x => x && typeof x.match === 'string' && isStrList(x.sourceHeaders)))) errors.push(`${where} amountOverrideByTradeType 형식이 잘못됐습니다.`);
    };

    MARKETS.forEach(m => {
      const ids = new Set(), names = new Set();
      sec[m].forEach((b, bi) => {
        const where = `${MARKET_LABEL[m]} ${b && b.name ? b.name : `${bi + 1}번째 증권사`}`;
        if (!b || !b.id || !b.name) { errors.push(`${where}: id 또는 name 이 없습니다.`); return; }
        if (ids.has(b.id)) errors.push(`${MARKET_LABEL[m]}: 증권사 id "${b.id}" 가 중복됩니다.`); ids.add(b.id);
        if (names.has(b.name)) errors.push(`${MARKET_LABEL[m]}: 증권사 이름 "${b.name}" 이 중복됩니다.`); names.add(b.name);
        if (b.aliases !== undefined && !isStrList(b.aliases)) errors.push(`${where} aliases 는 글자 목록이어야 합니다.`);
        if (!Array.isArray(b.profiles)) { errors.push(`${where}: profiles 목록이 없습니다.`); return; }
        const shared = b.shared ? checkBuckets(b.shared, `${where} 공통 분류`) : {};
        checkRules(b.specialRules, where);
        checkSplit(b.currencySplit, where);
        if (b.exceptions !== undefined) {
          if (!Array.isArray(b.exceptions)) errors.push(`${where} exceptions 는 목록이어야 합니다.`);
          else b.exceptions.forEach((x, xi) => {
            const xw = `${where} 예외 규칙 ${x && x.id ? x.id : xi + 1}`;
            if (!x || !x.id) errors.push(`${xw}: id 가 없습니다.`);
            if (!x || !isStrList(x.signature) || x.signature.length < 3) errors.push(`${xw}: signature(사례 파일의 전체 열 이름)가 없습니다 — 열 이름이 모두 같을 때만 적용하므로 필수입니다.`);
            if (x && x.classificationBasis !== undefined && x.classificationBasis !== 'balance') errors.push(`${xw}: classificationBasis 는 'balance' 만 지원합니다.`);
            if (x) checkSplit(x.currencySplit, xw);
            if (x && !(x.cases >= 1)) warnings.push(`${xw}: 사례 수(cases)가 없습니다.`);
          });
        }
        checkKeywordRules(b.keywordRules, `${where} 키워드 규칙`);
        b.profiles.forEach((p, pi) => {
          const pw = `${where} 형식${pi + 1}`;
          if (!p || typeof p.columns !== 'object' || !p.columns) { errors.push(`${pw}: columns 가 없습니다.`); return; }
          Object.entries(p.columns).forEach(([k, v]) => { if (!isStrList(v) || !v.length) errors.push(`${pw}: 열 ${k} 의 이름 목록이 비었거나 잘못됐습니다.`); });
          const miss = ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'].filter(k => !p.columns[k]);
          if (miss.length && !(miss.length === 1 && miss[0] === 'amountCol' && ((p.specialRules || b.specialRules) || {}).deriveAmountFromTradeTypeText)) {
            warnings.push(`${pw}: 필수 열(${miss.join(', ')})이 없어 열 일치만으로는 확정되지 않습니다.`);
          }
          if (p.label !== undefined && typeof p.label !== 'string') errors.push(`${pw}: label(형식 이름)은 글자여야 합니다.`);
          if (p.signature !== undefined && !isStrList(p.signature)) errors.push(`${pw}: signature 는 글자 목록이어야 합니다.`);
          if (p.inferredFeeMode !== undefined && !['included', 'separate'].includes(p.inferredFeeMode)) errors.push(`${pw}: inferredFeeMode 는 included 또는 separate 여야 합니다.`);
          checkSplit(p.currencySplit, pw);
          if (p.memo !== undefined) {
            if (!p.memo || typeof p.memo !== 'object') errors.push(`${pw}: memo 형식이 잘못됐습니다.`);
            else Object.entries(p.memo).forEach(([k, v]) => {
              if (!['additionalCol1', 'additionalCol2', 'excludedValues'].includes(k)) warnings.push(`${pw}: memo 의 알 수 없는 항목 "${k}"(무시됨)`);
              else if (!isStrList(v)) errors.push(`${pw}: memo.${k} 는 글자 목록이어야 합니다.`);
            });
          }
          const own = checkBuckets(p, pw);
          checkRules(p.specialRules, pw);
          Object.entries(own).forEach(([k, b2]) => { if (shared[k] && shared[k] !== b2) warnings.push(`${pw}: "${k}" 분류가 증권사 공통(${BUCKET_LABEL[shared[k]]})과 다릅니다 — 이 형식에서는 형식 분류(${BUCKET_LABEL[b2]})를 따릅니다.`); });
        });
      });
    });
    (dict.abbreviations || []).forEach((a, i) => {
      if (!a || !isStrList(a.keys) || !a.keys.length) errors.push(`파일명 약칭 ${i + 1}번째: keys 가 비었습니다.`);
      else if (a.securities && !sec.domestic.some(b => b.name === a.securities)) warnings.push(`파일명 약칭 "${a.keys[0]}" → 증권사 "${a.securities}" 가 사전 증권사 목록에 없습니다.`);
    });
    if (dict.accounts) warnings.push('accounts(계좌별 정보) 항목은 저장할 때 빠집니다.');
    return { errors, warnings };
  }

  // 사전 텍스트/객체를 받아 구조 변환 + 검사까지 한 번에. 오류가 있으면 예외(쓰지 않음).
  function prepare(dict) {
    const notes = migrate(dict);
    const result = validate(dict);
    if (result.errors.length) {
      const err = new Error(`사전 검사 오류 ${result.errors.length}건: ${result.errors.slice(0, 3).join(' / ')}${result.errors.length > 3 ? ' …' : ''}`);
      err.errors = result.errors; err.warnings = result.warnings;
      throw err;
    }
    return { dict, notes, warnings: result.warnings };
  }

  function applySecurities(dict, account, changes, opts) {
    const out = account.output || {};
    const mapping = out.mapping || {};
    const header = account.layout.header;
    const cls = account.config.classification;
    if (!cls || !Array.isArray(cls.depositTypes)) throw new Error('거래종류 분류가 아직 없습니다. 열을 지정해 변환을 먼저 완료하세요.');
    const columns = {};
    SEC_KEYS.forEach(k => {
      const idx = mapping[k];
      if (idx !== undefined && idx !== null && header[idx]) columns[k] = [header[idx]];
    });
    const missing = REQUIRED_SEC.filter(k => !columns[k]);
    if (missing.length) throw new Error('일자·거래종류·거래금액·잔액 열이 모두 지정되어야 사전에 반영할 수 있습니다.');
    const name = (account.institution || '').trim();
    if (!name) throw new Error('기관 이름을 입력하세요.');
    // 예외 규칙(잔액 기준 분류) 계좌: 분류는 파일마다 잔액으로 정하므로 사전에 저장하지 않는다
    const exr = account.broker && account.broker.exceptionRule;
    if (exr && exr.classificationBasis === 'balance') throw new Error(`예외 규칙 "${exr.label || exr.id}"을 쓰는 계좌는 입금/출금을 잔액으로 분류하므로 사전에 저장할 분류가 없습니다. 열 구성을 바꾸려면 예외 규칙을 직접 수정하세요(AGENTS.md 예외 계좌 분석 절차).`);
    if (account.currencySide) return applyCurrencySide(dict, account, columns, cls, name, changes);

    // 국내/해외는 서로 다른 사전 목록에 기록한다(같은 증권사·같은 열 이름이라도 해외 계좌는 국내 형식을 쓰지 않음)
    const market = marketOf(account);
    const ml = MARKET_LABEL[market];
    const list = dict.securities[market] = dict.securities[market] || [];
    let broker = list.find(b => b.name === name);
    if (!broker) {
      const twin = dict.securities[market === 'domestic' ? 'foreign' : 'domestic'].find(b => b.name === name);
      broker = { id: twin ? twin.id : slug(name), name, aliases: twin ? clone(twin.aliases || [name]) : [name], profiles: [], shared: { depositTypes: [], withdrawalTypes: [], removalTypes: [] } };
      list.push(broker);
      changes.push(`${ml} 새 증권사 "${name}" 추가`);
    }
    broker.profiles = broker.profiles || [];
    let profile = broker.profiles.find(p => sameColumns(p.columns, columns));
    if (!profile) {
      // signature: 이 파일의 전체 열 이름(형식 지문). 흔한 열 이름만으로 다른 증권사 파일과 혼동되지 않게 판별에 쓴다.
      profile = { columns, signature: [...new Set(header.filter(Boolean))], depositTypes: [], withdrawalTypes: [], removalTypes: [] };
      const lbl = String((opts && opts.label) || '').trim();
      if (lbl) profile.label = lbl.slice(0, 40);
      // 같은 증권사 형식에만 있던 특수 규칙은 이어받는다(증권사 단위 규칙 broker.specialRules 는 자동으로 적용됨)
      const det = account.broker;
      if (det && det.detected && det.profile && det.profile.specialRules && det.brokerId === broker.id && (det.market || 'domestic') === market) {
        profile.specialRules = clone(det.profile.specialRules);
      }
      broker.profiles.push(profile);
      changes.push(`${ml} ${name}: 새 파일 형식${profile.label ? ` "${profile.label}"` : ''}(형식${broker.profiles.length}, 열 ${Object.keys(columns).length}개) 추가`);
    }
    // 5단계: 추가 적요 열과 적요에서 뺄 값은 형식의 memo 에 따로 둔다(판별용 열 columns 와 분리 → 판별에 영향 없음)
    const memoNow = {};
    ['additionalCol1', 'additionalCol2'].forEach(k => { const i = mapping[k]; if (i !== undefined && i !== null && header[i]) memoNow[k] = [header[i]]; });
    const exVals = uniq(['additionalCol1', 'additionalCol2'].flatMap(k => { const i = mapping[k]; return i === undefined || i === null ? [] : (((account.config.excludedMemoValues || {})[i]) || []); })).sort();
    if (exVals.length && Object.keys(memoNow).length) memoNow.excludedValues = exVals;
    const beforeMemo = JSON.stringify(profile.memo || {});
    if (Object.keys(memoNow).length) profile.memo = memoNow; else delete profile.memo;
    if (JSON.stringify(profile.memo || {}) !== beforeMemo) {
      changes.push(`${ml} ${name}: 추가 적요 ${['additionalCol1', 'additionalCol2'].filter(k => memoNow[k]).map(k => memoNow[k][0]).join(' + ') || '없음'}${exVals.length && memoNow.excludedValues ? `, 뺄 값 ${exVals.length}개` : ''}`);
    }
    const keyCount = Object.keys(columns).length;
    if (keyCount < 5) profile.minMatchedKeys = Math.max(4, keyCount); // 기존 판별 규칙은 기본 5개 열 일치 필요

    // 이번 계좌에 나온 거래종류만 다시 배정하고, 다른 계좌에서 배운 값은 유지
    const seen = new Set((out.tradeTypes || []).map(E.norm));
    const keep = arr => (arr || []).filter(v => !seen.has(E.norm(v)));
    const before = JSON.stringify([profile.depositTypes, profile.withdrawalTypes, profile.removalTypes]);
    profile.depositTypes = uniq([...keep(profile.depositTypes), ...cls.depositTypes]).sort();
    profile.withdrawalTypes = uniq([...keep(profile.withdrawalTypes), ...cls.withdrawalTypes]).sort();
    profile.removalTypes = uniq([...keep(profile.removalTypes), ...cls.valuesToRemove]).sort();
    if (JSON.stringify([profile.depositTypes, profile.withdrawalTypes, profile.removalTypes]) !== before) {
      changes.push(`${ml} ${name}: 거래종류 분류 갱신 (입금 ${profile.depositTypes.length} · 출금 ${profile.withdrawalTypes.length} · 제거 ${profile.removalTypes.length})`);
    }
    // 증권사 공통 분류에도 기록 → 같은 증권사의 다른 형식(열 구조가 다른 파일)에서도 이 분류를 쓴다
    broker.shared = broker.shared || { depositTypes: [], withdrawalTypes: [], removalTypes: [] };
    const sh = broker.shared;
    const beforeShared = JSON.stringify(BUCKETS.map(b => sh[b] || []));
    const moved = [];
    const assign = (list2, bucket) => list2.forEach(t => {
      const k = typeKey(t);
      BUCKETS.filter(b => b !== bucket).forEach(b => {
        if ((sh[b] || []).some(x => typeKey(x) === k)) { sh[b] = sh[b].filter(x => typeKey(x) !== k); moved.push(`${t}(${BUCKET_LABEL[b]}→${BUCKET_LABEL[bucket]})`); }
      });
      sh[bucket] = uniq([...(sh[bucket] || []).filter(x => typeKey(x) !== k), t]).sort();
    });
    assign(cls.depositTypes, 'depositTypes'); assign(cls.withdrawalTypes, 'withdrawalTypes'); assign(cls.valuesToRemove, 'removalTypes');
    if (JSON.stringify(BUCKETS.map(b => sh[b] || [])) !== beforeShared) {
      changes.push(`${ml} ${name}: 증권사 공통 분류 갱신${moved.length ? ` (바뀐 분류: ${moved.join(', ')})` : ''}`);
    }
    const fee = cls.inferredFeeMode === 'included' ? 'included' : 'separate';
    if ((profile.inferredFeeMode || 'separate') !== fee) changes.push(`${ml} ${name}: 수수료 처리 → ${fee === 'included' ? '거래금액에 포함' : '별도 행'}`);
    if (fee === 'included') profile.inferredFeeMode = 'included'; else delete profile.inferredFeeMode;
  }

  // 통화별로 나뉜 계좌(3단계): 나눈 기준인 사전 형식 안에 기록한다. 새 형식을 만들지 않는다.
  //  - 원화 쪽(base): 그 형식의 currencySplit.base 분류(입금·출금·제거)와 수수료 처리
  //  - 외화 쪽: 형식 분류 + 증권사 공통 분류(일반 해외 계좌와 같음). 정산금액 열은 currencySplit.foreign 에 둔다.
  function applyCurrencySide(dict, account, columns, cls, name, changes) {
    const det = account.broker || {};
    const market = marketOf(account);
    const broker = (dict.securities[market] || []).find(b => b.id === det.brokerId) || (dict.securities[market] || []).find(b => b.name === name);
    const profile = broker && broker.profiles && broker.profiles[det.profileIndex];
    const spec = profile && profile.currencySplit ? profile.currencySplit : broker && broker.currencySplit;
    if (!broker || !profile || !spec) throw new Error('통화별로 나눈 계좌인데 나눈 기준(사전 형식의 currencySplit)을 찾지 못했습니다.');
    const ml = MARKET_LABEL[market];
    const label = `${ml} ${name} 형식${det.profileIndex + 1} ${account.currency}`;
    const seen = new Set((account.output.tradeTypes || []).map(typeKey));
    const keep = arr => (arr || []).filter(v => !seen.has(typeKey(v)));
    const merge = target => {
      const before = JSON.stringify([target.depositTypes || [], target.withdrawalTypes || [], target.removalTypes || [], target.inferredFeeMode || 'separate']);
      target.depositTypes = uniq([...keep(target.depositTypes), ...cls.depositTypes]).sort();
      target.withdrawalTypes = uniq([...keep(target.withdrawalTypes), ...cls.withdrawalTypes]).sort();
      target.removalTypes = uniq([...keep(target.removalTypes), ...cls.valuesToRemove]).sort();
      if (cls.inferredFeeMode === 'included') target.inferredFeeMode = 'included'; else delete target.inferredFeeMode;
      return JSON.stringify([target.depositTypes || [], target.withdrawalTypes || [], target.removalTypes || [], target.inferredFeeMode || 'separate']) !== before;
    };
    if (account.currencySide === 'base') {
      spec.base = spec.base || { columns: {} };
      const cols = Object.fromEntries(Object.entries(columns).filter(([k]) => ['amountCol', 'balanceCol', 'feeCol', 'taxCol', 'settlementCol'].includes(k)));
      if (JSON.stringify(cols) !== JSON.stringify(spec.base.columns || {})) { spec.base.columns = cols; changes.push(`${label}: 원화 쪽 금액·잔액 열 갱신`); }
      if (merge(spec.base)) changes.push(`${label}: 원화 쪽 거래종류 분류 갱신 (입금 ${spec.base.depositTypes.length} · 출금 ${spec.base.withdrawalTypes.length} · 제거 ${spec.base.removalTypes.length})`);
      return;
    }
    if (columns.settlementCol) {
      spec.foreign = spec.foreign || { columns: {} };
      if (JSON.stringify(spec.foreign.columns.settlementCol) !== JSON.stringify(columns.settlementCol)) { spec.foreign.columns.settlementCol = columns.settlementCol; changes.push(`${label}: 정산금액 열(숨은 비용 계산) 갱신`); }
    }
    if (merge(profile)) changes.push(`${label}: 거래종류 분류 갱신 (입금 ${profile.depositTypes.length} · 출금 ${profile.withdrawalTypes.length} · 제거 ${profile.removalTypes.length})`);
  }

  function applyBank(dict, account, changes) {
    const c = account.config;
    const header = account.layout.header;
    const name = (account.institution || '').trim();
    if (!name) throw new Error('은행 이름을 입력하세요.');
    const columns = {};
    BANK_KEYS.forEach(k => { if (c[k] !== undefined && header[c[k]]) columns[k] = [header[c[k]]]; });
    if (!columns.balanceCol || !(columns.signedAmountCol || (columns.depositCol && columns.withdrawalCol))) {
      throw new Error('잔액과 입금·출금(또는 거래금액) 열이 지정되어야 사전에 반영할 수 있습니다.');
    }
    columns.memoCols = uniq((c.memoCols || []).map(i => header[i]));
    const excluded = uniq((c.memoCols || []).flatMap(i => c.excludedMemoValues[i] || []));
    dict.bank.names = dict.bank.names || [];
    if (!dict.bank.names.includes(name)) { dict.bank.names = [...dict.bank.names, name].sort((x, y) => x.localeCompare(y, 'ko')); changes.push(`은행 이름 목록에 "${name}" 추가`); }
    const profiles = dict.bank.profiles;
    let profile = profiles.find(p => p.name === name && sameColumns(
      Object.fromEntries(Object.entries(p.columns).filter(([k]) => k !== 'memoCols')),
      Object.fromEntries(Object.entries(columns).filter(([k]) => k !== 'memoCols'))));
    if (!profile) {
      profile = { id: slug(name + profiles.length), name, aliases: [name], columns: {} };
      profiles.push(profile);
      changes.push(`${name}: 새 은행 파일 형식 추가`);
    }
    const before = JSON.stringify([profile.columns, profile.excludedMemoValues || []]);
    profile.columns = columns;
    if (excluded.length) profile.excludedMemoValues = excluded; else delete profile.excludedMemoValues;
    if (JSON.stringify([profile.columns, profile.excludedMemoValues || []]) !== before) {
      changes.push(`${name}: 적요 열 ${columns.memoCols.join(' + ') || '없음'}${excluded.length ? `, 제외 값 ${excluded.length}개` : ''}`);
    }
  }

  // 계좌 1개의 현재 편집 상태를 사전에 반영한다. 반환: 변경 설명 목록 (빈 배열이면 이미 같은 내용)
  // opts.label: 새 형식이 생길 때 붙일 이름(선택)
  function applyAccount(dict, account, opts) {
    if (!account || !account.config || !account.layout || !account.layout.ok) throw new Error('변환이 끝난 계좌만 반영할 수 있습니다.');
    const changes = [];
    if (account.kind === 'bank') applyBank(dict, account, changes);
    else if (account.kind === 'securities') applySecurities(dict, account, changes, opts);
    else throw new Error('은행/증권 유형이 정해지지 않은 계좌입니다.');
    applyAbbreviation(dict, account, changes);
    if (changes.length) {
      dict.version = (dict.version || 0) + 1;
      dict.updatedAt = new Date().toISOString();
      dict.changelog = dict.changelog || [];
      dict.changelog.push({ version: dict.version, at: dict.updatedAt, changes: changes.slice() });
    }
    return changes;
  }

  // 형식 이름 붙이기/바꾸기(빈 글자면 이름 지움). 반환: 변경 설명 목록
  function setProfileLabel(dict, market, brokerId, profileIndex, label) {
    const broker = ((dict.securities || {})[market] || []).find(b => b.id === brokerId);
    const p = broker && broker.profiles && broker.profiles[profileIndex];
    if (!p) throw new Error('이름을 붙일 사전 형식을 찾지 못했습니다.');
    const v = String(label || '').trim().slice(0, 40);
    if ((p.label || '') === v) return [];
    if (v) p.label = v; else delete p.label;
    const changes = [`${MARKET_LABEL[market]} ${broker.name} 형식${profileIndex + 1}: 이름 ${v ? `"${v}"` : '지움'}`];
    dict.version = (dict.version || 0) + 1;
    dict.updatedAt = new Date().toISOString();
    dict.changelog = dict.changelog || [];
    dict.changelog.push({ version: dict.version, at: dict.updatedAt, changes: changes.slice() });
    return changes;
  }

  // 계좌가 국내/해외 중 어느 사전을 쓰는지: 판별 결과 > 파일명 해외/외화 표시 > 국내
  function marketOf(account) {
    if (account.market && MARKETS.includes(account.market)) return account.market;
    if (account.broker && account.broker.market) return account.broker.market;
    return account.foreignHint ? 'foreign' : 'domestic';
  }

  function serialize(dict) {
    const out = clone(dict);
    delete out.accounts; // 계좌별 정보는 저장하지 않음
    out.schemaVersion = SCHEMA_VERSION;
    return '// 통합 통장내역 편집기 기관사전. 화면의 "사전 파일 저장"으로 만든 파일입니다.\n' +
      '// 이 파일을 저장소의 dictionary/institutions.js 와 교체해 올리면 팀 전체에 적용됩니다.\n' +
      '// 아래 { } 안은 순수 JSON 입니다.\n' +
      `(function (root) { root.BANKBOOK_DICTIONARY = ${JSON.stringify(out, null, 2)}; })(typeof window !== 'undefined' ? window : globalThis);\n`;
  }

  // institutions.js(감싼 형식) 또는 순수 JSON 텍스트 → 사전 객체. 형식이 틀리면 오류.
  function parse(textContent) {
    const t = String(textContent || '').replace(/^﻿/, '');
    let jsonText = t.trim();
    if (!jsonText.startsWith('{')) {
      const start = t.indexOf('BANKBOOK_DICTIONARY');
      const open = start >= 0 ? t.indexOf('{', start) : -1;
      const close = t.lastIndexOf('};');
      if (open < 0 || close < open) throw new Error('사전 파일 형식이 아닙니다(BANKBOOK_DICTIONARY 를 찾지 못함).');
      jsonText = t.slice(open, close + 1);
    }
    let d;
    try { d = JSON.parse(jsonText); } catch (e) { throw new Error(`사전 파일의 JSON 문법 오류: ${e.message}`); }
    if (!d || !Array.isArray(d.abbreviations) || !d.bank || !d.securities || !Array.isArray(d.securities.domestic)) {
      throw new Error('사전 파일에 필요한 항목(abbreviations, bank, securities)이 없습니다.');
    }
    const info = prepare(d); // 구조 변환 + 검사(오류면 예외)
    Object.defineProperty(d, 'loadInfo', { value: { notes: info.notes, warnings: info.warnings }, enumerable: false, configurable: true });
    return d;
  }

  // 배포 사전이 깨졌을 때 화면이 멈추지 않도록 쓰는 빈 사전(자동 판별 없음, 열 직접 지정)
  function emptyDictionary() {
    return { schemaVersion: SCHEMA_VERSION, version: 0, abbreviations: [], bank: { columns: {}, profiles: [] },
      securities: { domestic: [], foreign: [], keywordRules: clone(DEFAULT_KEYWORD_RULES) } };
  }

  const api = { applyAccount, setProfileLabel, serialize, parse, migrate, validate, prepare, marketOf, emptyDictionary, SCHEMA_VERSION, DEFAULT_KEYWORD_RULES };
  root.BankbookDictionary = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
