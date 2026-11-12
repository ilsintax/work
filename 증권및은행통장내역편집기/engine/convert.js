// 통합 통장내역 편집기 — 변환 엔진 (판별 결과 → 수지라 업로드 형식)
// 결과 형식: 결제일 / 입금 / 출금 / 잔액 / (참고용)잔액 / 의뢰인/수취인
// 은행형: 통장적요자동생성기 transformFinal 방식 (기초잔액 역산, 적요 열 여러 개 합치기, 제외 값)
// 증권형: 기존 증권 도구 최종본(engine/securities-legacy.js)을 원래 순서대로 호출
(function (root) {
  'use strict';
  const E = root.BankbookEngine || (typeof require !== 'undefined' ? require('./engine.js') : null);
  const OUTPUT_HEADER = ['결제일', '입금', '출금', '잔액', '(참고용)잔액', '의뢰인/수취인'];

  // 레거시 엔진은 엑셀 날짜를 숫자나 글자로 받는다고 가정하므로 Date 셀을 글자로 바꿔 넘긴다.
  function normalizeCell(v) {
    if (v instanceof Date) {
      const d = E.parseDateCell(v);
      if (d) return d;
      return isNaN(v.getTime()) ? '' : v.toTimeString().slice(0, 8);
    }
    return v;
  }

  function orderedRecords(account) {
    const L = account.layout;
    const pairs = L.records.map((rec, k) => ({ rec: rec.map(normalizeCell), src: L.sourceRows[k] }));
    // 날짜가 내림차순이면 뒤집는다 (두 기존 도구와 같은 규칙: 첫 날짜 > 마지막 날짜)
    const first = E.parseDateCell(pairs[0] && pairs[0].rec[L.dateCol]);
    const last = E.parseDateCell(pairs.length && pairs[pairs.length - 1].rec[L.dateCol]);
    if (first && last && first > last) pairs.reverse();
    if (account.config && account.config.orderMode === 'reverse') pairs.reverse();
    // 4단계: 잔액으로 복원한 순서가 채택돼 있으면 그 순서(원본 행 번호 목록)를 따른다. 자동 모드에서만.
    const rep = account.orderRepair;
    if (rep && (!account.config || !account.config.orderMode || account.config.orderMode === 'original') && rep.order.length === pairs.length) {
      const bySrc = new Map(pairs.map(p => [p.src, p]));
      if (rep.order.every(s => bySrc.has(s))) return rep.order.map(s => bySrc.get(s));
    }
    return pairs;
  }

  // ---------------------------------------------------------------- 4단계: 순서 복원(잔액 연결)
  // 위→아래·아래→위가 섞인 파일을 잔액이 이어지도록 다시 줄 세운다. 분류 없이 "|잔액 변화| = 금액(±비용)"만 본다.
  //  1) 날짜순(같은 날은 원래 순서) 2) 날마다 그대로/뒤집기 중 잔액이 더 잘 이어지는 쪽
  //  3) 그래도 끊기는 날은 전날 마지막 잔액에서 출발해 이어지는 행을 하나씩 붙임(그 날 모든 행에 잔액이 있을 때만, 다 이어질 때만 채택)
  // 반환: { order: 원본 행 번호 목록, flippedDays, rechainedDays, moved } 또는 null(바꿀 것 없음)
  function repairOrder(account, pairs, m, unit) {
    const L = account.layout;
    const num = (r, k) => (m[k] === undefined || m[k] === null ? 0 : Math.abs(E.parseAmount(r[m[k]]) || 0));
    const amountOf = r => (account.kind === 'bank'
      ? (m.depositCol !== undefined && m.withdrawalCol !== undefined ? num(r, 'depositCol') + num(r, 'withdrawalCol') : num(r, 'signedAmountCol'))
      : num(r, 'amountCol'));
    const costOf = r => (account.kind === 'bank' ? 0 : ['feeCol', 'taxCol', 'extraCostCol1', 'extraCostCol2', 'extraCostCol3'].reduce((s, k) => s + num(r, k), 0));
    const balOf = r => { const b = E.parseAmount(r[m.balanceCol]); return b === null || (b === 0 && amountOf(r) > 0) ? null : b; };
    const tol = unit * (1 + 1e-6);
    const fit = (p, r) => {
      const b = balOf(r);
      if (p === null || b === null) return 0;
      const dl = Math.abs(b - p), A = amountOf(r), c = costOf(r);
      // 정산금액 열이 있으면 그 값도 잔액 변화로 인정(합산 정산: 첫 행 정산 0 → 잔액 변화 0)
      const S = account.kind !== 'bank' && m.settlementCol !== undefined ? [num(r, 'settlementCol')] : [];
      return [A, A + c, Math.abs(A - c), ...S].some(x => Math.abs(dl - x) <= tol) ? 1 : -1;
    };
    const run = (prev, list) => { let p = prev, sc = 0, breaks = 0; list.forEach(x => { const f = fit(p, x.rec); sc += f; if (f < 0) breaks++; const b = balOf(x.rec); if (b !== null) p = b; }); return { sc, breaks, end: p }; };
    const dateOf = x => E.parseDateCell(x.rec[L.dateCol]) || '';
    const sorted = pairs.map((x, k) => ({ x, k })).sort((a, b) => (dateOf(a.x) < dateOf(b.x) ? -1 : dateOf(a.x) > dateOf(b.x) ? 1 : a.k - b.k)).map(o => o.x);
    const groups = [];
    sorted.forEach(x => { const g = groups[groups.length - 1]; if (g && dateOf(g[0]) === dateOf(x)) g.push(x); else groups.push([x]); });
    let prev = null, flippedDays = 0, rechainedDays = 0;
    const out = [];
    groups.forEach(g => {
      const a = run(prev, g), rv = [...g].reverse(), b = run(prev, rv);
      let best = b.sc > a.sc ? (flippedDays++, rv) : g;
      let r = run(prev, best);
      if (r.breaks && best.length <= 60 && best.every(x => balOf(x.rec) !== null)) {
        // 행 단위로 이어 붙이기: 시작 행 후보를 차례로 시험(전날 잔액이 있으면 그 잔액에서 출발)
        const tryFrom = (start, p0) => {
          const left = best.filter(x => x !== start); const chain = [start]; let p = balOf(start.rec);
          if (p0 !== null && fit(p0, start.rec) < 0) return null;
          while (left.length) {
            const i = left.findIndex(x => fit(p, x.rec) > 0);
            if (i < 0) return null;
            const [n] = left.splice(i, 1); chain.push(n); p = balOf(n.rec);
          }
          return chain;
        };
        let chain = null;
        for (const s of best) { chain = tryFrom(s, prev); if (chain) break; }
        if (chain) { best = chain; rechainedDays++; r = run(prev, best); }
      }
      out.push(...best);
      if (r.end !== null) prev = r.end;
    });
    const moved = out.filter((x, i) => x !== pairs[i]).length;
    if (!moved) return null;
    // 잔액 연결이 끊기는 곳이 원래보다 줄지 않으면 채택하지 않는다(거래 누락 파일에서 엉뚱한 순서를 고르지 않게)
    const breaksBefore = run(null, pairs).breaks, breaksAfter = run(null, out).breaks;
    if (breaksAfter >= breaksBefore) return null;
    return { order: out.map(x => x.src), flippedDays, rechainedDays, moved, days: groups.length, breaksBefore, breaksAfter };
  }

  // ---------------------------------------------------------------- 계좌 기본값
  function defaultTitle(result, sameFileSheetCount) {
    const base = result.file.base || result.fileName.replace(/\.[^.]+$/, '');
    const t = sameFileSheetCount > 1 && result.sheetName ? `${base} ${result.sheetName}` : base;
    return result.splitCount > 1 && result.currency ? `${t} [${result.currency}]` : t; // 통화별로 나눈 계좌
  }

  function createAccount(result, sameFileSheetCount) {
    const account = {
      ...result,
      title: defaultTitle(result, sameFileSheetCount),
      include: result.status !== 'fail',
      detectStatus: result.status,
      config: null,
      output: null,
      edited: false
    };
    // 사전 변경 후 다시 판별할 수 있도록 원본 행을 보관 (목록·직렬화에는 나오지 않게)
    Object.defineProperty(account, 'rows', { value: result.rows, enumerable: false, writable: true });
    if (!result.layout || !result.layout.ok || result.kind === 'unknown') return account;
    account.config = result.kind === 'bank' ? defaultBankConfig(account) : defaultSecuritiesConfig(account);
    return account;
  }

  function defaultBankConfig(account) {
    const c = account.columns || {};
    // 사전 프로필에 적요 열이 저장돼 있으면 그 열들을 모두, 없으면 첫 후보(보낸분/받는분 등)만 기본 선택
    const memoCols = (c.memoCols || []).slice(0, c.profileMemoCount || 1);
    const excluded = (account.bankProfile && account.bankProfile.excludedMemoValues) || [];
    const excludedMemoValues = {};
    if (excluded.length) memoCols.forEach(ci => { excludedMemoValues[ci] = excluded.slice(); });
    return {
      dateCol: account.layout.dateCol,
      depositCol: c.depositCol, withdrawalCol: c.withdrawalCol, balanceCol: c.balanceCol, signedAmountCol: c.signedAmountCol,
      memoCols, excludedMemoValues,
      baseBalance: null, baseBalanceManual: false, orderMode: 'original'
    };
  }

  function defaultSecuritiesConfig(account) {
    return { columnOverrides: {}, baseAmount: null, baseAmountManual: false, orderMode: 'original', classification: null, balanceAdjustment: false };
  }

  // ---------------------------------------------------------------- 은행형 변환
  function convertBank(account) {
    const cfg = account.config;
    const issues = [];
    const pairs = orderedRecords(account);
    const amount = (rec, idx, label, src) => {
      if (idx === undefined || idx === null) return 0;
      const v = E.parseAmount(rec[idx]);
      if (v === null) { issues.push(`${src + 1}행 ${label} "${rec[idx]}"을(를) 숫자로 읽지 못해 0으로 계산했습니다.`); return 0; }
      return v;
    };
    const flows = pairs.map(({ rec, src }) => {
      let dep = 0, wd = 0;
      if (cfg.depositCol !== undefined && cfg.withdrawalCol !== undefined) {
        dep = Math.abs(amount(rec, cfg.depositCol, '입금', src));
        wd = Math.abs(amount(rec, cfg.withdrawalCol, '출금', src));
      } else {
        const a = amount(rec, cfg.signedAmountCol, '거래금액', src);
        if (a > 0) dep = a; else wd = Math.abs(a);
      }
      const balRaw = cfg.balanceCol === undefined ? null : E.parseAmount(rec[cfg.balanceCol]);
      const memoParts = (cfg.memoCols || []).map(ci => {
        const v = E.text(rec[ci]).trim();
        const excluded = (cfg.excludedMemoValues[ci] || []);
        return v && !excluded.includes(v) ? v : null;
      }).filter(Boolean);
      const memo = [...new Set(memoParts)].join(' / ');
      return { date: E.parseDateCell(rec[cfg.dateCol]) || E.text(rec[cfg.dateCol]), dep, wd, balRaw, memo, src };
    });

    let base = cfg.baseBalance;
    if (!cfg.baseBalanceManual || base === null || base === undefined) {
      const firstWithBal = flows.find(f => f.balRaw !== null);
      base = firstWithBal ? firstWithBal.balRaw - firstWithBal.dep + firstWithBal.wd : 0;
      cfg.baseBalance = base;
    }

    let running = base;
    const sheet = [OUTPUT_HEADER.slice()];
    const mismatches = [];
    flows.forEach(f => {
      running += f.dep - f.wd;
      sheet.push([f.date, f.dep, f.wd, running, f.balRaw === null ? '' : f.balRaw, f.memo]);
      if (f.balRaw !== null && Math.abs(running - f.balRaw) >= 1) mismatches.push({ row: sheet.length - 1, src: f.src, diff: f.balRaw - running, date: f.date });
    });
    const lastBal = [...flows].reverse().find(f => f.balRaw !== null);
    // 검증 강도: 기초잔액을 첫 행에서 역산했다면 그 행은 검증에 쓸 수 없다
    const cpCount = flows.filter(f => f.balRaw !== null).length;
    const baseManual = !!cfg.baseBalanceManual;
    const independent = baseManual ? cpCount : Math.max(0, cpCount - 1);
    const reviewReasons = []; // 은행은 입금·출금이 열로 정해져 있어 잔액으로 추정하는 분류가 없다(기준점 수는 참고 정보로만 남김)
    return {
      sheet, removed: null, issues, reviewReasons,
      verification: {
        checkpointCount: cpCount, independentCheckpoints: independent, baseManual,
        isCorrect: mismatches.length === 0,
        baseBalance: base, calculated: running, original: lastBal ? lastBal.balRaw : null,
        diff: lastBal ? lastBal.balRaw - running : 0,
        mismatchCount: mismatches.length, mismatchRows: mismatches.map(m => m.row), firstMismatch: mismatches[0] || null
      }
    };
  }

  // ---------------------------------------------------------------- 증권형 변환
  function legacy() {
    const L = root.LegacySecurities || (typeof require !== 'undefined' ? require('./securities-legacy.js') : null);
    if (!L) throw new Error('증권 변환 엔진(engine/securities-legacy.js)을 불러오지 못했습니다.');
    return L;
  }

  // 기존 엔진에 고정된 키워드 목록을 사전(securities.keywordRules, 증권사별 keywordRules)의 값으로 바꿔 끼운다.
  // 사전에 없는 항목은 기존 엔진의 원래 값을 쓴다. (engine/securities-legacy.js 파일은 고치지 않음)
  let legacyKeywordDefaults = null;
  function applyKeywordRules(L, rules) {
    const DT = L.DataTransformer;
    if (!legacyKeywordDefaults) {
      legacyKeywordDefaults = { removal: DT._phase5RemovalKeywords, block: DT._phase5BlockDepositKeywords, mirror: DT._phase5MirrorPatterns, infer: DT._inferDirectionFromText };
    }
    const r = rules || {};
    DT._phase5RemovalKeywords = r.removalHints || legacyKeywordDefaults.removal;
    DT._phase5BlockDepositKeywords = r.blockDeposit || legacyKeywordDefaults.block;
    DT._phase5MirrorPatterns = r.mirrorPatterns || legacyKeywordDefaults.mirror;
    const dh = r.directionHints;
    DT._inferDirectionFromText = dh ? function (textValue) {
      const n = String(textValue || '').toLowerCase();
      const hit = list => (list || []).some(k => n.includes(String(k).toLowerCase()));
      if (hit(dh.deposit)) return 'deposit';
      if (hit(dh.withdrawal)) return 'withdrawal';
      return null;
    } : legacyKeywordDefaults.infer;
  }

  // 전처리: 기존 엔진은 증권사 id 가 'daishin' 일 때만 행 제외·금액 대체를 했다.
  // 여기서는 증권사 id 와 무관하게 사전 규칙이 있으면 적용한다(다른 증권사에도 같은 규칙을 사전으로 줄 수 있음).
  function preprocessByRules(BAM, data, phase3) {
    const r = phase3.specialRules || {};
    if (r.ignoreTradeTypeContains || r.amountOverrideByTradeType) return BAM.preprocessDaishinData(data, phase3);
    if (r.deriveAmountFromTradeTypeText) return BAM.preprocessAmountFromTradeTypeText(data, phase3);
    return data.map(row => (Array.isArray(row) ? [...row] : row));
  }

  // 통화별 허용오차: 원화 1원, 외화 0.01(센트). 통화를 모르면 원화로 본다.
  const isForeignCcy = account => !!account.currency && account.currency !== 'KRW';
  const toleranceOf = account => (isForeignCcy(account) ? 0.01 : 1);
  const roundTo = (v, unit) => Math.round(v / unit) * unit;

  // 숨은 비용: 정산금액 열이 있으면 "정산금액 ≠ 거래금액 ± 수수료·세금"인 차이를 "기타비용" 열로 만들어 별도 출금 행으로 낸다.
  //  - 매수: 정산 = 금액 + 수수료 + 세금 + 기타 → 기타 = 정산 − (금액+수수료+세금)
  //  - 매도: 정산 = 금액 − 수수료 − 세금 − 기타 → 기타 = (금액−수수료−세금) − 정산
  //  두 식 중 0 이상이면서 더 작은 값을 쓰고, 금액의 1%를 넘는 값은 비용이 아니라고 보고 쓰지 않는다.
  //  같은 날 같은 거래종류가 여러 행이고 일부 행의 정산금액이 0이면(합산 정산) 묶음 합계로 계산해 정산 행에 둔다.
  function addHiddenCostColumn(data, cfg, unit) {
    const slot = ['extraCostCol1', 'extraCostCol2', 'extraCostCol3'].find(k => cfg[k] === undefined);
    if (cfg.settlementCol === undefined || !slot) return { data, details: [], combined: 0 };
    const header = [...data[0], '기타비용'];
    const col = header.length - 1;
    const body = data.slice(1).map(r => { const c = r.slice(); c[col] = ''; return c; });
    const num = (row, k) => (cfg[k] === undefined ? 0 : Math.abs(E.parseAmount(row[cfg[k]]) || 0));
    const extras = row => ['extraCostCol1', 'extraCostCol2', 'extraCostCol3'].reduce((s, k) => s + num(row, k), 0);
    const key = r => `${E.parseDateCell(r[cfg.dateCol]) || E.text(r[cfg.dateCol])}|${E.text(r[cfg.tradeTypeCol]).trim()}`;
    const details = [];
    let combined = 0;
    for (let i = 0; i < body.length;) {
      let j = i;
      while (j + 1 < body.length && key(body[j + 1]) === key(body[i])) j++;
      const g = body.slice(i, j + 1);
      const isCombined = g.length > 1 && g.some(r => num(r, 'settlementCol') === 0 && num(r, 'amountCol') > 0) && g.some(r => num(r, 'settlementCol') > 0);
      if (isCombined) combined++;
      (isCombined ? [g] : g.map(r => [r])).forEach(u => {
        const sum = k => u.reduce((s, r) => s + num(r, k), 0);
        const a = sum('amountCol'), f = sum('feeCol'), t = sum('taxCol'), e = u.reduce((s, r) => s + extras(r), 0), s = sum('settlementCol');
        if (!a || !s) return;
        const cands = [s - (a + f + t + e), (a - f - t - e) - s].filter(x => x > -unit / 2).sort((x, y) => Math.abs(x) - Math.abs(y));
        if (!cands.length) return;
        const v = roundTo(cands[0], unit);
        if (v < unit || v > a * 0.01) return;
        const target = [...u].reverse().find(r => num(r, 'settlementCol') > 0);
        target[col] = Math.round(v * 100) / 100;
        details.push({ key: key(target), amount: target[col] });
      });
      i = j + 1;
    }
    cfg[slot] = col;
    return { data: [header, ...body], details, combined };
  }

  function convertSecurities(account, dictionary) {
    const L = legacy();
    const { DataTransformer: DT, BrokerAutoMapper: BAM, buildAutoClassifyConfig } = L;
    BAM.presets = dictionary.securities.domestic;          // 사전 파일 내용으로 판별
    BAM.foreignPresets = dictionary.securities.foreign;
    const cfg = account.config;
    const issues = [];
    const notes = []; // 안내(상태에 영향 없음)
    const pairs = orderedRecords(account);
    const data = [account.layout.header.slice(), ...pairs.map(p => p.rec)];

    // 기존 도구 3단계: 열 매핑(사전) + 전처리 + 적요 생성("결과 생성" 열)
    // 증권사·형식 결정은 기존 점수 방식(buildPhase3Config) 대신 엔진의 보수적 판별(account.broker) 결과를 쓴다.
    // 처리 방식(분류·특수 규칙·키워드)은 증권사별 사전에서 가져온다(E.brokerProcessing).
    // 아래 필드 구성은 기존 BrokerAutoMapper.buildPhase3Config 반환 형식과 같다.
    const det = account.broker && account.broker.detected ? account.broker : null;
    const proc = E.brokerProcessing(dictionary, account.broker);
    applyKeywordRules(L, proc.keywordRules);
    const phase3 = {
      ...(det ? det.mappedColumns : {}),
      detectedBrokerName: det ? det.name : '', detectedBrokerId: det ? det.brokerId : '', brokerDetected: !!det,
      autoMappedKeys: det ? Object.keys(det.mappedColumns) : [],
      presetRemovalTypes: proc.removalTypes, presetDepositTypes: proc.depositTypes, presetWithdrawalTypes: proc.withdrawalTypes,
      specialRules: proc.specialRules, marketType: proc.market, isComplete: false
    };
    if (phase3.specialRules && phase3.specialRules.deriveAmountFromTradeTypeText && phase3.amountCol === undefined) {
      const existing = E.findHeader(data[0], ['거래금액(KB)', '거래금액(내용)']);
      phase3.amountCol = existing >= 0 ? existing : data[0].length;
      phase3.autoMappedKeys = [...new Set([...phase3.autoMappedKeys, 'amountCol'])];
    }
    // 5단계: 추가 적요 열(사전 형식의 memo 설정 → 사용자 지정이 우선). 기존 엔진 transformPhase3 가 적요 끝에 " / "로 붙인다.
    if (proc.memo) ['additionalCol1', 'additionalCol2'].forEach(k => {
      if (phase3[k] === undefined && Array.isArray(proc.memo[k])) { const i = E.findHeader(data[0], proc.memo[k]); if (i >= 0) phase3[k] = i; }
    });
    Object.assign(phase3, cfg.columnOverrides || {});
    // 잔액 없는 형식: 원본에 잔액 열이 없음(판별 결과·사전 형식 noBalance) 또는 입출금 구분 열만 지정된 경우.
    //  잔액으로 검증·추정할 수 없으므로 입출금 구분 열 → 사전 분류 순으로 나누고, 상태는 항상 확인 필요.
    const noBalance = phase3.balanceCol === undefined && (!!account.noBalance || !!(det && det.profile && det.profile.noBalance) || phase3.directionCol !== undefined);
    const required = noBalance ? ['dateCol', 'tradeTypeCol', 'amountCol'] : ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'];
    const missing = required.filter(k => phase3[k] === undefined);
    if (missing.length) {
      return { sheet: null, error: `증권 필수 열을 찾지 못했습니다: ${missing.map(k => ({ dateCol: '일자', tradeTypeCol: '거래종류', amountCol: '거래금액', balanceCol: '잔액' }[k])).join(', ')}. 열 지정에서 선택하세요.`, mapping: phase3 };
    }
    const unit = toleranceOf(account);
    const hidden = addHiddenCostColumn(data, phase3, unit);
    if (hidden.details.length) notes.push(`정산금액과 차이 나는 숨은 비용 ${hidden.details.length}건(합계 ${roundTo(hidden.details.reduce((s, d) => s + d.amount, 0), unit).toLocaleString('ko-KR')})을 "(기타비용)" 출금 행으로 넣었습니다.`);
    if (hidden.combined) notes.push(`같은 날 여러 건을 한 번에 정산한 거래 ${hidden.combined}묶음 — 그날 마지막 잔액으로 검증합니다.`);
    const pre = preprocessByRules(BAM, preprocessByRules(BAM, hidden.data, phase3), phase3); // 기존 흐름(2단계·3단계)과 같이 두 번
    // 추가 적요에서 뺄 값: 열별 목록(cfg.excludedMemoValues[열 번호]). 처음에는 사전 형식의 memo.excludedValues 로 채운다.
    const ADD_KEYS = ['additionalCol1', 'additionalCol2'];
    if (!cfg.excludedMemoValues) {
      cfg.excludedMemoValues = {};
      const pre0 = (proc.memo && proc.memo.excludedValues) || [];
      if (pre0.length) ADD_KEYS.forEach(k => { if (phase3[k] !== undefined) cfg.excludedMemoValues[phase3[k]] = [...pre0]; });
    }
    ADD_KEYS.forEach(k => {
      const ci = phase3[k];
      const ex = ci === undefined ? [] : (cfg.excludedMemoValues[ci] || []);
      if (ex.length) pre.slice(1).forEach(r => { if (ex.includes(E.text(r[ci]).trim())) r[ci] = ''; });
    });
    const d3 = DT.transformPhase3(pre, phase3);
    // 사전의 앞부분 일치 거래종류("…(…)")를 이 파일의 실제 거래종류로 풀어 쓴다(기존 엔진은 정확히 같은 이름만 비교)
    {
      const actual = [...new Set(d3.slice(1).map(r => DT._normalizeTradeTypeValue(r[phase3.tradeTypeCol])).filter(Boolean))];
      const nf = v => BAM.normalizeTradeType(v);
      ['presetDepositTypes', 'presetWithdrawalTypes', 'presetRemovalTypes'].forEach(k => { phase3[k] = E.expandTypePatterns(phase3[k], actual, nf); });
    }

    const base5 = {
      dateCol: phase3.dateCol, amountCol: phase3.amountCol, balanceCol: phase3.balanceCol, tradeTypeCol: phase3.tradeTypeCol,
      feeCol: phase3.feeCol, taxCol: phase3.taxCol, extraCostCol1: phase3.extraCostCol1, extraCostCol2: phase3.extraCostCol2, extraCostCol3: phase3.extraCostCol3,
      detectedBrokerName: phase3.detectedBrokerName || '', detectedBrokerId: phase3.detectedBrokerId || '',
      presetRemovalTypes: [...(phase3.presetRemovalTypes || [])], presetDepositTypes: [...(phase3.presetDepositTypes || [])], presetWithdrawalTypes: [...(phase3.presetWithdrawalTypes || [])],
      specialRules: phase3.specialRules || null, orderMode: 'original',
      valuesToRemove: [], suggestedRemovals: [], depositTypes: [], withdrawalTypes: [],
      jukyoCol: Math.max(0, d3[0].length - 1), baseAmount: 0
    };

    // 기초잔액 = 첫 기준점(제거 대상이 아닌 행 중 원본 잔액이 있는 행)의 원본 잔액 − 그때까지 계산된 증감
    const firstCheckpointBase = conf => {
      const kept = DT.transformPhase4(d3, conf).kept;
      const proj = DT._buildPhase5Projection(kept, { ...conf, baseAmount: 0, balanceAdjustmentMode: null });
      const cp = proj.checkpoints.find(c => c.isCheckpoint);
      return cp ? cp.originalBalance - cp.calculatedBalance : 0;
    };

    // 행별 잔액 근거: 원본 잔액이 있는 연속 행에서 "잔액 변화 = +금액(−비용)"이면 입금, "−(금액+비용)"이면 출금.
    // 두 방향 중 정확히 하나만 맞을 때만 근거로 센다. 잔액 0 + 금액 있는 행(잔액 표시 없음)은 건너뛴다.
    // 행별 잔액 근거: 원본 잔액이 있는 연속 행에서 "잔액 변화 = +금액(−비용) 또는 +정산금액"이면 입금, 반대면 출금.
    // 두 방향 중 정확히 하나만 맞을 때만 근거로 센다. 잔액 0 + 금액 있는 행(잔액 표시 없음)은 건너뛴다.
    // 직접 입력한 기초잔액이 있으면 첫 행부터 판단한다.
    const nt0 = v => DT._normalizeTradeTypeValue(v);
    const rowDir = new Array(d3.length).fill(null);
    {
      const n = (r, k) => (phase3[k] === undefined ? 0 : Math.abs(E.parseAmount(r[phase3[k]]) || 0));
      let prevBal = cfg.baseAmountManual && cfg.baseAmount !== null && cfg.baseAmount !== undefined ? cfg.baseAmount : null;
      d3.forEach((r, i) => {
        if (!i) return;
        const A = n(r, 'amountCol');
        const raw = E.parseAmount(r[phase3.balanceCol]);
        const c = ['feeCol', 'taxCol', 'extraCostCol1', 'extraCostCol2', 'extraCostCol3'].reduce((s, k) => s + n(r, k), 0);
        // 잔액 0 + 금액 있는 행은 기존 엔진이 "잔액 표시 없음"으로 본다. 여기서는 직전 잔액에서 금액(+비용)을 빼면 정확히 0일 때만
        // 실제 0원(전액 출금)으로 인정한다(그다음 행의 방향 판단이 이어지도록).
        const drained = raw === 0 && A > 0 && prevBal !== null && [A, A + c].some(x => Math.abs(prevBal - x) <= unit * (1 + 1e-6));
        const bal = raw === null || (raw === 0 && A > 0 && !drained) ? null : raw;
        if (A > 0 && bal !== null && prevBal !== null) {
          const S = phase3.settlementCol !== undefined ? n(r, 'settlementCol') : null;
          const dl = bal - prevBal;
          const near = x => Math.abs(dl - x) <= unit * (1 + 1e-6);
          const dep = near(A) || near(A - c) || (S > 0 && near(S)), wd = near(-A) || near(-(A + c)) || (S > 0 && near(-S));
          if (dep !== wd) rowDir[i] = dep ? 'deposit' : 'withdrawal';
        }
        if (bal !== null) prevBal = bal;
      });
      // 합산 정산(정산금액 0인 행): 같은 날·같은 거래종류에서 뒤따라 정산된 행의 방향을 따른다
      if (phase3.settlementCol !== undefined) {
        d3.forEach((r, i) => {
          if (!i || rowDir[i] || !(n(r, 'amountCol') > 0) || n(r, 'settlementCol') !== 0) return;
          const key = x => `${E.text(x[phase3.dateCol])}|${nt0(x[phase3.tradeTypeCol])}`;
          for (let j = i + 1; j < d3.length && key(d3[j]) === key(r); j++) if (rowDir[j]) { rowDir[i] = rowDir[j]; break; }
        });
      }
    }
    // 잔액 없는 형식: 입출금 구분 열(예: 거래구분 = 입금/출금) 값으로 행마다 방향을 정한다
    if (noBalance && phase3.directionCol !== undefined) {
      const dh = (proc.keywordRules && proc.keywordRules.directionHints) || { deposit: ['입금', '매도', '배당', '환급', '이입'], withdrawal: ['출금', '매수', '수수료', '세금', '이체'] };
      const dirOf = v => {
        const s = E.text(v).replace(/\s+/g, '');
        if (!s) return null;
        if (/^입금$|입금$/.test(s) && !/출금/.test(s)) return 'deposit';
        if (/^출금$|출금$/.test(s) && !/입금/.test(s)) return 'withdrawal';
        if ((dh.deposit || []).some(k => s.includes(k))) return 'deposit';
        if ((dh.withdrawal || []).some(k => s.includes(k))) return 'withdrawal';
        return null;
      };
      d3.forEach((r, i) => { if (i) rowDir[i] = dirOf(r[phase3.directionCol]); });
    }
    const tally = () => {
      const ev = {};
      d3.forEach((r, i) => { if (i && rowDir[i]) { const t = nt0(r[phase3.tradeTypeCol]); (ev[t] = ev[t] || { deposit: 0, withdrawal: 0 })[rowDir[i]]++; } });
      return ev;
    };
    let evidence = tally();
    // 예외 규칙(잔액 기준): 같은 거래종류가 입금·출금 양쪽으로 움직이면(예: "입출금") 행마다 나눈다 → "거래종류 [입금]"/"[출금]".
    // 잔액으로 방향을 못 정한 행은 "[방향 미확인]"으로 두고 확인 필요 사유를 남긴다.
    const rowSplit = { types: [], unresolved: 0 };
    if (proc.balanceBasis || noBalance) {
      const mixedTypes = new Set(Object.entries(evidence).filter(([, e]) => e.deposit && e.withdrawal).map(([t]) => t));
      if (mixedTypes.size) {
        const label = { deposit: '[입금]', withdrawal: '[출금]' };
        d3.forEach((r, i) => {
          if (!i) return;
          const t = nt0(r[phase3.tradeTypeCol]);
          if (!mixedTypes.has(t)) return;
          if (!rowDir[i]) rowSplit.unresolved++;
          r[phase3.tradeTypeCol] = `${t} ${rowDir[i] ? label[rowDir[i]] : '[방향 미확인]'}`;
        });
        rowSplit.types = [...mixedTypes];
        evidence = tally();
      }
    }

    let conf;
    if (noBalance && !(cfg.classification && Array.isArray(cfg.classification.depositTypes))) {
      // 사전(형식·공통) 분류를 먼저 적용하고, 입출금 구분 열로 방향이 정해진 거래종류는 그 방향으로
      const ac = buildAutoClassifyConfig(d3, { ...base5, baseAmount: 0 });
      const dep = new Set(ac.depositTypes), wd = new Set(ac.withdrawalTypes), rm = new Set(ac.valuesToRemove);
      Object.entries(evidence).forEach(([t, ev]) => {
        if (ev.deposit && !ev.withdrawal) { dep.add(t); wd.delete(t); rm.delete(t); }
        else if (ev.withdrawal && !ev.deposit) { wd.add(t); dep.delete(t); rm.delete(t); }
      });
      conf = { ...ac, depositTypes: [...dep].sort(), withdrawalTypes: [...wd].sort(), valuesToRemove: [...rm].sort(),
        inferredFeeMode: proc.inferredFeeMode || 'separate' };
      cfg.classification = { valuesToRemove: conf.valuesToRemove, depositTypes: conf.depositTypes, withdrawalTypes: conf.withdrawalTypes, inferredFeeMode: conf.inferredFeeMode };
    } else if (cfg.classification && Array.isArray(cfg.classification.depositTypes)) {
      // 사용자가 편집한 분류를 그대로 사용
      conf = { ...base5, ...cfg.classification };
    } else {
      // 기존 도구의 "잔액 기준 자동 분류" 버튼과 같은 계산. 기초잔액을 모르면 1차 분류로 역산 후 한 번 더 분류.
      const auto = baseAmount => {
        const ac = buildAutoClassifyConfig(d3, { ...base5, baseAmount });
        const { result, error } = DT.autoDistributePhase5(d3, ac);
        if (error) throw new Error(error);
        return { ...ac, valuesToRemove: result.valuesToRemove, suggestedRemovals: result.suggestedRemovals, depositTypes: result.depositTypes, withdrawalTypes: result.withdrawalTypes, inferredFeeMode: result.inferredFeeMode || ac.inferredFeeMode || 'separate' };
      };
      const manualBase = cfg.baseAmountManual && cfg.baseAmount !== null && cfg.baseAmount !== undefined;
      conf = auto(manualBase ? cfg.baseAmount : 0);
      if (!manualBase) {
        const b = firstCheckpointBase(conf);
        if (Math.abs(b) >= 1) conf = auto(b);
      }
      // 잔액 기준 분류(예외 규칙 classificationBasis 'balance'): 행마다 "직전 잔액 → 이 행 잔액" 변화로 방향을 정하고,
      // 거래종류별로 한 방향 근거만 있으면 그 방향으로 확정한다. 근거가 없는 거래종류만 기존 자동 분류 결과를 쓴다.
      if (proc.balanceBasis) {
        const nt = DT._normalizeTradeTypeValue.bind(DT);
        const fixed = [];
        Object.entries(evidence).forEach(([t, ev]) => {
          if (!!ev.deposit === !!ev.withdrawal) return; // 근거 없음 또는 양쪽 → 그대로
          const to = ev.deposit ? 'depositTypes' : 'withdrawalTypes';
          const from = to === 'depositTypes' ? 'withdrawalTypes' : 'depositTypes';
          const was = [...conf[from], ...conf.valuesToRemove].map(nt).includes(t);
          conf[from] = conf[from].filter(x => nt(x) !== t);
          conf.valuesToRemove = conf.valuesToRemove.filter(x => nt(x) !== t);
          if (!conf[to].map(nt).includes(t)) { conf[to] = [...conf[to], t].sort(); }
          if (was) fixed.push(`${t}→${to === 'depositTypes' ? '입금' : '출금'}`);
        });
        if (fixed.length) notes.push(`잔액 기준으로 분류를 바로잡음: ${fixed.join(', ')} (행별 잔액 변화가 한 방향으로만 맞음)`);
      }
      // 수수료 처리 방식: 사용자가 먼저 고른 값 > 사전 프로필에 저장된 값 > 자동 계산 값
      const profileFee = proc.inferredFeeMode;
      const feePref = (cfg.classification && cfg.classification.inferredFeeMode) || profileFee;
      if (feePref) conf.inferredFeeMode = feePref;
      cfg.classification = {
        valuesToRemove: conf.valuesToRemove, depositTypes: conf.depositTypes, withdrawalTypes: conf.withdrawalTypes,
        inferredFeeMode: conf.inferredFeeMode
      };
    }

    if (!cfg.baseAmountManual || cfg.baseAmount === null || cfg.baseAmount === undefined) cfg.baseAmount = firstCheckpointBase(conf);
    conf = { ...conf, baseAmount: cfg.baseAmount };

    // 분류되지 않은 거래종류 알림 (결과에서 빠짐)
    const allTypes = [...new Set(d3.slice(1).map(r => DT._normalizeTradeTypeValue(r[conf.tradeTypeCol])).filter(Boolean))];
    const assigned = new Set([...conf.valuesToRemove, ...conf.depositTypes, ...conf.withdrawalTypes].map(v => DT._normalizeTradeTypeValue(v)));
    const unassigned = allTypes.filter(t => !assigned.has(t));
    if (unassigned.length) issues.push(`분류되지 않아 결과에서 빠진 거래종류: ${unassigned.join(', ')}`);
    // 증권사 공통 분류(다른 형식에서 배운 값)가 이 파일에 쓰였으면 알린다
    const sharedKeys = new Set(proc.sharedAdded.map(t => DT._normalizeTradeTypeValue(t)));
    const sharedApplied = allTypes.filter(t => sharedKeys.has(t));
    if (sharedApplied.length) notes.push(`${proc.entry ? proc.entry.name : ''} 공통 분류 적용(${proc.market === 'foreign' ? '해외' : '국내'}): ${sharedApplied.join(', ')}`);

    // 잔액 보정 행(선택): 날짜별 말 잔액이 원본과 다르면 그 날 끝에 "(잔액보정)" 행을 넣는다(기존 도구의 mismatch_dates 방식)
    conf.balanceAdjustmentMode = cfg.balanceAdjustment ? 'mismatch_dates' : null;
    // 제거(입출금외) 거래에 붙은 비용(수수료·세금·기타비용)은 실제로 빠져나간 돈이다(예: 주식 타사대체출고 수수료 5,000원).
    // 기존 엔진은 제거 행을 통째로 빼므로, 비용만 "(입출금외 거래 비용)" 출금 행으로 바로 뒤에 넣는다.
    // 원래 행의 원본 잔액은 비용 행으로 옮겨(비용이 빠진 뒤 잔액) 그 행에서 잔액을 검증한다. 주식 이동 기록은 입출금외 내역에 그대로.
    const costRowTypes = new Set();
    {
      const rmSet0 = new Set(conf.valuesToRemove.map(t => DT._normalizeTradeTypeValue(t)));
      const costKeys = ['feeCol', 'taxCol', 'extraCostCol1', 'extraCostCol2', 'extraCostCol3'].filter(k => conf[k] !== undefined);
      if (rmSet0.size && costKeys.length) {
        const out = [d3[0]];
        let prevBal = null; // 직전 행까지의 마지막 원본 잔액(잔액 0 = 표시 없음은 제외)
        d3.slice(1).forEach(row => {
          out.push(row);
          const raw = E.parseAmount(row[conf.balanceCol]);
          const hasBal = raw !== null && E.text(row[conf.balanceCol]).trim() !== '' && raw !== 0;
          const before = prevBal;
          if (hasBal) prevBal = raw;
          const t = DT._normalizeTradeTypeValue(row[conf.tradeTypeCol]);
          if (!rmSet0.has(t)) return;
          const cost = costKeys.reduce((sum, k) => sum + Math.abs(E.parseAmount(row[conf[k]]) || 0), 0);
          if (!(cost > 0)) return;
          // 잔액이 정확히 그 비용만큼 줄었을 때만(요약 행의 수수료처럼 이미 다른 행에 반영된 비용은 넣지 않음)
          if (!hasBal || before === null || Math.abs((before - cost) - raw) > unit * (1 + 1e-6)) return;
          const c = row.slice();
          const ct = `${t} (비용)`;
          c[conf.tradeTypeCol] = ct;
          c[conf.amountCol] = cost;
          costKeys.forEach(k => { c[conf[k]] = 0; });
          ['stockQtyCol', 'unitPriceCol'].forEach(k => { if (phase3[k] !== undefined) c[phase3[k]] = ''; });
          c[conf.jukyoCol] = `(입출금외 거래 비용) ${E.text(row[conf.jukyoCol])}`.trim();
          row[conf.balanceCol] = ''; // 원래 행 잔액은 비용 행이 가진다
          out.push(c);
          costRowTypes.add(ct);
        });
        if (costRowTypes.size) {
          d3.splice(0, d3.length, ...out);
          conf = { ...conf, withdrawalTypes: [...conf.withdrawalTypes, ...costRowTypes] };
          notes.push(`입출금외(제거) 거래에 붙은 비용을 출금 행으로 넣었습니다: ${[...costRowTypes].map(t => t.replace(/ \(비용\)$/, '')).join(', ')} (거래 자체는 입출금외 내역에 보관)`);
        }
      }
    }
    const split = DT.transformPhase4(d3, conf);
    const sheet = DT.transformPhase5(split.kept, conf);
    const v = DT.verifyPhase5(split.kept, conf);
    sheet[0] = OUTPUT_HEADER.slice();
    // 외화 계좌: 입금·출금·잔액을 센트 단위로 반올림(누적 계산의 부동소수 오차 제거, 예: 312262.49999999994 → 312262.5)
    if (isForeignCcy(account)) sheet.slice(1).forEach(r => [1, 2, 3].forEach(i => { if (typeof r[i] === 'number') r[i] = Math.round(r[i] * 100) / 100; }));
    const projection = DT._buildPhase5Projection(split.kept, { ...conf, balanceAdjustmentMode: null });

    // 하루 단위 검증: 같은 날 안의 순서·합산 정산 때문에 생기는 중간 차이는 넘기고, 날짜별 마지막 잔액으로 판정한다.
    // 허용오차는 통화별(원화 1원, 외화 0.01).
    const over = c => Math.abs(c.diff) > unit * (1 + 1e-6);
    const dayEnds = DT._summarizePhase5MismatchDates(projection).dailyCheckpoints;
    const dayMismatches = dayEnds.filter(over);
    const badDates = new Set(dayMismatches.map(c => c.date));
    const intraday = projection.checkpoints.filter(c => c.isCheckpoint && over(c) && !badDates.has(c.date));
    const mismatchRows = projection.checkpoints.filter(c => c.isCheckpoint && over(c) && badDates.has(c.date)).map(c => c.lastSheetRowIndex);
    if (intraday.length) notes.push(`같은 날 안에서만 생긴 중간 잔액 차이 ${intraday.length}곳은 그날 마지막 잔액이 맞아 문제로 보지 않았습니다(같은 날 순서·합산 정산).`);
    const lastDay = dayEnds[dayEnds.length - 1];
    const adjustRows = sheet.slice(1).filter(r => String(r[5] || '').startsWith('(잔액보정)')).length;
    const earlyReasons = [];
    let isCorrect = dayMismatches.length === 0 && (!lastDay || !over(lastDay));
    if (adjustRows) {
      isCorrect = true;
      earlyReasons.push(`잔액 보정 행 ${adjustRows}개를 넣었습니다(원본 거래내역에 없는 금액). 누락된 거래가 있는지 확인하세요.`);
    }
    // 원본 잔액이 거래 없이 늘거나 줄어든 날 → 원본에 빠진 거래가 있을 가능성 안내
    let missingSide = '';
    if (!isCorrect && dayMismatches.length) {
      const fm = dayMismatches[0];
      notes.push(`원본 잔액과 계산 잔액이 ${dayMismatches.length}개 날짜에서 다릅니다(처음 ${fm.date}, 차이 ${roundTo(fm.diff, unit).toLocaleString('ko-KR')}). 원본에 빠진 입출금이 있거나 분류가 틀렸을 수 있습니다. 누락이 확실하면 "잔액 보정 행 넣기"로 맞출 수 있습니다.`);
      // 누락 형태: 날짜별 차이가 늘어날 때마다 같은 방향이면(원본 잔액 > 계산 = 입금 누락, 반대 = 출금 누락)
      // 거래내역을 "입금" 또는 "출금"만 골라 조회한 파일일 가능성이 높다.
      const steps = [];
      dayEnds.reduce((prev, c) => { const d = c.diff - prev; if (Math.abs(d) > unit * (1 + 1e-6)) steps.push(d); return c.diff; }, 0);
      const missing = steps.length && steps.every(d => d > 0) ? '입금' : steps.length && steps.every(d => d < 0) ? '출금' : '';
      if (missing) {
        const sum = roundTo(steps.reduce((s, d) => s + Math.abs(d), 0), unit);
        notes.push(`잔액이 맞지 않는 곳이 모두 "${missing}이 빠진" 형태입니다(${steps.length}곳, 합계 ${sum.toLocaleString('ko-KR')}). 거래내역을 ${missing === '입금' ? '출금' : '입금'}만 조회해 내려받았을 수 있습니다 — 조회 조건을 "전체"로 다시 받으면 맞을 가능성이 높습니다.`);
        missingSide = missing;
      }
    }

    // 입출금에 영향을 주지 않는 내역(제거 항목, 미분류, 금액 0)은 버리지 않고 따로 보관한다.
    const depositSet = new Set(conf.depositTypes.map(t => DT._normalizeTradeTypeValue(t)));
    const withdrawalSet = new Set(conf.withdrawalTypes.map(t => DT._normalizeTradeTypeValue(t)));
    const removalSet = new Set(conf.valuesToRemove.map(t => DT._normalizeTradeTypeValue(t)));
    const toNonCash = (row, reason) => {
      const g = k => (phase3[k] === undefined ? '' : row[phase3[k]]);
      // 주식 입고·출고 등은 현금 금액이 없으므로 수량 × 거래단가로 금액을 계산해 적요에 함께 적는다.
      const qtyNum = E.parseAmount(g('stockQtyCol'));
      const priceNum = E.parseAmount(g('unitPriceCol'));
      const valueAmount = qtyNum && priceNum ? Math.round(Math.abs(qtyNum) * Math.abs(priceNum) * 100) / 100 : null;
      const baseMemo = E.text(row[conf.jukyoCol]);
      const memo = valueAmount ? [baseMemo, `${valueAmount.toLocaleString('ko-KR')}원`].filter(Boolean).join(' / ') : baseMemo;
      return {
        date: E.parseDateCell(g('dateCol')) || E.text(g('dateCol')), tradeType: E.text(g('tradeTypeCol')).trim(),
        stockName: E.text(g('stockNameCol')).trim(), qty: g('stockQtyCol'), price: g('unitPriceCol'), valueAmount,
        amount: g('amountCol'), balance: g('balanceCol'), memo, reason
      };
    };
    // ---- 검증 강도와 분류 근거 점검 → "확인 필요" 사유 (거짓 자동 완료 방지)
    const reviewReasons = [...earlyReasons];
    // 제거(입출금외)로 분류된 거래에서 원본 잔액이 계산 잔액과 다르게 움직이면, 실제 돈이 움직인 거래일 수 있다 → 확인 필요
    // (제거 행은 잔액 검증에서 빠지므로 이 점검이 없으면 거짓 자동 완료가 될 수 있음)
    {
      const rmSet1 = new Set(conf.valuesToRemove.map(t => DT._normalizeTradeTypeValue(t)));
      const cps1 = projection.checkpoints || [];
      // 기준 잔액: 직접 입력한 기초잔액, 없으면 입출금 거래 전에는 처음 보이는 잔액(주식 입고·출고만 있는 계좌 대비)
      // 같은 날 안의 순서는 확실하지 않으므로(순서 복원·합산 정산 등), 그날 나온 계산 잔액(그날 시작 잔액 포함) 중 하나와 같으면 정상으로 본다.
      const dayOf = row => E.parseDateCell(row[conf.dateCol]) || E.text(row[conf.dateCol]).slice(0, 10);
      const runsByDay = new Map(); // 날짜 → 그날의 계산 잔액들
      {
        let k = 0, running = cfg.baseAmountManual && cfg.baseAmount !== null && cfg.baseAmount !== undefined ? cfg.baseAmount : null;
        d3.slice(1).forEach(row => {
          const day = dayOf(row);
          if (!runsByDay.has(day)) runsByDay.set(day, running === null ? [] : [running]);
          if (rmSet1.has(DT._normalizeTradeTypeValue(row[conf.tradeTypeCol]))) return;
          if (cps1[k]) { running = cps1[k].calculatedBalance; runsByDay.get(day).push(running); }
          k++;
        });
      }
      const moved = [];
      let firstSeen = null; // 입출금 거래가 하나도 없을 때의 기준(처음 보이는 잔액)
      d3.slice(1).forEach(row => {
        const t = DT._normalizeTradeTypeValue(row[conf.tradeTypeCol]);
        if (!rmSet1.has(t)) return;
        const raw = E.parseAmount(row[conf.balanceCol]);
        const shown = raw !== null && E.text(row[conf.balanceCol]).trim() !== '' && raw !== 0; // 잔액 0 = 표시 없음으로 보고 판단 안 함
        if (!shown) return;
        const runs = runsByDay.get(dayOf(row)) || [];
        const refs = runs.length ? runs : (firstSeen === null ? [raw] : [firstSeen]);
        if (firstSeen === null) firstSeen = raw;
        if (!refs.some(x => Math.abs(raw - x) <= unit * (1 + 1e-6))) {
          moved.push(`${t}(${dayOf(row)}, 원본 잔액 ${Math.round(raw).toLocaleString('ko-KR')} / 그날 계산 잔액 ${refs.map(x => Math.round(x).toLocaleString('ko-KR')).slice(-2).join('·')})`);
        }
      });
      if (moved.length) reviewReasons.push(`입출금외(제거)로 분류한 거래에서 잔액이 바뀌었습니다 — 실제 돈이 움직였다면 입금/출금으로 바꾸세요: ${moved.slice(0, 3).join(', ')}${moved.length > 3 ? ' …' : ''}`);
    }
    if (noBalance) {
      reviewReasons.push('원본에 잔액 열이 없어 잔액으로 검증할 수 없습니다 — 분류(입출금 구분 열·사전 기준)와 기초잔액(전월 말 잔액)을 확인한 뒤 "확인 완료"를 누르세요.');
      const minRow = sheet.slice(1).reduce((m, r) => (typeof r[3] === 'number' && (m === null || r[3] < m[3]) ? r : m), null);
      if (minRow && minRow[3] < -unit) reviewReasons.push(`계산 잔액이 음수가 되는 곳이 있습니다(${minRow[0]}, ${Math.round(minRow[3]).toLocaleString('ko-KR')}) — 기초잔액을 입력하거나 수수료 처리(포함/별도)를 확인하세요.`);
    }
    // 같은 거래종류가 잔액상 입금·출금 양쪽으로 움직이면 거래종류 단위 분류로는 맞출 수 없다 → 확인 필요
    const mixed = Object.entries(evidence).filter(([, ev]) => ev.deposit && ev.withdrawal).map(([t, ev]) => `${t}(입금 ${ev.deposit}·출금 ${ev.withdrawal}건)`);
    if (mixed.length) reviewReasons.push(`같은 거래종류가 잔액상 입금·출금 양쪽으로 움직임: ${mixed.join(', ')}`);
    if (rowSplit.types.length) {
      const cnt = suffix => d3.slice(1).filter(r => String(r[phase3.tradeTypeCol]).endsWith(suffix)).length;
      notes.push(`입금·출금이 섞인 거래종류(${rowSplit.types.join(', ')})를 행마다 잔액으로 나눔: 입금 ${cnt('[입금]')}행 · 출금 ${cnt('[출금]')}행`);
      if (rowSplit.unresolved) reviewReasons.push(`잔액으로 방향을 정하지 못한 행 ${rowSplit.unresolved}개("[방향 미확인]") — 첫 거래라면 기초잔액(전월 말 잔액)을 입력하면 자동으로 정해집니다. 아니면 분류를 직접 확인하세요.`);
    }
    if (proc.balanceBasis) notes.push('예외 규칙: 입금/출금을 사전 목록이 아니라 잔액 기준으로 분류했습니다.');
    const cps = projection.checkpoints;
    const cpCount = cps.filter(c => c.isCheckpoint).length;
    const baseManual = cfg.baseAmountManual && cfg.baseAmount !== null && cfg.baseAmount !== undefined;
    const independent = baseManual ? cpCount : Math.max(0, cpCount - 1);
    const weakNote = `잔액으로 확인할 기준점이 부족합니다(원본 잔액이 있는 행 ${cpCount}개${!baseManual && cpCount ? ', 그중 1개는 기초잔액 계산에 사용' : ''}). 전월 말 잔액을 기초잔액에 입력하면 첫 거래부터 검증합니다.`;
    if (cfg.baseAmount < -0.5) reviewReasons.push(`기초잔액이 음수(${Math.round(cfg.baseAmount).toLocaleString('ko-KR')})입니다. 분류나 기초잔액을 확인하세요.`);
    // 사전(형식별 분류·증권사 공통 분류·키워드 규칙)이 알려주는 거래종류별 분류
    const known = {};
    {
      const kconf = { ...base5 };
      BAM.getPresetClassificationValues(d3, kconf, 'presetDepositTypes').forEach(t => { known[t] = 'deposit'; });
      BAM.getPresetClassificationValues(d3, kconf, 'presetWithdrawalTypes').forEach(t => { known[t] = 'withdrawal'; });
      const kw = BAM.getKeywordClassificationValues(d3, kconf);
      kw.deposit.forEach(t => { known[t] = 'deposit'; }); kw.withdrawal.forEach(t => { known[t] = 'withdrawal'; });
      [...BAM.getPresetRemovalValues(d3, kconf), ...kw.remove].forEach(t => { known[t] = 'remove'; });
    }
    const bucket = t => depositSet.has(t) ? 'deposit' : withdrawalSet.has(t) ? 'withdrawal' : removalSet.has(t) ? 'remove' : 'none';
    // 거래종류별 근거(화면 표시·사전 반영 후보): 현재 분류, 사전 분류, 잔액 근거(행 수), 건수
    //  - 후보: 사전에 없고(또는 사전과 다르고) 잔액 근거가 한 방향뿐이며 현재 분류와 같은 거래종류
    const typeInfo = allTypes.map(t => {
      const rowsOf = d3.slice(1).filter(r => nt0(r[conf.tradeTypeCol]) === t);
      const ev = evidence[t] || { deposit: 0, withdrawal: 0 };
      const cur = bucket(t);
      const evDir = ev.deposit && !ev.withdrawal ? 'deposit' : ev.withdrawal && !ev.deposit ? 'withdrawal' : ev.deposit && ev.withdrawal ? 'mixed' : null;
      return {
        type: t, bucket: cur, dict: known[t] || null, rows: rowsOf.length,
        amountRows: rowsOf.filter(r => Math.abs(DT._toNumber(r[conf.tradeTypeCol === undefined ? 0 : conf.amountCol]) || 0) > 0).length,
        evidence: ev, evidenceDir: evDir,
        conflict: !!evDir && evDir !== 'mixed' && (cur === 'deposit' || cur === 'withdrawal') && evDir !== cur,
        candidate: !proc.balanceBasis && !!evDir && evDir === cur && known[t] !== cur
      };
    });
    if (!cfg.classificationEdited && !noBalance) {
      const L = { deposit: '입금', withdrawal: '출금', remove: '제거', none: '미분류' };
      const firstCpIndex = baseManual ? -1 : cps.findIndex(c => c.isCheckpoint);
      const differs = [], unverified = [], removedGuess = [];
      allTypes.forEach(t => {
        const fb = bucket(t), kb = known[t];
        if (kb) { if (kb !== fb && fb !== 'none') differs.push(`${t}(사전 ${L[kb]}→${L[fb]})`); return; }
        if (fb === 'remove') {
          const hasAmount = d3.slice(1).some(r => DT._normalizeTradeTypeValue(r[conf.tradeTypeCol]) === t && Math.abs(DT._toNumber(r[conf.amountCol]) || 0) > 0);
          if (hasAmount) removedGuess.push(t);
        } else if (fb !== 'none') {
          const verified = cps.some((c, i) => i !== firstCpIndex && c.isCheckpoint && c.tradeType === t);
          if (!verified) unverified.push(t);
        }
      });
      if (differs.length) reviewReasons.push(`사전과 다르게 분류한 거래종류: ${differs.join(', ')}`);
      if (unverified.length) reviewReasons.push(`사전에 없고 잔액으로도 확인되지 않은 거래종류(추정 분류): ${unverified.join(', ')}`);
      if (removedGuess.length) reviewReasons.push(`사전에 없는데 금액이 있는 거래를 제거(입출금외)로 추정: ${removedGuess.join(', ')}`);
      // 잔액으로 추정한 분류가 있는데 확인할 기준점까지 부족하면 맨 앞에 알린다 (사전으로 정해진 분류뿐이면 기준점이 적어도 틀릴 곳이 없음)
      const guessed = differs.length + unverified.length + removedGuess.length + allTypes.filter(x => !known[x] && bucket(x) !== 'none').length;
      if (independent < 1 && guessed) reviewReasons.unshift(weakNote);
    }

    const nonCash = [];
    d3.slice(1).forEach((row, i) => {
      const t = DT._normalizeTradeTypeValue(row[conf.tradeTypeCol]);
      if (removalSet.has(t)) { nonCash.push(toNonCash(row, '제거 항목(입출금 아님)')); return; }
      if (!depositSet.has(t) && !withdrawalSet.has(t)) { nonCash.push(toNonCash(row, '미분류(결과에서 빠짐)')); return; }
      const amt = DT._toNumber(row[conf.amountCol]);
      if (!amt) nonCash.push(toNonCash(row, '금액 0'));
    });

    return {
      sheet, issues, notes, evidence, typeInfo, noBalance, mapping: phase3, tradeTypes: allTypes, unassigned, nonCash, reviewReasons, currency: account.currency || null, tolerance: unit, adjustRows,
      removed: split.removed, classification: cfg.classification,
      verification: {
        checkpointCount: cpCount, independentCheckpoints: independent, baseManual,
        isCorrect, legacyFinalCorrect: !!v.isCorrect, baseBalance: cfg.baseAmount,
        // 기말 잔액 표시는 기존 도구와 같이 "원본 잔액이 있는 마지막 행" 기준. 보정 행을 넣었으면 그 시점 잔액이 원본과 같아진다.
        calculated: (c => (isForeignCcy(account) && typeof c === 'number' ? Math.round(c * 100) / 100 : c))(adjustRows ? v.original : v.calculated), original: v.original, diff: adjustRows ? 0 : (v.diff || 0),
        mismatchCount: dayMismatches.length, intradayDiffCount: intraday.length, mismatchRows, missingSide,
        firstMismatch: dayMismatches[0] ? { date: dayMismatches[0].date, diff: dayMismatches[0].diff } : null
      }
    };
  }

  // ---------------------------------------------------------------- 공통
  function convertAccount(account, dictionary) {
    account.output = null;
    if (!account.config) { account.status = account.detectStatus; account.autoStatus = account.status; if (account.confirmed) account.status = 'confirmed'; return account; }
    try {
      const once = () => (account.kind === 'bank' ? convertBank(account) : convertSecurities(account, dictionary));
      let out = once();
      // 4단계: 잔액이 맞지 않으면 잔액 연결로 순서를 복원해 보고, 불일치가 "확실히 줄 때만" 채택한다(자동 모드에서만).
      const mode = account.config.orderMode || 'original';
      const badness = o => (!o || o.error || !o.verification ? Infinity : (o.verification.isCorrect ? 0 : 1e6) + (o.verification.mismatchCount || 0));
      if (mode === 'original' && !account.orderRepair && out && !out.error && out.verification && !out.verification.isCorrect && (account.kind === 'bank' || out.mapping)) {
        const cfg = account.config;
        const unit = account.kind === 'bank' ? 1 : toleranceOf(account);
        const rep = repairOrder(account, orderedRecords(account), account.kind === 'bank' ? cfg : out.mapping, unit);
        if (rep) {
          const saved = { classification: cfg.classification, baseAmount: cfg.baseAmount };
          if (account.kind !== 'bank') {
            if (!cfg.classificationEdited) cfg.classification = null; // 섞인 순서로 추정한 분류는 버리고 다시
            if (!cfg.baseAmountManual) cfg.baseAmount = null;
          }
          account.orderRepair = rep;
          let tried = null;
          try { tried = once(); } catch (e) { tried = null; }
          if (badness(tried) < badness(out)) {
            const before = out.verification.mismatchCount || 0, after = tried.verification.mismatchCount || 0;
            rep.note = `순서 자동 복원: 날짜순 정렬${rep.flippedDays ? ` · 방향을 바꾼 날 ${rep.flippedDays}일` : ''}${rep.rechainedDays ? ` · 잔액으로 행을 다시 이은 날 ${rep.rechainedDays}일` : ''} · 위치가 바뀐 행 ${rep.moved}개 — 잔액 불일치 ${before}곳 → ${after}곳. 원본 탭의 행 번호는 원래 그대로입니다.`;
            tried.notes = [...(tried.notes || []), rep.note];
            out = tried;
          } else {
            account.orderRepair = null;
            Object.assign(cfg, saved);
            out = once();
          }
        }
      } else if (account.orderRepair && mode === 'original' && out && account.orderRepair.note) {
        out.notes = [...(out.notes || []), account.orderRepair.note];
      }
      account.output = out;
      if (out.error) { account.status = 'review'; }
      else {
        const v = out.verification;
        if (account.detectStatus === 'nocash' || (out.sheet && out.sheet.length <= 1)) {
          // 금액이 있는 거래를 추정으로 전부 제거한 경우는 "현금거래 없음"이 아니라 확인 필요 (거짓 완료 방지)
          const guessed = account.detectStatus !== 'nocash' && (out.reviewReasons || []).length > 0;
          account.status = guessed ? 'review' : 'nocash';
          out.issues = [...(out.issues || []), guessed ? '입출금으로 분류된 거래가 없습니다. 금액이 있는 거래를 모두 제거로 추정했으니 분류를 확인하세요.' : '입출금으로 분류된 거래가 없습니다(입고·출고 등 제거 항목만 있음).'];
        } else if (!v.isCorrect) account.status = 'mismatch';
        else if (account.detectStatus === 'review' || (out.issues || []).length || (out.reviewReasons || []).length) account.status = 'review';
        else account.status = 'ready';
      }
    } catch (e) {
      console.error('[변환 실패]', account.title, e);
      account.output = { sheet: null, error: `변환 중 오류: ${e.message}` };
      account.status = 'fail';
    }
    // 사용자가 "확인 완료"를 누른 계좌: 자동 판정(autoStatus)은 남기고 상태만 '사용자 확인'으로
    account.autoStatus = account.status;
    if (account.confirmed) account.status = 'confirmed';
    return account;
  }

  // 사전이 바뀐 뒤 같은 원본으로 다시 판별·변환한다. 계좌명·저장 여부·id 는 유지.
  // 통화 분리 결과가 여러 계좌면 배열로 돌려준다: 같은 통화(없으면 첫 번째)가 원래 계좌를 이어받고, 나머지는 새 계좌.
  function rebuildAccounts(account, dictionary, sameFileSheetCount) {
    if (!account.rows) return [account];
    const r = E.analyzeSheet({ fileName: account.fileName, sheetName: account.sheetName, rows: account.rows, dictionary, readWarnings: account.readWarnings || [], pick: account.pick || null, merges: account.merges || [] });
    r.format = account.format;
    r.readWarnings = account.readWarnings;
    const parts = E.splitByCurrency(r, dictionary);
    const keep = parts.find(p => account.currency && p.currency === account.currency) || parts[0];
    // 원래 계좌를 이어받는 쪽을 맨 앞에 둔다(rebuildAccount 는 그 계좌를 돌려줌)
    const ordered = [keep, ...parts.filter(p => p !== keep)];
    const baseTitle = String(account.title || '').replace(/ \[[A-Z]{3}\]$/, '');
    return ordered.map(p => {
      Object.defineProperty(p, 'rows', { value: account.rows, enumerable: false, writable: true });
      const next = convertAccount(createAccount(p, sameFileSheetCount), dictionary);
      if (p === keep) {
        next.id = account.id;
        next.include = account.include;
        // 원래 계좌명은 유지하되, 이번에 통화별로 나뉘었으면 통화 표시를 붙인다
        next.title = p.splitCount > 1 && p.currency && !account.splitCount ? `${account.title} [${p.currency}]` : account.title;
      } else {
        next.id = `${account.id}|${p.currency || parts.indexOf(p)}`;
        next.title = p.currency ? `${baseTitle} [${p.currency}]` : next.title;
      }
      return next;
    });
  }
  const rebuildAccount = (account, dictionary, sameFileSheetCount) => rebuildAccounts(account, dictionary, sameFileSheetCount)[0];

  function buildAccounts(results, dictionary) {
    // 계좌명에 시트명을 붙일지는 "시트 수"로 정한다(통화별로 나뉜 계좌는 같은 시트)
    const perFile = {};
    results.forEach(r => { (perFile[r.fileName] = perFile[r.fileName] || new Set()).add(r.sheetName); });
    return results.map(r => convertAccount(createAccount(r, perFile[r.fileName].size), dictionary));
  }

  // ---------------------------------------------------------------- 다운로드용 워크북
  function BankbookEngineNum(v) {
    if (v === null || v === undefined || E.text(v).trim() === '') return null;
    return E.parseAmount(v);
  }

  function safeSheetName(name, used) {
    let base = String(name || '시트').replace(/[\\/?*:[\]]/g, '').trim().slice(0, 31) || '시트';
    let candidate = base, n = 2;
    while (used.has(candidate.toLowerCase())) { const suffix = ` (${n++})`; candidate = base.slice(0, 31 - suffix.length) + suffix; }
    used.add(candidate.toLowerCase());
    return candidate;
  }

  function buildWorkbook(XLSX, accounts, options = {}) {
    const wb = XLSX.utils.book_new();
    const used = new Set();
    const list = accounts.filter(a => a.include && a.output && a.output.sheet);
    const statusLabel = { ready: '자동 완료', review: '확인 필요', mismatch: '잔액 불일치', fail: '실패', nocash: '현금거래 없음', confirmed: '사용자 확인' };
    const summary = [['번호', '계좌명', '유형', '기관', '계좌번호', '거래건수', '기초잔액', '계산 기말잔액', '원본 기말잔액', '차이', '상태', '자동 판정', '확인 사유', '확인 메모', '판별 근거', '원본 파일']];
    list.forEach((a, i) => {
      const v = a.output.verification || {};
      summary.push([i + 1, a.title, a.kind === 'bank' ? '은행' : '증권', a.institution || '', a.account || '', a.output.sheet.length - 1,
        v.baseBalance ?? '', v.calculated ?? '', v.original ?? '', v.diff ?? '', statusLabel[a.status] || a.status, statusLabel[a.autoStatus] || a.autoStatus || '',
        [...(a.reviewReasons || []), ...(a.output.reviewReasons || []), ...(a.output.issues || [])].join(' / '),
        a.confirmed ? a.confirmed.note || '(메모 없음)' : '', a.basis || '', `${a.fileName}${a.sheetName ? ' / ' + a.sheetName : ''}`]);
    });
    const fmt = ws => Object.keys(ws).forEach(k => { if (k[0] !== '!' && ws[k].t === 'n') ws[k].z = '#,##0.########'; });
    const sws = XLSX.utils.aoa_to_sheet(summary); fmt(sws);
    XLSX.utils.book_append_sheet(wb, sws, safeSheetName('요약', used));

    // 입출금에 영향을 주지 않는 내역(주식 입고·출고 등) — 모든 계좌를 한 시트에 모아 보관
    const nonCashRows = [['계좌명', '결제일', '거래종류', '종목명', '수량', '단가', '수량×단가', '거래금액', '원본잔액', '적요', '사유']];
    list.forEach(a => (a.output.nonCash || []).forEach(n => {
      const num = v => { const x = BankbookEngineNum(v); return x === null ? (v ?? '') : x; };
      nonCashRows.push([a.title, n.date, n.tradeType, n.stockName, num(n.qty), num(n.price), n.valueAmount ?? '', num(n.amount), num(n.balance), n.memo, n.reason]);
    }));
    const nws = XLSX.utils.aoa_to_sheet(nonCashRows); fmt(nws);
    XLSX.utils.book_append_sheet(wb, nws, safeSheetName('입출금외내역', used));

    list.filter(a => a.output.sheet.length > 1).forEach(a => { const ws = XLSX.utils.aoa_to_sheet(a.output.sheet); fmt(ws); XLSX.utils.book_append_sheet(wb, ws, safeSheetName(a.title, used)); });
    if (options.includeOriginal) {
      list.forEach(a => {
        const L = a.layout;
        const rows = [['원본 행', ...L.header], ...L.records.map((r, k) => [L.sourceRows[k] + 1, ...r.map(normalizeCell)])];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), safeSheetName(`원본_${a.title}`, used));
      });
    }
    return wb;
  }

  const api = { OUTPUT_HEADER, normalizeCell, createAccount, convertAccount, rebuildAccount, rebuildAccounts, buildAccounts, buildWorkbook, safeSheetName };
  root.BankbookConvert = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
