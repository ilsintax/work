// 통합 통장내역 편집기 — 판별 엔진 (화면과 분리, 브라우저와 Node 테스트에서 함께 사용)
// 담당: 파일 읽기 → 파일명 해석 → 표 구조(제목/헤더/거래 줄 수) 자동 감지 → 1행 1거래로 평탄화
//       → 은행/증권 판별 → 기관 판별 → 상태(자동/확인필요/실패/현금거래없음)
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- 공통 도구
  const text = v => (v === null || v === undefined) ? '' : String(v);
  const norm = v => text(v).replace(/[\s 　]+/g, '');
  const isBlank = v => norm(v) === '';
  const isBlankRow = row => !row || row.every(isBlank);
  const nonEmptyCount = row => (row || []).filter(v => !isBlank(v)).length;

  function isNumericCell(v) {
    if (typeof v === 'number') return Number.isFinite(v);
    const s = text(v).trim();
    return /^[-+]?[\d,]+(\.\d*)?$/.test(s) && /\d/.test(s);
  }

  // 금액 읽기. 읽을 수 없으면 null (0으로 바꾸지 않는다). 빈 칸과 '-'는 0.
  function parseAmount(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const s = text(v).trim();
    if (s === '' || s === '-') return 0;
    const cleaned = s.replace(/,/g, '').replace(/원$/, '');
    if (!/^[-+]?\d+(\.\d+)?$/.test(cleaned)) return null;
    return parseFloat(cleaned);
  }

  const pad2 = n => String(n).padStart(2, '0');
  function validYmd(y, m, d) {
    if (y < 1990 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const dt = new Date(y, m - 1, d);
    if (dt.getMonth() !== m - 1) return null;
    return `${y}-${pad2(m)}-${pad2(d)}`;
  }

  // 날짜 칸이면 'YYYY-MM-DD', 아니면 null. 엑셀 일련번호(숫자)는 금액과 구분할 수 없어 날짜로 보지 않는다.
  function parseDateCell(v) {
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      return validYmd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }
    if (typeof v === 'number') {
      if (Number.isInteger(v) && v >= 19900101 && v <= 21001231) {
        return validYmd(Math.floor(v / 10000), Math.floor(v / 100) % 100, v % 100);
      }
      return null;
    }
    const s = text(v).trim();
    let m = s.match(/^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})(?!\d)/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{4})(\d{2})(\d{2})(?!\d)/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);
    return null;
  }

  // ---------------------------------------------------------------- 파일 읽기
  // .xls 로 저장된 HTML(예: NH투자증권)은 문서에 적힌 문자셋(EUC-KR 등)으로 직접 풀어 읽는다.
  function sniffHtml(bytes) {
    const head = new TextDecoder('latin1').decode(bytes.slice(0, 2048)).replace(/^﻿/, '').trimStart().toLowerCase();
    if (!(head.startsWith('<') || head.startsWith('<!doctype'))) return null;
    const m = head.match(/charset\s*=\s*["']?([a-z0-9_-]+)/);
    return m ? m[1] : 'utf-8';
  }

  // bytes: Uint8Array → { sheets: [{ name, rows }], format, warnings }
  function readWorkbook(XLSX, bytes, fileName) {
    const warnings = [];
    let wb;
    let format = /\.xls$/i.test(fileName) ? 'xls' : (/\.csv$/i.test(fileName) ? 'csv' : 'xlsx');
    const charset = sniffHtml(bytes);
    if (charset) {
      let decoded;
      try { decoded = new TextDecoder(charset).decode(bytes); } catch (e) { decoded = new TextDecoder('euc-kr').decode(bytes); }
      wb = XLSX.read(decoded, { type: 'string', cellDates: true });
      format = 'html';
      warnings.push(`확장자는 엑셀이지만 실제로는 웹페이지(HTML, ${charset}) 파일이라 문자셋을 풀어서 읽었습니다.`);
    } else if (format === 'csv') {
      let decoded = new TextDecoder('utf-8').decode(bytes);
      if (decoded.includes('�')) decoded = new TextDecoder('euc-kr').decode(bytes);
      wb = XLSX.read(decoded, { type: 'string', cellDates: true, raw: true });
    } else {
      wb = XLSX.read(bytes, { type: 'array', cellDates: true });
      if (format === 'xls') warnings.push('구형 .xls 파일입니다. 숫자·날짜를 한 번 더 확인하세요.');
    }
    let sheets = wb.SheetNames.map(name => ({
      name,
      rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true, blankrows: true }),
      merges: ((wb.Sheets[name] && wb.Sheets[name]['!merges']) || []).map(m => ({ r0: m.s.r, c0: m.s.c, r1: m.e.r, c1: m.e.c }))
    }));
    if (format === 'html') sheets = mergeHtmlTables(sheets, warnings);
    return { sheets, format, warnings };
  }

  // 웹페이지(HTML) 파일은 표(<table>)마다 시트가 하나씩 생긴다. 엑셀에서는 한 장으로 보이므로,
  // 거래 행이 없는 표(제목·안내 표)는 따로 계좌로 만들지 않고 바로 다음 거래 표의 제목 줄로 붙인다.
  // 거래 표가 하나도 없으면 원래대로 둔다(판별 단계에서 안내).
  function mergeHtmlTables(sheets, warnings) {
    const nonEmpty = sheets.filter(s => !s.rows.every(isBlankRow));
    const hasData = s => { try { const L = detectLayout(s.rows); return !!(L.ok && L.records.length); } catch (e) { return false; } };
    const flags = nonEmpty.map(hasData);
    if (!flags.some(Boolean) || flags.every(Boolean)) return sheets;
    const out = [];
    let pending = [];
    nonEmpty.forEach((s, i) => {
      if (!flags[i]) { pending.push(...s.rows.filter(r => !isBlankRow(r)), []); return; }
      const shift = pending.length;
      out.push({ name: s.name, rows: [...pending, ...s.rows], merges: (s.merges || []).map(m => ({ r0: m.r0 + shift, c0: m.c0, r1: m.r1 + shift, c1: m.c1 })) });
      pending = [];
    });
    const merged = nonEmpty.length - out.length;
    warnings.push(`웹페이지 안의 제목·안내 표 ${merged}개는 따로 계좌로 만들지 않고 거래 표의 제목으로 붙였습니다.`);
    return out;
  }

  // ---------------------------------------------------------------- 병합 머리글 해석
  // 엑셀에서 머리글 한 칸이 가로로 여러 열에 병합되면 글자는 첫 열에만 있다(예: 한국투자증권 "거래종류" B1:C1 → B=순번, C=거래종류 이름).
  // 병합 범위 안에서 첫 열의 값이 숫자(순번)뿐이고 다른 열에 글자가 있으면, 머리글 이름을 글자 열로 옮기고 첫 열은 "이름(번호)"로 바꾼다.
  // 거래당 1줄·머리글 1줄인 표에만 적용(여러 줄 머리글은 기존 방식 유지). 반환: 바꾼 내용 설명 목록
  function applyHeaderMerges(layout, merges) {
    const notes = [];
    if (!layout || !layout.ok || layout.stride !== 1 || layout.headerRows.length !== 1 || !(merges || []).length) return notes;
    const hr = layout.headerRows[0];
    merges.filter(m => m.r0 <= hr && m.r1 >= hr && m.c1 > m.c0).forEach(m => {
      const name = text(layout.header[m.c0]).trim();
      if (!name) return;
      const vals = c => layout.records.map(r => r[c]).filter(v => v !== null && v !== undefined && text(v).trim() !== '');
      const isNum = v => typeof v === 'number' || /^[\d,.\-\s]+$/.test(text(v).trim());
      const first = vals(m.c0);
      if (first.length && !first.every(isNum)) return; // 첫 열에 글자가 있으면 그대로
      for (let c = m.c0 + 1; c <= m.c1; c++) {
        if (text(layout.header[c]).trim()) continue;
        const v = vals(c);
        if (v.length && v.some(x => !isNum(x))) {
          layout.header[c] = name;
          layout.header[m.c0] = `${name}(번호)`;
          notes.push(`"${name}" 머리글이 ${m.c1 - m.c0 + 1}칸에 병합돼 있어, 첫 칸(순번) 대신 글자가 있는 ${c + 1}번째 열을 "${name}" 열로 읽었습니다.`);
          return;
        }
      }
    });
    return notes;
  }

  // ---------------------------------------------------------------- 파일명 해석
  // 예: "12_KB(123-456-789-01)_거래내역.xlsx", "0_KB국민은행_거래내역_10월.xlsx"
  function parseFileName(fileName) {
    const base = text(fileName).replace(/\.[^.]+$/, '');
    let m = base.match(/^\s*(\d+)\s*_\s*([^(_]+?)\s*\(([^)]*)\)\s*_?(.*)$/);
    if (m) return { seq: +m[1], abbr: m[2].trim(), account: m[3].trim(), docType: m[4].trim(), base };
    m = base.match(/^\s*(\d+)\s*_\s*([^_]+?)\s*_(.*)$/);
    if (m) return { seq: +m[1], abbr: m[2].trim(), account: '', docType: m[3].trim(), base };
    return { seq: null, abbr: '', account: '', docType: '', base };
  }

  function resolveAbbreviation(abbr, dictionary) {
    const key = norm(abbr).toUpperCase();
    if (!key) return null;
    const list = dictionary.abbreviations || [];
    const exact = list.find(e => e.keys.some(k => norm(k).toUpperCase() === key));
    if (exact) return exact;
    // "KB국민은행"처럼 약칭 뒤에 기관명이 붙은 경우: 가장 긴 약칭부터 접두 일치
    const candidates = [];
    list.forEach(e => e.keys.forEach(k => { const nk = norm(k).toUpperCase(); if (nk && key.startsWith(nk)) candidates.push({ e, len: nk.length }); }));
    candidates.sort((a, b) => b.len - a.len);
    if (!candidates.length) return null;
    const hit = candidates[0].e;
    // "KB국민은행"처럼 이름에 은행/증권이 드러나면 그쪽만 남긴다
    if (/은행|뱅크/.test(abbr)) return { ...hit, securities: undefined };
    if (/증권|투자/.test(abbr)) return { ...hit, bank: undefined };
    return hit;
  }

  // 파일명에서 기관 후보 찾기. ① "번호_약칭(계좌)" 형식의 약칭 ② 없으면 "_"·공백으로 나눈 각 단어의 앞부분 약칭
  //   (예: "(고객명)_키움해외_26년상반기" → 키움). 후보가 서로 다른 기관이면 확정하지 않는다(보수적).
  function institutionsFromName(file, dictionary) {
    const structured = resolveAbbreviation(file.abbr, dictionary);
    if (structured) return { hit: structured, source: `파일명 약칭 "${file.abbr}"`, foreignHint: /해외|외화/.test(file.base) };
    const tokens = String(file.base || '').split(/[_\s()[\]-]+/).filter(t => t && !/^\d+$/.test(t));
    const hits = tokens.map(t => ({ t, h: resolveAbbreviation(t, dictionary) })).filter(x => x.h);
    const names = new Set(hits.map(x => `${x.h.bank || ''}|${x.h.securities || ''}`));
    const foreignHint = /해외|외화/.test(file.base);
    if (hits.length && names.size === 1) return { hit: hits[0].h, source: `파일명 단어 "${hits[0].t}"`, foreignHint };
    return { hit: null, source: hits.length > 1 ? `파일명에 기관 후보가 여러 개(${hits.map(x => x.t).join(', ')})` : '', foreignHint };
  }

  // ---------------------------------------------------------------- 표 구조 감지
  // 날짜가 있는 열을 찾고, 날짜 행 간격으로 "거래 1건 = N줄"을 정한 뒤 그 위의 헤더 묶음을 찾는다.
  function detectLayout(rows) {
    const warnings = [];
    const width = rows.reduce((w, r) => Math.max(w, (r || []).length), 0);
    const scanCols = Math.min(width, 15);
    let dateCol = -1, best = 0;
    for (let c = 0; c < scanCols; c++) {
      let count = 0;
      rows.forEach(r => { if (r && parseDateCell(r[c])) count++; });
      if (count > best) { best = count; dateCol = c; }
    }
    if (dateCol < 0) return { ok: false, error: '날짜가 들어 있는 열을 찾지 못했습니다.', warnings };

    const dateRows = [];
    rows.forEach((r, i) => { if (r && parseDateCell(r[dateCol])) dateRows.push(i); });

    // 거래 줄 수(stride): 날짜 행 간격의 최빈값
    let stride = null;
    if (dateRows.length > 1) {
      const freq = {};
      for (let i = 1; i < dateRows.length; i++) { const d = dateRows[i] - dateRows[i - 1]; freq[d] = (freq[d] || 0) + 1; }
      stride = +Object.entries(freq).sort((a, b) => b[1] - a[1] || a[0] - b[0])[0][0];
    }
    const dataStart = dateRows[0];

    // 헤더 묶음: 첫 날짜 행 바로 위(빈 줄은 건너뜀)부터 위로, 날짜·숫자가 없는 줄을 모은다.
    let i = dataStart - 1;
    while (i >= 0 && isBlankRow(rows[i])) i--;
    const block = [];
    while (i >= 0) {
      const r = rows[i];
      if (isBlankRow(r)) break;
      if (r.some(v => parseDateCell(v))) break;
      if (r.filter(isNumericCell).length > 1) break;
      block.unshift(i);
      i--;
    }
    if (!block.length) return { ok: false, error: '날짜 행 위에서 열 제목(헤더)을 찾지 못했습니다.', warnings, dateCol };

    const maxNE = Math.max(...block.map(idx => nonEmptyCount(rows[idx])));
    const threshold = Math.max(2, Math.ceil(maxNE * 0.4));
    const titleRows = [];
    const mainRows = [];      // { index, cells }
    const continuationRows = [];
    block.forEach(idx => {
      const r = rows[idx];
      if (nonEmptyCount(r) >= threshold) {
        mainRows.push({ index: idx, cells: Array.from({ length: width }, (_, c) => text(r[c]).trim()) });
      } else if (!mainRows.length) {
        titleRows.push(idx);
      } else {
        // 줄바꿈된 헤더(예: 신한 "예탁금" + "이용료")는 바로 위 헤더 칸에 이어 붙인다.
        const target = mainRows[mainRows.length - 1].cells;
        r.forEach((v, c) => { if (!isBlank(v)) target[c] = (target[c] || '') + text(v).trim(); });
        continuationRows.push(idx);
      }
    });

    if (stride === null) stride = mainRows.length;
    if (stride !== mainRows.length) {
      warnings.push(`헤더 줄 수(${mainRows.length})와 거래 1건의 줄 수(${stride})가 다릅니다.`);
      if (mainRows.length > stride) {
        mainRows.splice(0, mainRows.length - stride).forEach(m => titleRows.push(m.index));
      } else {
        while (mainRows.length < stride) mainRows.push({ index: -1, cells: Array.from({ length: width }, () => '') });
      }
    }

    const header = [];
    mainRows.forEach(m => m.cells.forEach(c => header.push(c)));

    const records = [];
    const sourceRows = [];
    const covered = new Set();
    dateRows.forEach((r, k) => {
      const next = k + 1 < dateRows.length ? dateRows[k + 1] : rows.length;
      const flat = [];
      for (let j = 0; j < stride; j++) {
        const idx = r + j;
        const row = (idx < next && rows[idx]) ? rows[idx] : [];
        if (idx < next) covered.add(idx);
        for (let c = 0; c < width; c++) flat.push(row[c] === undefined ? null : row[c]);
      }
      records.push(flat);
      sourceRows.push(r);
    });

    const skippedRows = [];
    for (let r = dataStart; r < rows.length; r++) {
      if (!covered.has(r) && !isBlankRow(rows[r])) skippedRows.push(r);
    }
    if (skippedRows.length) warnings.push(`거래로 묶이지 않은 줄 ${skippedRows.length}개가 있습니다(합계·안내 문구 등). 확인이 필요합니다.`);

    return {
      ok: true, warnings, width, dateCol, stride, dataStart,
      titleRows, headerRows: mainRows.map(m => m.index).filter(x => x >= 0), continuationRows,
      header, records, sourceRows, skippedRows
    };
  }

  // 제목 줄에서 "계좌번호 | 123-45" 같은 정보를 읽는다(파일명 계좌번호 확인용).
  function readTitleMeta(rows, titleRows) {
    const meta = {};
    titleRows.forEach(idx => {
      const r = rows[idx] || [];
      for (let c = 0; c < r.length; c++) {
        const k = norm(r[c]);
        if (/^계좌(번호)?$|^계좌정보$/.test(k) && !isBlank(r[c + 1])) meta.account = text(r[c + 1]).trim();
        const m = text(r[c]).match(/(\d[\d-]{6,}\d)/);
        if (!meta.account && m && /계좌|종합/.test(text(r[c]))) meta.account = m[1];
      }
    });
    return meta;
  }

  // ---------------------------------------------------------------- 열 찾기
  // 열 이름 비교용: 공백 제거 + 앞에 붙은 따옴표(' ‘ ’ ` ") 무시 — 엑셀 내보내기마다 따옴표 모양이 다름
  const headerKey = v => norm(v).replace(/^['‘’`"]+/, '');
  function findHeader(header, candidates, exclude) {
    const cands = (candidates || []).map(headerKey).filter(Boolean);
    for (const cand of cands) {
      const idx = header.findIndex((h, i) => headerKey(h) === cand && !(exclude && exclude.has(i)));
      if (idx >= 0) return idx;
    }
    return -1;
  }

  // 은행형 열 찾기: 입금/출금/잔액 열이 따로 있거나(일반), 부호 있는 거래금액 하나(토스형)
  // 은행 사전 프로필: 메모 열을 뺀 모든 열 이름이 헤더에 있으면 일치
  function findBankProfile(header, dictionary) {
    return (dictionary.bank.profiles || []).find(p => Object.entries(p.columns || {})
      .filter(([k]) => k !== 'memoCols').every(([, v]) => findHeader(header, v) >= 0)) || null;
  }

  function mapBankColumns(header, dictionary, layout, profile) {
    const generic = dictionary.bank.columns;
    const pc = (profile && profile.columns) || {};
    const cols = {};
    Object.keys(generic).forEach(k => { cols[k] = [...(pc[k] || []), ...(generic[k] || [])]; });
    const used = new Set();
    const pick = (cands, regex) => {
      let idx = findHeader(header, cands, used);
      if (idx < 0 && regex) idx = header.findIndex((h, i) => !used.has(i) && regex.test(norm(h)));
      if (idx >= 0) used.add(idx);
      return idx;
    };
    const map = {};
    map.dateCol = layout.dateCol;
    used.add(layout.dateCol);
    map.depositCol = pick(cols.depositCol, /^(입금|맡기신|받으신)/);
    map.withdrawalCol = pick(cols.withdrawalCol, /^(출금|찾으신|지급)/);
    map.balanceCol = pick(cols.balanceCol, /잔액|잔고/);
    if (map.depositCol < 0 || map.withdrawalCol < 0) {
      map.signedAmountCol = pick(cols.signedAmountCol, null);
    }
    map.memoCols = [...new Set((cols.memoCols || []).map(c => findHeader(header, [c], used)).filter(i => i >= 0))];
    map.profileMemoCount = (pc.memoCols || []).map(c => findHeader(header, [c])).filter(i => i >= 0).length;
    Object.keys(map).forEach(k => { if (map[k] === -1) delete map[k]; });
    const complete = map.balanceCol !== undefined && ((map.depositCol !== undefined && map.withdrawalCol !== undefined) || map.signedAmountCol !== undefined);
    return { map, complete };
  }

  // 기존 증권 도구 BrokerAutoMapper.detect 와 같은 점수 규칙 (필수 4개 열 2점, 나머지 1점, 이름 일치 2.5점)
  const REQUIRED_SEC_KEYS = ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'];
  function detectBroker(header, sourceName, presets) {
    const source = norm(sourceName);
    const candidates = [];
    (presets || []).forEach(broker => {
      (broker.profiles || []).forEach((profile, profileIndex) => {
        const mappedColumns = {};
        let requiredCount = 0, score = 0;
        Object.entries(profile.columns || {}).forEach(([key, headers]) => {
          const idx = findHeader(header, headers);
          if (idx >= 0) {
            mappedColumns[key] = idx;
            if (REQUIRED_SEC_KEYS.includes(key)) { requiredCount++; score += 2; } else score += 1;
          }
        });
        const aliasMatched = (broker.aliases || []).some(a => source.includes(norm(a)));
        if (aliasMatched) score += 2.5;
        candidates.push({ brokerId: broker.id, name: broker.name, profileIndex, profile, mappedColumns, requiredCount, score, aliasMatched, minMatchedKeys: profile.minMatchedKeys || 5 });
      });
    });
    candidates.sort((a, b) => b.score - a.score || b.requiredCount - a.requiredCount ||
      Object.keys(b.mappedColumns).length - Object.keys(a.mappedColumns).length || a.minMatchedKeys - b.minMatchedKeys);
    const best = candidates[0], second = candidates[1];
    const none = { detected: false };
    if (!best) return none;
    const tied = second && second.score === best.score && second.requiredCount === best.requiredCount &&
      Object.keys(second.mappedColumns).length === Object.keys(best.mappedColumns).length && second.brokerId !== best.brokerId;
    if (best.requiredCount < 3 || Object.keys(best.mappedColumns).length < best.minMatchedKeys || tied) {
      return { ...none, nearest: best.name, tied: !!tied };
    }
    return { detected: true, ...best };
  }

  // 보수적 증권사 판별 (사용자 규칙 16: 열 이름 매칭으로 다른 증권사로 잘못 잡히지 않게)
  //  - 파일명에 증권사가 있으면: 그 증권사의 형식 안에서만 고른다. 다른 증권사로 바꾸지 않는다.
  //      모든 열 일치 → 확정 / 필수 4열 + 최소 열 수 일치 → 확정하되 "일부 열 다름" 확인 필요 / 그 외 → 형식 미확인
  //  - 파일명에 증권사가 없으면: 모든 열이 일치하는 증권사가 딱 한 곳일 때만 확정.
  //  - 반환값 basis: 판별 근거(화면 표시용)
  function evaluateProfiles(header, presets) {
    const rows = [];
    (presets || []).forEach(broker => (broker.profiles || []).forEach((profile, profileIndex) => {
      const mappedColumns = {}; const missing = [];
      let requiredCount = 0;
      Object.entries(profile.columns || {}).forEach(([key, headers]) => {
        const idx = findHeader(header, headers);
        if (idx >= 0) { mappedColumns[key] = idx; if (REQUIRED_SEC_KEYS.includes(key)) requiredCount++; } else missing.push(headers[0]);
      });
      const total = Object.keys(profile.columns || {}).length;
      const matched = Object.keys(mappedColumns).length;
      // 형식 지문(signature): 사전 반영 때 저장한 그 파일의 전체 열 이름. 흔한 열 몇 개만으로 다른 증권사 파일과
      // "완전 일치"가 되지 않도록, 지문이 있으면 그 90% 이상이 있어야 완전 일치로 본다. (기존 사전 형식에는 지문 없음)
      const sig = Array.isArray(profile.signature) ? profile.signature : null;
      const sigRatio = sig && sig.length ? sig.filter(name => findHeader(header, [name]) >= 0).length / sig.length : 1;
      rows.push({ brokerId: broker.id, name: broker.name, profileIndex, profile, mappedColumns, requiredCount, matched, total, missing, sigRatio,
        full: total > 0 && matched === total && sigRatio >= 0.9 && requiredCount === REQUIRED_SEC_KEYS.filter(k => profile.columns[k]).length,
        minMatchedKeys: profile.minMatchedKeys || 5 });
    }));
    return rows.sort((a, b) => (b.full - a.full) || (b.matched / (b.total || 1) - a.matched / (a.total || 1)) || (b.matched - a.matched));
  }

  function decideBroker(header, nameBroker, presets) {
    const all = evaluateProfiles(header, presets);
    const fullOthers = name => [...new Set(all.filter(r => r.full && r.name !== name).map(r => r.name))];
    const label = r => `${r.name} 형식${r.profileIndex + 1}${r.profile && r.profile.label ? ` · ${r.profile.label}` : ''}`;
    if (nameBroker) {
      const own = all.filter(r => r.name === nameBroker);
      const best = own[0];
      const conflict = fullOthers(nameBroker);
      const conflictNote = conflict.length ? `열 구조는 ${conflict.join(', ')} 형식과 완전히 같지만 파일명(${nameBroker})을 따릅니다.` : '';
      if (best && best.full) return { detected: true, ...best, basis: `파일명(${nameBroker}) + ${label(best)}와 모든 열 일치`, conflictNote };
      if (best && best.requiredCount === 4 && best.matched >= Math.min(best.minMatchedKeys, best.total)) {
        return { detected: true, partial: true, ...best, basis: `파일명(${nameBroker}) + ${label(best)}와 ${best.matched}/${best.total}열 일치`, conflictNote };
      }
      return { detected: false, name: nameBroker, nearest: best ? `${label(best)} ${best.matched}/${best.total}열` : '', basis: `파일명(${nameBroker}) — 사전에 맞는 형식 없음`, conflictNote };
    }
    const full = all.filter(r => r.full);
    const brokers = [...new Set(full.map(r => r.name))];
    if (brokers.length === 1) return { detected: true, ...full[0], basis: `파일명에 기관 없음 — ${label(full[0])}와 모든 열 일치(유일)` };
    const top = all[0];
    return {
      detected: false,
      nearest: brokers.length > 1 ? brokers.join(', ') : (top ? `${label(top)} ${top.matched}/${top.total}열` : ''),
      basis: brokers.length > 1 ? `여러 증권사 형식과 모두 일치(${brokers.join(', ')}) — 판별 보류` : '파일명에 기관 없음, 모든 열이 일치하는 형식 없음 — 판별 보류'
    };
  }

  // 국내/해외 사전 선택. 해외(외화) 계좌는 같은 증권사·같은 열 이름이라도 국내 형식을 쓰지 않는다.
  //  - 파일명에 "해외/외화" → 해외 사전만 사용 (맞는 형식이 없으면 확정하지 않음)
  //  - 그 외 → 국내 사전 우선. 국내에서 확정되지 않고 해외 형식과만 모든 열이 일치하면 해외로 본다.
  function chooseMarket(header, secName, dictionary, foreignHint) {
    const sec = (dictionary && dictionary.securities) || {};
    const dom = { ...decideBroker(header, secName, sec.domestic || []), market: 'domestic' };
    const fx = { ...decideBroker(header, secName, sec.foreign || []), market: 'foreign' };
    if (foreignHint) {
      fx.basis = `파일명 해외/외화 → 해외 사전: ${fx.basis}`;
      if (!fx.detected && dom.detected) fx.marketNote = `열 구조는 국내 ${dom.name} 형식과 같지만, 파일명이 해외/외화라서 해외 사전을 씁니다(해외 계좌에 국내 형식을 쓰지 않음).`;
      return fx;
    }
    if (dom.detected) return dom;
    if (fx.detected && !fx.partial) return { ...fx, basis: `해외 사전 — ${fx.basis}` };
    return dom;
  }

  // 판별 결과(broker)가 가리키는 사전의 증권사 항목
  function brokerEntry(dictionary, broker) {
    if (!broker) return null;
    const list = (((dictionary || {}).securities || {})[broker.market || 'domestic']) || [];
    return (broker.brokerId && list.find(b => b.id === broker.brokerId)) || (broker.name && list.find(b => b.name === broker.name)) || null;
  }

  // ---------------------------------------------------------------- 거래종류 앞부분 일치
  // 거래종류 끝 괄호에 거래처(상대 금융기관·회사 이름)가 붙는 증권사가 있다(예: 한화 "[입금] 은행이체입금(국민은행 주식회사○○)").
  // 사전에는 "[입금] 은행이체입금(…)"처럼 괄호 시작까지만 저장하고, 그 앞부분으로 시작하는 거래종류를 같은 분류로 본다.
  // (외화)·(이자)·(연계은행)처럼 짧은 구분어는 그대로 둔다.
  const PATTERN_TAIL = '(…)';
  const isPatternType = t => String(t || '').trim().endsWith(PATTERN_TAIL);
  const patternPrefix = t => String(t || '').trim().slice(0, -PATTERN_TAIL.length + 1); // "…(" 까지
  // 괄호 안이 거래처로 보이는지: 7자 이상, 또는 회사 표시(주식회사·(주)·㈜), 또는 공백이 있고 금융기관 이름이 들어 있음
  function looksLikeCounterparty(inner) {
    const s = String(inner || '').trim();
    if (!s) return false;
    if (/주식회사|㈜|\(주\)|유한회사|합자회사/.test(s)) return true;
    if (s.replace(/\s+/g, '').length >= 7) return true;
    return /\s/.test(s) && /은행|증권|투자|금고|저축|캐피탈|자산운용|보험|신협|농협|수협/.test(s);
  }
  // 거래종류 → 사전에 저장할 모양(거래처 괄호면 앞부분 일치로)
  function generalizeType(t) {
    const s = String(t || '').trim();
    if (isPatternType(s)) return s;
    const m = s.match(/^(.*\()([^()]*)\)\s*$/);
    return m && m[1].length > 1 && looksLikeCounterparty(m[2]) ? `${m[1]}…)` : s;
  }
  // 사전 목록(앞부분 일치 포함)을 실제 거래종류 목록에 맞춰 풀어 쓴다. norm: 비교용 정리 함수
  function expandTypePatterns(list, actualTypes, normFn) {
    const nf = normFn || (v => String(v || '').replace(/\s+/g, ''));
    const out = [];
    (list || []).forEach(t => {
      if (!isPatternType(t)) { out.push(t); return; }
      const pre = nf(patternPrefix(t));
      (actualTypes || []).forEach(a => { if (nf(a).startsWith(pre)) out.push(a); });
    });
    return [...new Set(out)];
  }

  // 예외 규칙(사례에서 나온 규칙): 증권사 항목의 exceptions[]. 그 파일의 열 이름이 사례 파일과 "모두 같을 때만" 적용한다.
  //  { id, label, cases, source, signature: [사례 파일의 전체 열 이름], classificationBasis: 'balance', currencySplit? }
  //  classificationBasis 'balance' = 사전의 입금/출금 목록을 쓰지 않고 잔액으로만 분류(같은 거래종류도 계좌·통화에 따라 입금/출금이 바뀔 수 있음)
  function sameHeaderSet(a, b) {
    const s = list => [...new Set((list || []).map(headerKey).filter(Boolean))].sort().join('|');
    return s(a) === s(b);
  }
  function findException(dictionary, broker, header) {
    const entry = brokerEntry(dictionary, broker);
    return ((entry && entry.exceptions) || []).find(x => Array.isArray(x.signature) && sameHeaderSet(x.signature, header)) || null;
  }

  // 증권사별 처리 방식 = 형식 분류(우선) + 증권사 공통 분류(shared) + 특수 규칙(형식 > 증권사) + 키워드 규칙(증권사 > 공통)
  // 형식이 확정되지 않아도 파일명으로 증권사를 알면 증권사 공통 분류·규칙은 쓴다.
  function brokerProcessing(dictionary, broker) {
    const market = (broker && broker.market) || 'domestic';
    const entry = brokerEntry(dictionary, broker);
    const ex = broker && broker.exceptionRule;
    if (ex && ex.classificationBasis === 'balance') {
      // 예외(잔액 기준 분류): 사전 분류·공통 분류를 쓰지 않는다. 특수 규칙·키워드는 그대로.
      const profile0 = broker.detected ? (broker.profile || {}) : {};
      const common0 = (((dictionary || {}).securities || {}).keywordRules) || null;
      const own0 = entry && entry.keywordRules;
      return { market, entry, depositTypes: [], withdrawalTypes: [], removalTypes: [], specialRules: profile0.specialRules || (entry && entry.specialRules) || null,
        inferredFeeMode: null, sharedAdded: [], keywordRules: common0 || own0 ? { ...(common0 || {}), ...(own0 || {}) } : null, balanceBasis: true, memo: profile0.memo || null };
    }
    const profile = broker && broker.detected ? (broker.profile || {}) : {};
    const specialRules = profile.specialRules || (entry && entry.specialRules) || null;
    const lists = { depositTypes: [...(profile.depositTypes || [])], withdrawalTypes: [...(profile.withdrawalTypes || [])], removalTypes: [...(profile.removalTypes || [])] };
    const sharedAdded = [];
    // 간편 형식(classificationKeywords)은 제거 목록을 "포함" 검색하므로 공통 분류를 섞지 않는다
    // 해외 계좌의 원화 쪽(통화 분리 base)은 외화 기준으로 배운 공통 분류를 섞지 않는다(예: 외화매도 = 외화 출금이지만 원화 입금)
    if (entry && entry.shared && !(specialRules && specialRules.classificationKeywords) && !(broker && broker.currencySide === 'base')) {
      const key = v => norm(v).replace(/\s+/g, '');
      const known = new Set(Object.values(lists).flat().map(key));
      ['removalTypes', 'depositTypes', 'withdrawalTypes'].forEach(b => (entry.shared[b] || []).forEach(t => {
        const k = key(t);
        if (!k || known.has(k)) return;
        known.add(k); lists[b].push(t); sharedAdded.push(t);
      }));
    }
    const common = (((dictionary || {}).securities || {}).keywordRules) || null;
    const own = entry && entry.keywordRules;
    const keywordRules = common || own ? { ...(common || {}), ...(own || {}) } : null;
    return { market, entry, ...lists, specialRules, inferredFeeMode: profile.inferredFeeMode || null, sharedAdded, keywordRules, memo: profile.memo || null };
  }

  const SEC_HINTS = {
    tradeType: /^(거래종류|거래구분|적요명|거래명|거래적요|구분|상세내용|거래유형|적요)$/,
    amount: /거래금액|거래대금|입출금액|세전금액|정산금액/,
    balance: /예수금|최종금액|현금잔액|잔고금액/
  };
  // 잔액 열이 없는 증권 파일(홈페이지·앱 간편 조회 등): 거래종류 + 거래금액 + 종목·수량·수수료 중 하나, 잔액 열 없음
  function looksLikeSecuritiesNoBalance(header) {
    const n = header.map(norm);
    return n.some(h => SEC_HINTS.tradeType.test(h)) && n.some(h => SEC_HINTS.amount.test(h)) &&
      n.some(h => /종목명|거래수량|^수량$|수수료/.test(h)) && !n.some(h => SEC_HINTS.balance.test(h) || /잔액|잔고$/.test(h) && !/유가/.test(h));
  }
  function looksLikeSecurities(header) {
    const n = header.map(norm);
    return n.some(h => SEC_HINTS.tradeType.test(h)) && n.some(h => SEC_HINTS.amount.test(h)) && n.some(h => SEC_HINTS.balance.test(h));
  }

  // ---------------------------------------------------------------- 시트 1개 분석
  // pick: 사용자가 화면에서 고른 기관 { kind: 'securities'|'bank', name, market?: 'domestic'|'foreign' }.
  //       파일명에 그 기관 이름이 있는 것과 똑같이 판별한다(보수적 판별 규칙 그대로).
  function analyzeSheet({ fileName, sheetName, rows, dictionary, readWarnings = [], pick = null, merges = [] }) {
    const result = {
      fileName, sheetName, file: parseFileName(fileName), messages: [...readWarnings],
      status: 'fail', kind: 'unknown', institution: '', account: '', layout: null, columns: null, broker: null
    };
    const fromName = institutionsFromName(result.file, dictionary);
    if (pick && pick.name) {
      fromName.hit = pick.kind === 'bank' ? { bank: pick.name } : { securities: pick.name };
      fromName.source = `사용자가 고른 기관 "${pick.name}${pick.market === 'foreign' ? ' (해외)' : ''}"`;
      if (pick.market) fromName.foreignHint = pick.market === 'foreign';
      result.pick = pick;
    }
    const nameHit = fromName.hit;
    result.nameSource = fromName.source;
    result.foreignHint = fromName.foreignHint;
    result.account = result.file.account;

    if (!rows || !rows.length || rows.every(isBlankRow)) { result.messages.push('빈 시트입니다.'); return result; }
    const layout = detectLayout(rows);
    const mergeNotes = applyHeaderMerges(layout, merges);
    result.merges = merges && merges.length ? merges : undefined; // 다시 판별할 때 같은 병합 정보를 쓰도록 보관
    result.layout = layout;
    if (!layout.ok) { result.messages.push(layout.error); return result; }
    result.messages.push(...layout.warnings, ...mergeNotes);
    if (!layout.records.length) { result.messages.push('거래 행이 없습니다.'); return result; }

    const meta = readTitleMeta(rows, layout.titleRows);
    if (!result.account && meta.account) result.account = meta.account;
    if (result.file.account && meta.account && norm(meta.account).replace(/\D/g, '') !== norm(result.file.account).replace(/\D/g, '')
        && !norm(meta.account).replace(/\D/g, '').startsWith(norm(result.file.account).replace(/\D/g, ''))) {
      result.messages.push(`파일명 계좌번호(${result.file.account})와 파일 안 계좌번호(${meta.account})가 다릅니다.`);
    }

    const header = layout.header;
    const bankProfile = findBankProfile(header, dictionary);
    const bank = mapBankColumns(header, dictionary, layout, bankProfile);
    const hasDeposit = header.some(h => /예수금/.test(norm(h)));
    const secName = nameHit && nameHit.securities ? nameHit.securities : '';
    // 파일명이 은행만 가리키면(예: KB국민은행) 증권사 이름으로 쓰지 않는다
    const broker = chooseMarket(header, secName, dictionary, fromName.foreignHint);
    // 예외 규칙: 열 이름이 사례 파일과 모두 같을 때만
    const exRule = broker.detected ? findException(dictionary, broker, header) : null;
    if (exRule) {
      broker.exceptionRule = exRule;
      broker.basis = `${broker.basis} · 예외 규칙 "${exRule.label || exRule.id}"(사례 ${exRule.cases || 1}건, 열 이름 모두 일치)`;
    }
    if (pick && pick.name && broker.basis) broker.basis = broker.basis.split('파일명(').join('고른 기관(').replace('파일명 해외/외화 → 해외 사전', '고른 기관 해외 → 해외 사전');
    result.broker = broker;
    if (fromName.source && !nameHit) result.messages.push(fromName.source + ' — 기관을 확정하지 않았습니다.');

    if (bank.complete && !hasDeposit && !broker.detected) {
      result.kind = 'bank';
      result.columns = bank.map;
      result.bankProfile = bankProfile;
      const profile = bankProfile;
      result.institution = (nameHit && nameHit.bank) || '';
      result.basis = nameHit && nameHit.bank ? `${fromName.source} → ${nameHit.bank}, 열 구조 은행형` : '열 구조 은행형, 은행 이름 미확인';
      if (!result.institution) result.messages.push(`은행 이름을 파일명에서 찾지 못했습니다${profile ? ` (열 구조는 ${profile.name} 형식과 같음 — 확인 후 입력)` : ''}.`);
      if (nameHit && !nameHit.bank) result.messages.push(`파일명은 증권사(${nameHit.securities})인데 열 구조는 은행형입니다.`);
    } else if (broker.detected || looksLikeSecurities(header) || secName || looksLikeSecuritiesNoBalance(header)) {
      result.kind = 'securities';
      result.basis = broker.basis;
      if (broker.detected) {
        result.columns = broker.mappedColumns;
        result.institution = broker.name;
        if (broker.partial) result.messages.push(`${broker.name} 사전 형식과 일부 열이 다릅니다(없는 열: ${broker.missing.join(', ')}). 열 지정을 확인하세요.`);
      } else {
        result.institution = secName;
        result.messages.push(secName
          ? `${secName}의 사전 형식과 맞지 않습니다${broker.nearest ? ` (가장 가까운 형식: ${broker.nearest})` : ''}. 열을 직접 지정해야 합니다.`
          : `증권사를 확정하지 못했습니다(${broker.basis}${broker.nearest ? `; 가장 가까운 후보: ${broker.nearest}` : ''}). 기관과 열을 직접 지정하세요.`);
      }
      if (broker.conflictNote) result.messages.push(broker.conflictNote);
      if (broker.marketNote) result.messages.push(broker.marketNote);
      if (broker.exceptionRule) result.messages.push(`예외 규칙 적용: ${broker.exceptionRule.label || broker.exceptionRule.id} — 사례 ${broker.exceptionRule.cases || 1}건에서 나온 규칙(열 이름이 사례 파일과 모두 같음)${broker.exceptionRule.classificationBasis === "balance" ? ". 입금/출금은 사전 목록이 아니라 잔액으로 분류합니다" : ""}.`);
      result.market = broker.market;
      if (result.foreignHint) result.messages.push('파일명에 "해외/외화"가 있습니다. 외화 거래가 섞여 있는지 확인하세요.');
      if (nameHit && !nameHit.securities) result.messages.push(`파일명은 은행(${nameHit.bank})인데 열 구조는 증권형입니다.`);
    } else {
      result.messages.push('은행형(입금/출금/잔액)도 증권형(거래종류/거래금액/예수금)도 아닙니다.');
      result.institution = nameHit ? (nameHit.securities || nameHit.bank || '') : '';
      result.status = 'review';
      return result;
    }

    // 잔액 열이 없는 증권 파일: 잔액 검증 없이 입출금 구분 열·사전으로 분류(변환에서 처리)
    if (result.kind === 'securities' && !(result.columns && result.columns.balanceCol !== undefined) && !header.some(h => SEC_HINTS.balance.test(norm(h)))) {
      result.noBalance = true;
      result.messages.push('이 파일에는 잔액(예수금) 열이 없습니다 — 잔액 검증 없이 입출금 구분 열·사전으로 분류합니다. 전월 말 잔액을 기초잔액에 넣으면 잔액이 계산됩니다.');
    }
    if (/잔고증명/.test(result.file.docType)) result.messages.push('파일명이 "잔고증명"입니다. 거래내역인지 확인하세요(자동 제외하지 않음).');

    // 현금 거래 없음: 증권형에서 금액·잔액이 모두 비었거나 0
    if (result.kind === 'securities' && result.columns && result.columns.amountCol !== undefined) {
      const allZero = idx => idx === undefined || layout.records.every(r => { const a = parseAmount(r[idx]); return a === 0; });
      if (allZero(result.columns.amountCol) && allZero(result.columns.balanceCol)) {
        result.status = 'nocash';
        result.messages.push('현금 거래가 없습니다(주식 입고·출고만 있음).');
        return result;
      }
    }

    // 금액을 숫자로 읽을 수 없는 칸 점검 (0으로 바꾸지 않고 알림)
    if (result.columns) {
      const amountKeys = result.kind === 'bank' ? ['depositCol', 'withdrawalCol', 'balanceCol', 'signedAmountCol'] : ['amountCol', 'balanceCol'];
      const bad = [];
      layout.records.forEach((r, k) => amountKeys.forEach(key => {
        const idx = result.columns[key];
        if (idx !== undefined && parseAmount(r[idx]) === null) bad.push(`${layout.sourceRows[k] + 1}행 ${header[idx]}="${text(r[idx])}"`);
      }));
      if (bad.length) result.messages.push(`숫자로 읽지 못한 금액 ${bad.length}칸: ${bad.slice(0, 3).join(', ')}${bad.length > 3 ? ' …' : ''}`);
    }

    // 확인 필요 사유를 따로 모아 화면·요약 시트에 보여준다
    result.reviewReasons = [];
    if (!result.institution) result.reviewReasons.push('기관 미확정');
    if (layout.warnings.length) result.reviewReasons.push('표 구조 경고');
    if (result.kind === 'securities' && !broker.detected) result.reviewReasons.push('사전 형식 없음(열 직접 지정)');
    if (result.kind === 'securities' && broker.partial) result.reviewReasons.push('사전 형식과 일부 열 다름');
    if (broker.conflictNote) result.reviewReasons.push('다른 증권사 형식과도 일치');
    if (result.foreignHint && result.kind === 'securities') result.reviewReasons.push('파일명 해외/외화 표시');
    if (result.kind === 'securities' && result.market === 'foreign') result.reviewReasons.push('해외 계좌(원화·외화 구분 확인)');
    if (result.messages.some(m => /읽지 못한/.test(m))) result.reviewReasons.push('숫자로 못 읽은 금액');
    if (result.noBalance) result.reviewReasons.push('잔액 열 없음(검증 불가)');
    if (result.messages.some(m => /잔고증명/.test(m))) result.reviewReasons.push('잔고증명 파일');
    if (result.messages.some(m => /인데 열 구조는/.test(m))) result.reviewReasons.push('파일명 기관 유형과 열 구조 불일치');
    if (result.messages.some(m => /계좌번호.*다릅니다/.test(m))) result.reviewReasons.push('계좌번호 불일치');
    const needsReview = result.reviewReasons.length > 0;
    result.status = needsReview ? 'review' : 'ready';
    return result;
  }

  // ---------------------------------------------------------------- 통화별 분리 (3단계)
  // 한 시트에 원화·외화 거래가 섞인 해외 계좌를 통화별 계좌로 나눈다. 나누는 방법은 사전에 적힌 대로만 한다(추측하지 않음).
  // 사전 형식(또는 증권사)의 currencySplit:
  //   { currencyCol: ['통화'], baseCurrency: 'KRW',
  //     base:    { columns: { amountCol, balanceCol, settlementCol? ... }, depositTypes, withdrawalTypes, removalTypes },  ← 원화 쪽 열·분류
  //     foreign: { columns: { settlementCol? ... } } }                                                               ← 외화 쪽은 형식 열 + 추가 열
  // 행 배정: 그 통화 행 + (원화 쪽) 원화 금액이 있는 외화 행(환전 등) → 환전 행은 두 계좌 모두에 들어간다.
  const CURRENCY_HEADERS = ['통화', '통화코드', '통화구분', '거래통화', '결제통화'];
  const MONEY_KEYS = ['amountCol', 'balanceCol', 'feeCol', 'taxCol', 'extraCostCol1', 'extraCostCol2', 'extraCostCol3', 'settlementCol', 'kbDepositAmountCol', 'kbWithdrawalAmountCol'];
  const SPLIT_REASONS = ['해외 계좌(원화·외화 구분 확인)', '파일명 해외/외화 표시'];
  function splitByCurrency(result, dictionary) {
    if (!result || result.kind !== 'securities' || !result.layout || !result.layout.ok || !result.broker || !['ready', 'review'].includes(result.status)) return [result];
    const header = result.layout.header;
    const entry = brokerEntry(dictionary, result.broker);
    const ex = result.broker.exceptionRule;
    const spec = (ex && ex.currencySplit) || (result.broker.detected && result.broker.profile && result.broker.profile.currencySplit) || (entry && entry.currencySplit) || null;
    const ccyIdx = findHeader(header, (spec && spec.currencyCol) || CURRENCY_HEADERS);
    if (ccyIdx < 0) return [result];
    const ccyOf = rec => text(rec[ccyIdx]).trim().toUpperCase();
    const codes = [...new Set(result.layout.records.map(ccyOf).filter(Boolean))];
    if (codes.length === 1) { result.currency = codes[0]; return [result]; }
    if (codes.length < 2) return [result];
    const block = msg => {
      result.messages.push(msg);
      if (!result.reviewReasons.includes('통화 혼재(분리 안 됨)')) result.reviewReasons.push('통화 혼재(분리 안 됨)');
      result.status = 'review';
      return [result];
    };
    if (!spec) return block(`통화가 여러 개(${codes.join(', ')}) 섞여 있습니다. 사전에 이 형식의 통화 분리 방법(currencySplit)이 없어 한 계좌로 둡니다 — 통화별 잔액이 섞여 맞지 않습니다.`);
    if (!result.broker.detected) return block(`통화가 여러 개(${codes.join(', ')}) 섞여 있지만 사전 형식이 확정되지 않아 나누지 않았습니다.`);
    const baseCcy = String(spec.baseCurrency || 'KRW').toUpperCase();
    const mapCols = cols => { const m = {}; Object.entries(cols || {}).forEach(([k, names]) => { const i = findHeader(header, names); if (i >= 0) m[k] = i; }); return m; };
    const profileCols = result.broker.mappedColumns || {};
    const baseCols = { ...Object.fromEntries(Object.entries(profileCols).filter(([k]) => !MONEY_KEYS.includes(k))), ...mapCols(spec.base && spec.base.columns) };
    const fxCols = { ...profileCols, ...mapCols(spec.foreign && spec.foreign.columns) };
    const need = ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'];
    const lack = [...need.filter(k => baseCols[k] === undefined).map(k => `${baseCcy} ${k}`), ...need.filter(k => fxCols[k] === undefined).map(k => `외화 ${k}`)];
    if (lack.length) return block(`통화 분리 설정의 열을 이 파일에서 찾지 못했습니다(${lack.join(', ')}).`);
    const nz = (rec, idx) => idx !== undefined && Math.abs(parseAmount(rec[idx]) || 0) > 0;
    const order = [...(codes.includes(baseCcy) ? [baseCcy] : []), ...codes.filter(c => c !== baseCcy)];
    const fxCodes = codes.filter(c => c !== baseCcy);
    const total = result.layout.records.length;
    const parts = order.map(ccy => {
      const isBase = ccy === baseCcy;
      const cols = isBase ? baseCols : fxCols;
      const idx = [];
      result.layout.records.forEach((rec, i) => {
        const c = ccyOf(rec);
        const inBase = c === baseCcy || nz(rec, baseCols.amountCol);
        // 외화가 하나뿐이면 원화 행이라도 외화 금액이 움직이면(외화매수 등) 외화 계좌에도 넣는다
        const inFx = c === ccy || (fxCodes.length === 1 && c === baseCcy && nz(rec, fxCols.amountCol));
        if (isBase ? inBase : inFx) idx.push(i);
      });
      const both = idx.filter(i => { const rec = result.layout.records[i]; return nz(rec, baseCols.amountCol) && nz(rec, fxCols.amountCol); }).length;
      const layout = { ...result.layout, records: idx.map(i => result.layout.records[i]), sourceRows: idx.map(i => result.layout.sourceRows[i]) };
      const sideProfile = isBase
        ? { depositTypes: [...((spec.base || {}).depositTypes || [])], withdrawalTypes: [...((spec.base || {}).withdrawalTypes || [])], removalTypes: [...((spec.base || {}).removalTypes || [])], ...((spec.base || {}).inferredFeeMode ? { inferredFeeMode: spec.base.inferredFeeMode } : {}) }
        : result.broker.profile;
      const part = {
        ...result, layout, columns: cols, currency: ccy, currencySide: isBase ? 'base' : 'foreign', splitCount: order.length,
        messages: [...result.messages, `통화별로 나눔: ${ccy} ${idx.length}건(원본 ${total}건 중)${both ? ` — 원화·외화가 함께 움직인 행(환전 등) ${both}건은 두 계좌 모두에 포함` : ''}.`],
        reviewReasons: result.reviewReasons.filter(r => !SPLIT_REASONS.includes(r)),
        broker: { ...result.broker, mappedColumns: cols, profile: sideProfile, currency: ccy, currencySide: isBase ? 'base' : 'foreign' },
        basis: `${result.basis} · 통화 분리 ${ccy}${isBase ? '(원화 쪽: 형식의 currencySplit.base)' : ''}`
      };
      if (!idx.length) return null;
      part.status = part.reviewReasons.length ? 'review' : 'ready';
      return part;
    }).filter(Boolean);
    return parts.length ? parts : [result];
  }

  function analyzeFile(XLSX, bytes, fileName, dictionary) {
    let book;
    try {
      book = readWorkbook(XLSX, bytes, fileName);
    } catch (e) {
      return [{ fileName, sheetName: '', file: parseFileName(fileName), status: 'fail', kind: 'unknown', institution: '', account: '', messages: [`파일을 읽지 못했습니다: ${e.message}`] }];
    }
    const nonEmpty = book.sheets.filter(s => !s.rows.every(isBlankRow));
    const sheets = nonEmpty.length ? nonEmpty : book.sheets;
    return sheets.flatMap(s => {
      const r = analyzeSheet({ fileName, sheetName: s.name, rows: s.rows, dictionary, readWarnings: book.warnings, merges: s.merges || [] });
      r.format = book.format;
      r.readWarnings = book.warnings;
      // 사전이 바뀌면 다시 판별할 수 있도록 원본 행을 보관(목록 표시·직렬화에는 나오지 않게)
      return splitByCurrency(r, dictionary).map(p => {
        Object.defineProperty(p, 'rows', { value: s.rows, enumerable: false, writable: true });
        return p;
      });
    });
  }

  const api = {
    text, norm, isBlankRow, isNumericCell, parseAmount, parseDateCell,
    readWorkbook, parseFileName, resolveAbbreviation, detectLayout, mapBankColumns, findBankProfile, findHeader, headerKey, detectBroker, evaluateProfiles, decideBroker, chooseMarket, brokerEntry, brokerProcessing, institutionsFromName,
    analyzeSheet, analyzeFile, splitByCurrency, findException, sameHeaderSet,
    isPatternType, patternPrefix, looksLikeCounterparty, generalizeType, expandTypePatterns
  };
  root.BankbookEngine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
