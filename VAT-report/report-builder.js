(function () {
  const QUARTERS = [["1분기", "D"], ["2분기", "F"], ["3분기", "H"], ["4분기", "J"]];
  const MAIN_ROWS = { sales_total: 15, purchase_total: 26, vat_due: 35, labor_total: 37, profit_est: 49 };
  const SALES_BREAKDOWN_ROWS = {
    "총매출": [15],
    "세금계산서": [17],
    "현금영수증/카드매출": [18],
    "영세(수출)": [20],
    "면세매출(합계)": [22],
    "기타": [19, 21, 23, 24, 25],
  };
  const SERIES_COLORS = {
    "총매출": "#8fd3ff",
    "세금계산서": "#4f8cff",
    "현금영수증/카드매출": "#ffbf69",
    "영세(수출)": "#4fd1b5",
    "면세매출(합계)": "#b7a6ff",
    "기타": "#ff8f8f",
  };

  const fileInput = document.getElementById("fileInput");
  const generateBtn = document.getElementById("generateBtn");
  const resetBtn = document.getElementById("resetBtn");
  const statusEl = document.getElementById("status");
  const errorBox = document.getElementById("errorBox");
  const resultPanel = document.getElementById("resultPanel");
  const desktopDownload = document.getElementById("desktopDownload");
  const mobileDownload = document.getElementById("mobileDownload");
  let selectedFile = null;
  let desktopUrl = null;
  let mobileUrl = null;

  const fmtMoney = value => value == null ? "-" : `${Math.round(value).toLocaleString("ko-KR")}원`;
  const fmtPct = value => value == null || Number.isNaN(value) ? "-" : `${(value * 100).toFixed(1)}%`;
  const esc = str => String(str).replace(/[&<>"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
  const num = value => {
    if (value == null || value === "" || value === "-") return null;
    if (typeof value === "number") return value;
    const n = Number(String(value).replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : null;
  };
  const ratio = (a, b) => a == null || b == null || b === 0 ? null : a / b;

  function cell(ws, address) {
    return ws[address] ? ws[address].v : null;
  }

  function rowValue(ws, row, col) {
    return num(cell(ws, `${col}${row}`));
  }

  function rowTotal(ws, row) {
    return rowValue(ws, row, "L") ?? 0;
  }

  function rowQuarters(ws, row) {
    return QUARTERS.map(([label, col]) => ({ label, value: rowValue(ws, row, col) }));
  }

  function sumRows(ws, rows, col) {
    return rows.reduce((sum, row) => sum + (rowValue(ws, row, col) ?? 0), 0);
  }

  function activeQuarters(ws) {
    const items = [];
    for (const [label, col] of QUARTERS) {
      const sales = rowValue(ws, MAIN_ROWS.sales_total, col);
      const purchase = rowValue(ws, MAIN_ROWS.purchase_total, col);
      const active = !((sales == null || sales === 0) && (purchase == null || purchase === 0));
      if (active) items.push({ label, vat_due: rowValue(ws, MAIN_ROWS.vat_due, col) });
    }
    return items;
  }

  function filingPattern(active) {
    const labels = active.map(v => v.label).join("|");
    if (labels === "1분기|2분기|3분기|4분기") return "분기 신고형";
    if (labels === "2분기|4분기") return "반기 신고형";
    if (labels === "4분기") return "연 1회 신고형";
    if (labels === "2분기|3분기|4분기") return "반기 후 분기 신고형";
    return active.length ? "혼합 신고형" : "신고 분기 미확인";
  }

  function displayPeriods(active) {
    const by = Object.fromEntries(active.map(v => [v.label, v]));
    const labels = active.map(v => v.label).join("|");
    if (labels === "1분기|2분기|3분기|4분기") return active;
    if (labels === "2분기") return [{ ...by["2분기"], label: "1~2분기" }];
    if (labels === "2분기|4분기") return [{ ...by["2분기"], label: "1~2분기" }, { ...by["4분기"], label: "3~4분기" }];
    if (labels === "4분기") return [{ ...by["4분기"], label: "1~4분기" }];
    if (labels === "2분기|3분기|4분기") return [{ ...by["2분기"], label: "1~2분기" }, { ...by["3분기"], label: "3분기" }, { ...by["4분기"], label: "4분기" }];
    return active;
  }

  function elapsedMonthsFromLabel(label) {
    if (label === "1분기") return 3;
    if (label === "1~2분기" || label === "2분기") return 6;
    if (label === "3분기") return 9;
    if (label === "3~4분기" || label === "4분기" || label === "1~4분기") return 12;
    return null;
  }

  function compressMetric(values, periods) {
    const map = Object.fromEntries(values.map(v => [v.label, v.value]));
    return periods.map(period => {
      const key = period.label === "1~2분기" ? "2분기" : period.label === "3~4분기" || period.label === "1~4분기" ? "4분기" : period.label;
      return { label: period.label, value: map[key] ?? null };
    });
  }

  function buildContext(workbook, filename) {
    const years = workbook.SheetNames.filter(name => /^\d+$/.test(name)).sort((a, b) => Number(a) - Number(b));
    if (!years.length) throw new Error("연도 시트를 찾지 못했습니다.");
    const latestYear = years[years.length - 1];
    const previousYear = years.length > 1 ? years[years.length - 2] : null;
    const ws = workbook.Sheets[latestYear];
    const active = activeQuarters(ws);
    const periods = displayPeriods(active);
    const latestSales = rowTotal(ws, MAIN_ROWS.sales_total);
    const latestPurchase = rowTotal(ws, MAIN_ROWS.purchase_total);
    const latestLabor = rowTotal(ws, MAIN_ROWS.labor_total);
    const annualDatasets = {};
    Object.keys(SALES_BREAKDOWN_ROWS).forEach(name => {
      annualDatasets[name] = [];
    });
    years.forEach(year => {
      const yearWs = workbook.Sheets[year];
      Object.entries(SALES_BREAKDOWN_ROWS).forEach(([name, rows]) => {
        annualDatasets[name].push(sumRows(yearWs, rows, "L"));
      });
    });
    const currentPeriodLabel = periods.length ? periods[periods.length - 1].label : null;
    const elapsedMonths = currentPeriodLabel ? elapsedMonthsFromLabel(currentPeriodLabel) : null;
    const projectedAnnualSales = elapsedMonths && elapsedMonths < 12 ? (latestSales / elapsedMonths) * 12 : latestSales;
    const projectedAnnualSeries = annualDatasets["총매출"].slice();
    if (projectedAnnualSeries.length) projectedAnnualSeries[projectedAnnualSeries.length - 1] = projectedAnnualSales;
    let previousQuarterMetrics = null;
    if (previousYear) {
      const prevWs = workbook.Sheets[previousYear];
      const prevActive = activeQuarters(prevWs);
      const prevPeriods = displayPeriods(prevActive);
      previousQuarterMetrics = {
        year: previousYear,
        periods: prevPeriods,
        metrics: {
          "매출": compressMetric(rowQuarters(prevWs, MAIN_ROWS.sales_total), prevPeriods),
          "매입": compressMetric(rowQuarters(prevWs, MAIN_ROWS.purchase_total), prevPeriods),
          "부가세 납부세액": compressMetric(rowQuarters(prevWs, MAIN_ROWS.vat_due), prevPeriods),
          "인건비": compressMetric(rowQuarters(prevWs, MAIN_ROWS.labor_total), prevPeriods),
          "예상손익": compressMetric(rowQuarters(prevWs, MAIN_ROWS.profit_est), prevPeriods),
        },
      };
    }

    return {
      company: cell(ws, "E5") || filename.replace(/\.[^.]+$/, ""),
      latest_year: latestYear,
      filing_pattern: filingPattern(active),
      active_periods: periods,
      current_filing: periods[periods.length - 1] || null,
      latest_metrics: {
        sales: latestSales,
        purchase: latestPurchase,
        labor: latestLabor,
        profit: rowTotal(ws, MAIN_ROWS.profit_est),
        vat_total: rowTotal(ws, MAIN_ROWS.vat_due),
        purchase_ratio: ratio(latestPurchase, latestSales),
        labor_ratio: ratio(latestLabor, latestSales),
        previous_actual_year: String(Number(latestYear) - 1),
        previous_actual_profit: num(cell(ws, "H9")),
        elapsed_months: elapsedMonths,
        projected_annual_sales: projectedAnnualSales,
      },
      quarter_metrics: {
        "매출": compressMetric(rowQuarters(ws, MAIN_ROWS.sales_total), periods),
        "매입": compressMetric(rowQuarters(ws, MAIN_ROWS.purchase_total), periods),
        "부가세 납부세액": compressMetric(rowQuarters(ws, MAIN_ROWS.vat_due), periods),
        "인건비": compressMetric(rowQuarters(ws, MAIN_ROWS.labor_total), periods),
        "예상손익": compressMetric(rowQuarters(ws, MAIN_ROWS.profit_est), periods),
      },
      previous_quarter_metrics: previousQuarterMetrics,
      sales_series: {
        annual: { labels: years, datasets: annualDatasets },
        projected_total_sales: projectedAnnualSeries,
      },
    };
  }

  function makeReport(ctx, mode) {
    const mobile = mode === "mobile";
    const data = JSON.stringify(ctx).replace(/</g, "\\u003c");
    const previousHeaders = ctx.previous_quarter_metrics ? ctx.previous_quarter_metrics.periods.map(v => v.label) : [];
    const sharedHeaders = [...ctx.active_periods.map(v => v.label)];
    previousHeaders.forEach(label => {
      if (!sharedHeaders.includes(label)) sharedHeaders.push(label);
    });
    const normalizeForHeaders = (metrics, headers) => {
      const out = {};
      Object.entries(metrics).forEach(([label, values]) => {
        const map = Object.fromEntries(values.map(item => [item.label, item.value]));
        out[label] = headers.map(header => ({ label: header, value: map[header] ?? null }));
      });
      return out;
    };
    const latestDesktopMetrics = normalizeForHeaders(ctx.quarter_metrics, sharedHeaders);
    const previousDesktopMetrics = ctx.previous_quarter_metrics ? normalizeForHeaders(ctx.previous_quarter_metrics.metrics, sharedHeaders) : null;
    const previousSection = ctx.previous_quarter_metrics ? (mobile
      ? `<section class="panel"><h2>${ctx.previous_quarter_metrics.year}년 분기별 수치</h2>${Object.entries(ctx.previous_quarter_metrics.metrics).map(([label, values]) => { const total = values.reduce((sum, item) => sum + (item.value ?? 0), 0); return `<div class="q-card"><div style="font-weight:700;margin-bottom:10px;">${label}</div><div class="q-grid">${values.map(item => `<div class="q-item"><span>${item.label}</span><strong>${fmtMoney(item.value)}</strong></div>`).join("")}<div class="q-item total"><span>합계</span><strong>${fmtMoney(total)}</strong></div></div></div>`; }).join("")}</section>`
      : `<section class="panel"><h2>${ctx.previous_quarter_metrics.year}년 분기별 수치</h2><table><thead><tr><th>항목</th>${sharedHeaders.map(v => `<th>${v}</th>`).join("")}<th>합계</th></tr></thead><tbody>${Object.entries(previousDesktopMetrics).map(([label, values]) => { const total = values.reduce((sum, item) => sum + (item.value ?? 0), 0); return `<tr><th>${label}</th>${values.map(item => `<td>${fmtMoney(item.value)}</td>`).join("")}<td>${fmtMoney(total)}</td></tr>`; }).join("")}</tbody></table><div class="note">최신연도 수치와 비교할 수 있도록 전년도 신고단위 기준 수치를 참고용으로 함께 표시했습니다.</div></section>`)
      : "";
    return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(ctx.company)} ${ctx.latest_year}</title><style>
      :root{--bg:#0d1825;--card:rgba(20,34,50,.96);--card2:rgba(16,29,43,.98);--line:rgba(153,184,216,.14);--line2:rgba(153,184,216,.26);--text:#e7eff8;--muted:#9eb0c4;--accent:#8bbce8;--shadow:0 14px 34px rgba(0,0,0,.22)}
      body.theme-white{--bg:#fff;--card:#fff;--card2:#fff;--line:rgba(34,58,84,.12);--line2:rgba(34,58,84,.2);--text:#22384b;--muted:#647789;--accent:#365c84;--shadow:0 10px 24px rgba(26,46,68,.08)}
      *{box-sizing:border-box}body{margin:0;font-family:"Malgun Gothic","Apple SD Gothic Neo",sans-serif;background:radial-gradient(circle at top right,rgba(184,166,132,.14),transparent 28%),linear-gradient(180deg,#dfe4e8 0%,#cfd6dc 100%);color:var(--text)}body.theme-white{background:#fff}
      .wrap{max-width:${mobile ? "480px" : "1240px"};margin:0 auto;padding:${mobile ? "16px 14px 32px" : "28px 18px 56px"}}.topbar{display:flex;justify-content:flex-end;margin-bottom:12px}.theme-switch{display:inline-flex;gap:6px;border:1px solid var(--line);border-radius:999px;padding:6px;background:rgba(18,31,46,.92)}body.theme-white .theme-switch{background:#f5f8fb}.theme-btn{border:0;background:transparent;color:var(--muted);padding:8px 12px;border-radius:999px;font:inherit;cursor:pointer}.theme-btn.active{background:rgba(139,188,232,.14);color:var(--text)}
      .hero,.panel{background:linear-gradient(180deg,var(--card) 0%,var(--card2) 100%);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow)}.hero{padding:${mobile ? "20px" : "30px"}}.panel{margin-top:14px;padding:${mobile ? "18px" : "24px"}}.eyebrow{color:var(--accent);font-size:11px;letter-spacing:.12em;text-transform:uppercase;font-weight:700}h1{margin:10px 0 12px;font-size:${mobile ? "26px" : "32px"};line-height:1.25}.sub,.note{color:var(--muted);line-height:1.8}.panel h2{margin:0 0 12px;font-size:${mobile ? "18px" : "22px"}}ul{margin:0;padding-left:18px;line-height:1.8}
      .kpis{display:grid;grid-template-columns:${mobile ? "1fr" : "repeat(4,minmax(0,1fr))"};gap:12px;margin-top:14px}.kpi{background:rgba(25,42,60,.92);border:1px solid var(--line);border-radius:8px;padding:16px}.label{color:var(--muted);font-size:11px;text-transform:uppercase}.value{margin-top:8px;font-size:${mobile ? "24px" : "27px"};font-weight:700}
      .toggle-wrap,.legend{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}.toggle,.legend-item{display:inline-flex;align-items:center;gap:8px;padding:8px 10px;border-radius:4px;border:1px solid var(--line);background:rgba(19,34,50,.88);color:var(--muted);font-size:12px}.toggle.active{background:rgba(139,188,232,.12);color:var(--text)}.toggle input{display:none}.dot,.legend-color{width:10px;height:10px;border-radius:50%}
      .chart-frame{border:1px solid var(--line);background:rgba(18,31,46,.92);border-radius:8px;padding:${mobile ? "16px 12px 14px 56px" : "22px 22px 20px 80px"}}.chart-title{font-size:${mobile ? "18px" : "22px"};font-weight:700}.chart-sub{color:var(--muted);font-size:12px;margin-top:4px}.plot{position:relative;height:${mobile ? "220px" : "320px"};margin-top:12px}.y-axis{position:absolute;left:${mobile ? "-52px" : "-78px"};top:0;bottom:0;width:${mobile ? "48px" : "72px"}}.y-tick{position:absolute;left:0;width:${mobile ? "48px" : "72px"};transform:translateY(-50%);text-align:right;color:var(--muted);font-size:${mobile ? "10px" : "12px"}}.grid-line{position:absolute;left:0;right:0;height:1px;background:rgba(153,184,216,.12)}svg{width:100%;height:${mobile ? "220px" : "320px"};display:block}.x-axis{display:grid;margin-top:8px;color:var(--muted);font-size:${mobile ? "11px" : "13px"};text-align:center}.x-axis div{white-space:pre-line}
      .tooltip{position:fixed;z-index:30;pointer-events:none;padding:8px 10px;border-radius:6px;border:1px solid var(--line2);background:rgba(12,22,34,.98);color:var(--text);font-size:12px;line-height:1.5;opacity:0}.tooltip strong{display:block;color:var(--accent);font-size:11px}
      table{width:100%;border-collapse:collapse;font-size:14px;table-layout:fixed}thead th{background:rgba(139,188,232,.06);color:var(--muted);font-size:12px;text-transform:uppercase}th,td{padding:12px 10px;border-bottom:1px solid rgba(153,184,216,.14);text-align:right}th:first-child,td:first-child{text-align:left;width:160px}.summary-table th:first-child{width:${mobile ? "160px" : "280px"};white-space:${mobile ? "normal" : "nowrap"}}
      .q-card{border:1px solid var(--line);border-radius:8px;padding:14px;background:rgba(18,31,46,.88);margin-top:10px}.q-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.q-item{border:1px solid var(--line);border-radius:6px;padding:10px;background:rgba(139,188,232,.08)}.q-item span{display:block;color:var(--muted);font-size:11px}.q-item strong{display:block;margin-top:6px;font-size:13px}.q-item.total{grid-column:1 / -1}
      body.theme-white .hero,body.theme-white .panel,body.theme-white .kpi,body.theme-white .toggle,body.theme-white .legend-item,body.theme-white .chart-frame,body.theme-white .q-card,body.theme-white .q-item{background:#fff}body.theme-white .toggle.active,body.theme-white .legend-item,body.theme-white thead th{background:#f4f7fa}body.theme-white .tooltip{background:#fff}
      @media print{body{--bg:#fff;--card:#fff;--card2:#fff;--line:rgba(34,58,84,.12);--line2:rgba(34,58,84,.2);--text:#22384b;--muted:#647789;--accent:#365c84;--shadow:none;background:#fff}.topbar{display:none}.hero,.panel,.kpi,.toggle,.legend-item,.chart-frame,.q-card,.q-item{background:#fff !important;box-shadow:none !important}}
    </style></head><body class="theme-dark"><div class="wrap"><div class="topbar"><div class="theme-switch"><button class="theme-btn active" id="darkBtn">Dark</button><button class="theme-btn" id="whiteBtn">White</button></div></div><section class="hero"><div class="eyebrow">Tax & Accounting Brief</div><h1>${esc(ctx.company)} ${ctx.latest_year}년 신고 분석</h1><div class="sub">이번 신고 대상과 핵심 수치를 먼저 확인하고, 아래에서 매출 유형별 흐름을 비교하실 수 있습니다.</div><div class="kpis">${[
      ["이번 신고 대상", ctx.current_filing?.label ?? "-"],
      ["이번 납부세액", fmtMoney(ctx.current_filing?.vat_due ?? null)],
      ["매출 대비 매입 비율", fmtPct(ctx.latest_metrics.purchase_ratio)],
      ["매출 대비 인건비 비율", fmtPct(ctx.latest_metrics.labor_ratio)],
    ].map(([label, value]) => `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div></div>`).join("")}</div></section><section class="panel"><h2>요약 의견</h2><ul><li>이번 신고 대상은 ${esc(ctx.current_filing?.label ?? "-")}이며 납부세액은 ${fmtMoney(ctx.current_filing?.vat_due ?? null)}입니다.</li><li>매출 대비 매입 비율은 ${fmtPct(ctx.latest_metrics.purchase_ratio)}, 매출 대비 인건비 비율은 ${fmtPct(ctx.latest_metrics.labor_ratio)}입니다.</li><li>부가세자료 기준 예상손익은 ${fmtMoney(ctx.latest_metrics.profit)}이며, ${ctx.latest_metrics.previous_actual_year}년도 실제 손익은 ${fmtMoney(ctx.latest_metrics.previous_actual_profit)}입니다.</li></ul></section><section class="panel"><h2>최신 연도 핵심 수치</h2><table class="summary-table"><tbody><tr><th>총매출</th><td>${fmtMoney(ctx.latest_metrics.sales)}</td></tr><tr><th>총매입</th><td>${fmtMoney(ctx.latest_metrics.purchase)}</td></tr><tr><th>총 인건비</th><td>${fmtMoney(ctx.latest_metrics.labor)}</td></tr><tr><th>연간 부가세 납부세액 합계</th><td>${fmtMoney(ctx.latest_metrics.vat_total)}</td></tr><tr><th>부가세자료 기준 예상손익</th><td>${fmtMoney(ctx.latest_metrics.profit)}</td></tr><tr><th>${ctx.latest_metrics.previous_actual_year}년도 실제 손익</th><td>${fmtMoney(ctx.latest_metrics.previous_actual_profit)}</td></tr>${ctx.latest_metrics.elapsed_months && ctx.latest_metrics.elapsed_months < 12 ? `<tr><th>${ctx.latest_year} 연환산 예상매출</th><td>${fmtMoney(ctx.latest_metrics.projected_annual_sales)}</td></tr>` : ""}</tbody></table></section><section class="panel"><h2>매출 추이 비교</h2><div class="toggle-wrap">${Object.keys(ctx.sales_series.annual.datasets).map((name, idx) => `<label class="toggle ${idx === 0 ? "active" : ""}" data-series-toggle="${esc(name)}"><input type="checkbox" data-series="${esc(name)}" ${idx === 0 ? "checked" : ""}><span class="dot" style="background:${SERIES_COLORS[name]}"></span><span>${esc(name)}</span></label>`).join("")}</div><div class="chart-frame"><div class="chart-title">연도별 매출 추이</div><div class="chart-sub">${ctx.latest_metrics.elapsed_months && ctx.latest_metrics.elapsed_months < 12 ? `최신연도는 ${ctx.latest_metrics.elapsed_months}개월 기준 실적이며, 연환산 예상매출 점선을 함께 표시합니다.` : `연도별 총매출 기준으로 비교합니다.`}</div><div class="plot"><div class="y-axis" id="yAxis"></div><div class="grid-line" style="top:0%"></div><div class="grid-line" style="top:20%"></div><div class="grid-line" style="top:40%"></div><div class="grid-line" style="top:60%"></div><div class="grid-line" style="top:80%"></div><div class="grid-line" style="top:100%"></div><svg id="salesChart" viewBox="0 0 960 ${mobile ? 220 : 320}"></svg></div><div class="x-axis" id="xAxis"></div><div class="legend" id="chartLegend"></div></div></section><section class="panel"><h2>최신 연도 분기별 수치</h2>${mobile ? Object.entries(ctx.quarter_metrics).map(([label, values]) => {
      const total = values.reduce((sum, item) => sum + (item.value ?? 0), 0);
      return `<div class="q-card"><div style="font-weight:700;margin-bottom:10px;">${label}</div><div class="q-grid">${values.map(item => `<div class="q-item"><span>${item.label}</span><strong>${fmtMoney(item.value)}</strong></div>`).join("")}<div class="q-item total"><span>합계</span><strong>${fmtMoney(total)}</strong></div></div></div>`;
    }).join("") : `<table><thead><tr><th>항목</th>${sharedHeaders.map(v => `<th>${v}</th>`).join("")}<th>합계</th></tr></thead><tbody>${Object.entries(latestDesktopMetrics).map(([label, values]) => {
      const total = values.reduce((sum, item) => sum + (item.value ?? 0), 0);
      return `<tr><th>${label}</th>${values.map(item => `<td>${fmtMoney(item.value)}</td>`).join("")}<td>${fmtMoney(total)}</td></tr>`;
    }).join("")}</tbody></table>`}<div class="note">본 보고서의 손익은 부가세 신고자료를 기준으로 산출한 값입니다. 적격증빙 외 수익 또는 비용은 반영되어 있지 않으며, 4대보험, 각종 지원금, 이자비용 등에 따라 실제 손익과 차이가 발생할 수 있습니다.</div></section>${previousSection}</div><div class="tooltip" id="tooltip"></div><script>
      const DATA=${data};const COLORS=${JSON.stringify(SERIES_COLORS)};const body=document.body,svg=document.getElementById("salesChart"),legend=document.getElementById("chartLegend"),tooltip=document.getElementById("tooltip"),yAxis=document.getElementById("yAxis"),xAxis=document.getElementById("xAxis");
      const chartHeight=${mobile ? 220 : 320}, padding={top:24,bottom:38}, plotW=960, plotH=chartHeight-padding.top-padding.bottom;
      const money=v=>Math.round(v||0).toLocaleString("ko-KR")+"원"; const short=v=>v>=100000000?(v/100000000).toFixed(1)+"억원":v>=10000?Math.round(v/10000)+"만원":money(v);
      function setTheme(t){body.classList.toggle("theme-white",t==="white");document.getElementById("darkBtn").classList.toggle("active",t!=="white");document.getElementById("whiteBtn").classList.toggle("active",t==="white")}
      document.getElementById("darkBtn").onclick=()=>setTheme("dark"); document.getElementById("whiteBtn").onclick=()=>setTheme("white");
      function currentSeries(){ return DATA.sales_series.annual; }
      function maxValue(){ const vals=Object.values(currentSeries().datasets).flat(); if(DATA.latest_metrics.elapsed_months && DATA.latest_metrics.elapsed_months < 12){ vals.push(...DATA.sales_series.projected_total_sales); } return Math.max(1,...vals); }
      function buildAxes(){ const max=maxValue(); yAxis.innerHTML=Array.from({length:6},(_,i)=>'<div class="y-tick" style="top:'+(i/5)*100+'%;">'+short(max*(5-i)/5)+'</div>').join(''); xAxis.style.gridTemplateColumns='repeat('+currentSeries().labels.length+', minmax(0,1fr))'; xAxis.innerHTML=currentSeries().labels.map(v=>'<div>'+String(v)+'</div>').join(''); }
      svg.setAttribute('preserveAspectRatio','none');
      function xCenters(total){
        const svgRect=svg.getBoundingClientRect();
        const axisRect=xAxis.getBoundingClientRect();
        const labels=[...xAxis.children];
        if(labels.length===total && svgRect.width>0 && axisRect.width>0){
          return labels.map(label=>{
            const rect=label.getBoundingClientRect();
            const ratio=(rect.left + rect.width/2 - axisRect.left) / axisRect.width;
            return Math.max(0, Math.min(960, ratio * 960));
          });
        }
        return Array.from({length: total}, (_, i) => 960 * ((i + 0.5) / total));
      }
      const point=(x,v,max)=>({x,y:padding.top+plotH-((v||0)/max)*plotH});
      function render(){const selected=[...document.querySelectorAll('input[data-series]:checked')].map(el=>el.dataset.series); svg.innerHTML=''; legend.innerHTML=''; buildAxes(); const series=currentSeries(); const max=maxValue();
        const centers=xCenters(series.labels.length);
        selected.forEach(name=>{const values=series.datasets[name]||[], color=COLORS[name], pts=values.map((v,i)=>point(centers[i],v,max));
          const poly=document.createElementNS('http://www.w3.org/2000/svg','polyline'); poly.setAttribute('points',pts.map(p=>p.x.toFixed(1)+','+p.y.toFixed(1)).join(' ')); poly.setAttribute('fill','none'); poly.setAttribute('stroke',color); poly.setAttribute('stroke-width',name==='총매출'?'5':'4'); poly.setAttribute('stroke-linecap','round'); poly.setAttribute('stroke-linejoin','round'); svg.appendChild(poly);
          pts.forEach((p,i)=>{const c=document.createElementNS('http://www.w3.org/2000/svg','circle'); c.setAttribute('cx',p.x); c.setAttribute('cy',p.y); c.setAttribute('r',name==='총매출'?'5.5':'4.5'); c.setAttribute('fill',color); const show=e=>{const ev=e.touches?e.touches[0]:e; tooltip.innerHTML='<strong>'+name+'</strong>'+String(series.labels[i])+'<br>'+money(values[i]); tooltip.style.left=(ev.clientX+12)+'px'; tooltip.style.top=(ev.clientY+12)+'px'; tooltip.style.opacity='1';}; c.addEventListener('mousemove',show); c.addEventListener('touchstart',show,{passive:true}); c.addEventListener('touchmove',show,{passive:true}); c.addEventListener('mouseleave',()=>tooltip.style.opacity='0'); c.addEventListener('touchend',()=>tooltip.style.opacity='0'); svg.appendChild(c);});
          const item=document.createElement('div'); item.className='legend-item'; item.innerHTML='<span class="legend-color" style="background:'+color+'"></span><span>'+name+'</span>'; legend.appendChild(item);});
        if(selected.includes('총매출') && DATA.latest_metrics.elapsed_months && DATA.latest_metrics.elapsed_months < 12){
          const values=DATA.sales_series.projected_total_sales, color='#d6e9ff', pts=values.map((v,i)=>point(centers[i],v,max));
          const poly=document.createElementNS('http://www.w3.org/2000/svg','polyline'); poly.setAttribute('points',pts.map(p=>p.x.toFixed(1)+','+p.y.toFixed(1)).join(' ')); poly.setAttribute('fill','none'); poly.setAttribute('stroke',color); poly.setAttribute('stroke-width','3'); poly.setAttribute('stroke-dasharray','8 6'); poly.setAttribute('stroke-linecap','round'); poly.setAttribute('stroke-linejoin','round'); svg.appendChild(poly);
          const lastPoint=pts[pts.length-1]; const lastVal=values[values.length-1]; const c=document.createElementNS('http://www.w3.org/2000/svg','circle'); c.setAttribute('cx',lastPoint.x); c.setAttribute('cy',lastPoint.y); c.setAttribute('r','5'); c.setAttribute('fill',color); const show=e=>{const ev=e.touches?e.touches[0]:e; tooltip.innerHTML='<strong>총매출(연환산 예상)</strong>'+String(series.labels[series.labels.length-1])+'<br>'+money(lastVal); tooltip.style.left=(ev.clientX+12)+'px'; tooltip.style.top=(ev.clientY+12)+'px'; tooltip.style.opacity='1';}; c.addEventListener('mousemove',show); c.addEventListener('touchstart',show,{passive:true}); c.addEventListener('touchmove',show,{passive:true}); c.addEventListener('mouseleave',()=>tooltip.style.opacity='0'); c.addEventListener('touchend',()=>tooltip.style.opacity='0'); svg.appendChild(c);
          const item=document.createElement('div'); item.className='legend-item'; item.innerHTML='<span class="legend-color" style="background:'+color+'"></span><span>총매출(연환산 예상)</span>'; legend.appendChild(item);
        }
      }
      document.querySelectorAll('input[data-series]').forEach(input=>input.addEventListener('change',()=>{input.closest('[data-series-toggle]').classList.toggle('active',input.checked); render();})); render();
    </script></body></html>`;
  }

  function setError(message) {
    errorBox.textContent = message;
    errorBox.classList.toggle("hidden", !message);
  }

  function setDownloadLink(anchor, filename, content) {
    const blob = new Blob([content], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    anchor.href = url;
    anchor.download = filename;
    anchor.classList.remove("hidden");
    return url;
  }

  fileInput.addEventListener("change", event => {
    selectedFile = event.target.files && event.target.files[0] ? event.target.files[0] : null;
    generateBtn.disabled = !selectedFile;
    resultPanel.classList.add("hidden");
    desktopDownload.classList.add("hidden");
    mobileDownload.classList.add("hidden");
    setError("");
    statusEl.textContent = selectedFile ? `선택한 파일: ${selectedFile.name}` : "업로드할 파일을 선택해 주세요.";
  });

  resetBtn.addEventListener("click", () => {
    fileInput.value = "";
    selectedFile = null;
    generateBtn.disabled = true;
    resultPanel.classList.add("hidden");
    desktopDownload.classList.add("hidden");
    mobileDownload.classList.add("hidden");
    setError("");
    statusEl.textContent = "업로드할 파일을 선택해 주세요.";
    if (desktopUrl) URL.revokeObjectURL(desktopUrl);
    if (mobileUrl) URL.revokeObjectURL(mobileUrl);
  });

  generateBtn.addEventListener("click", async () => {
    if (!selectedFile) return;
    try {
      if (!window.XLSX) throw new Error("엑셀 파서 라이브러리를 불러오지 못했습니다.");
      statusEl.textContent = "엑셀을 분석하고 보고서를 생성하는 중입니다...";
      const buffer = await selectedFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const ctx = buildContext(workbook, selectedFile.name);
      if (desktopUrl) URL.revokeObjectURL(desktopUrl);
      if (mobileUrl) URL.revokeObjectURL(mobileUrl);
      desktopUrl = setDownloadLink(desktopDownload, `${ctx.company}_${ctx.latest_year}_desktop.html`, makeReport(ctx, "desktop"));
      mobileUrl = setDownloadLink(mobileDownload, `${ctx.company}_${ctx.latest_year}_mobile.html`, makeReport(ctx, "mobile"));
      document.getElementById("companyName").textContent = ctx.company;
      document.getElementById("latestYear").textContent = ctx.latest_year;
      document.getElementById("filingPattern").textContent = ctx.filing_pattern;
      document.getElementById("currentFiling").textContent = ctx.current_filing ? ctx.current_filing.label : "-";
      document.getElementById("vatDue").textContent = fmtMoney(ctx.current_filing ? ctx.current_filing.vat_due : null);
      document.getElementById("salesTotal").textContent = fmtMoney(ctx.latest_metrics.sales);
      document.getElementById("purchaseRatio").textContent = fmtPct(ctx.latest_metrics.purchase_ratio);
      document.getElementById("laborRatio").textContent = fmtPct(ctx.latest_metrics.labor_ratio);
      resultPanel.classList.remove("hidden");
      statusEl.textContent = "생성이 완료되었습니다. 아래 버튼으로 HTML 파일을 다운로드해 주세요.";
      setError("");
    } catch (error) {
      console.error(error);
      setError(`생성 중 오류가 발생했습니다: ${error.message}`);
      statusEl.textContent = "오류가 발생했습니다.";
    }
  });
})();
