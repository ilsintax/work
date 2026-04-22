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
  const METRIC_LABELS = {
    sales: "매출",
    purchase: "매입",
    vat: "부가세 납부세액",
    labor: "인건비",
    profit: "예상손익",
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

  function displayPeriods(active) {
    const by = Object.fromEntries(active.map(v => [v.label, v]));
    const [q1, q2, q3, q4] = QUARTERS.map(([label]) => label);
    const periods = [];

    if (by[q1]) periods.push({ ...by[q1], sourceLabels: [q1] });
    if (by[q2]) {
      periods.push(by[q1]
        ? { ...by[q2], sourceLabels: [q2] }
        : { ...by[q2], label: "1~2분기", sourceLabels: [q1, q2] });
    }

    if (by[q3]) periods.push({ ...by[q3], sourceLabels: [q3] });
    if (by[q4]) {
      periods.push(by[q3]
        ? { ...by[q4], sourceLabels: [q4] }
        : { ...by[q4], label: "3~4분기", sourceLabels: [q3, q4] });
    }

    return periods;
  }

  function filingPattern(periods) {
    return periods.length ? periods.map(v => v.label).join(" / ") : "-";
  }

  function elapsedMonthsFromLabel(label) {
    if (label === "1분기") return 3;
    if (label === "1~2분기" || label === "2분기") return 6;
    if (label === "3분기") return 9;
    if (label === "3~4분기" || label === "4분기") return 12;
    return null;
  }

  function compressMetric(values, periods) {
    const map = Object.fromEntries(values.map(v => [v.label, v.value]));
    return periods.map(period => {
      const sourceLabels = period.sourceLabels?.length ? period.sourceLabels : [period.label];
      const sourceValues = sourceLabels.map(label => map[label]).filter(value => value != null);
      return {
        label: period.label,
        value: sourceValues.length ? sourceValues.reduce((sum, item) => sum + item, 0) : null,
      };
    });
  }

  function buildQuarterMetrics(ws, periods) {
    return {
      [METRIC_LABELS.sales]: compressMetric(rowQuarters(ws, MAIN_ROWS.sales_total), periods),
      [METRIC_LABELS.purchase]: compressMetric(rowQuarters(ws, MAIN_ROWS.purchase_total), periods),
      [METRIC_LABELS.vat]: compressMetric(rowQuarters(ws, MAIN_ROWS.vat_due), periods),
      [METRIC_LABELS.labor]: compressMetric(rowQuarters(ws, MAIN_ROWS.labor_total), periods),
      [METRIC_LABELS.profit]: compressMetric(rowQuarters(ws, MAIN_ROWS.profit_est), periods),
    };
  }

  function buildContext(workbook, filename) {
    const years = workbook.SheetNames.filter(name => /^\d+$/.test(name)).sort((a, b) => Number(a) - Number(b));
    if (!years.length) throw new Error("연도 시트를 찾지 못했습니다.");

    const latestYear = years[years.length - 1];
    const previousYear = years.length > 1 ? years[years.length - 2] : null;
    const ws = workbook.Sheets[latestYear];
    const periods = displayPeriods(activeQuarters(ws));

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
      const prevPeriods = displayPeriods(activeQuarters(prevWs));
      previousQuarterMetrics = {
        year: previousYear,
        periods: prevPeriods,
        metrics: buildQuarterMetrics(prevWs, prevPeriods),
      };
    }

    return {
      company: cell(ws, "E5") || filename.replace(/\.[^.]+$/, ""),
      latest_year: latestYear,
      filing_pattern: filingPattern(periods),
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
      quarter_metrics: buildQuarterMetrics(ws, periods),
      previous_quarter_metrics: previousQuarterMetrics,
      sales_series: {
        annual: { labels: years, datasets: annualDatasets },
        projected_total_sales: projectedAnnualSeries,
      },
    };
  }

  function buildQuarterLayout(periods) {
    const starts = { "1분기": 0, "2분기": 1, "3분기": 2, "4분기": 3 };
    const layout = Array.from({ length: 4 }, () => null);
    periods.forEach(period => {
      const sourceLabels = period.sourceLabels?.length ? period.sourceLabels : [period.label];
      const start = starts[sourceLabels[0]];
      if (start == null) return;
      layout[start] = { period, span: sourceLabels.length };
    });
    return layout;
  }

  function renderMetricCards(metrics, periods) {
    const layout = buildQuarterLayout(periods);
    return Object.entries(metrics).map(([label, values]) => {
      const total = values.reduce((sum, item) => sum + (item.value ?? 0), 0);
      const valueMap = Object.fromEntries(values.map(item => [item.label, item.value]));
      const cards = [];
      for (let idx = 0; idx < layout.length; idx += 1) {
        const slot = layout[idx];
        if (!slot) {
          cards.push(`<div class="q-item ghost" aria-hidden="true"></div>`);
          continue;
        }
        cards.push(`<div class="q-item period span-${slot.span}"><span>${slot.period.label}</span><strong>${fmtMoney(valueMap[slot.period.label] ?? null)}</strong></div>`);
        idx += slot.span - 1;
      }
      return `<div class="q-card"><div style="font-weight:700;margin-bottom:10px;">${label}</div><div class="q-grid">${cards.join("")}<div class="q-item total"><span>합계</span><strong>${fmtMoney(total)}</strong></div></div></div>`;
    }).join("");
  }

  function renderMetricTable(metrics, periods) {
    const layout = buildQuarterLayout(periods);
    const headerCells = [];
    for (let idx = 0; idx < layout.length; idx += 1) {
      const slot = layout[idx];
      if (!slot) {
        headerCells.push(`<th class="quarter-head ghost"></th>`);
        continue;
      }
      headerCells.push(`<th class="quarter-head auto-fit-print" colspan="${slot.span}">${slot.period.label}</th>`);
      idx += slot.span - 1;
    }

    return `<table class="report-table"><thead><tr><th class="auto-fit-print">항목</th>${headerCells.join("")}<th class="auto-fit-print">합계</th></tr></thead><tbody>${Object.entries(metrics).map(([label, values]) => {
      const total = values.reduce((sum, item) => sum + (item.value ?? 0), 0);
      const valueMap = Object.fromEntries(values.map(item => [item.label, item.value]));
      const rowCells = [];
      for (let idx = 0; idx < layout.length; idx += 1) {
        const slot = layout[idx];
        if (!slot) {
          rowCells.push(`<td class="quarter-cell ghost"></td>`);
          continue;
        }
        rowCells.push(`<td class="quarter-cell auto-fit-print" colspan="${slot.span}">${fmtMoney(valueMap[slot.period.label] ?? null)}</td>`);
        idx += slot.span - 1;
      }
      return `<tr><th class="auto-fit-print">${label}</th>${rowCells.join("")}<td class="auto-fit-print">${fmtMoney(total)}</td></tr>`;
    }).join("")}</tbody></table>`;
  }

  function makeReport(ctx, mode) {
    const mobile = mode === "mobile";
    const data = JSON.stringify(ctx).replace(/</g, "\\u003c");
    const heroTitle = mobile
      ? `<span class="hero-company">${esc(ctx.company)}</span><span class="hero-report">${ctx.latest_year}년 신고 분석</span>`
      : `${esc(ctx.company)} ${ctx.latest_year}년 신고 분석`;

    const latestQuarterSection = mobile
      ? renderMetricCards(ctx.quarter_metrics, ctx.active_periods)
      : renderMetricTable(ctx.quarter_metrics, ctx.active_periods);

    const summaryItems = [
      `이번 신고 대상은 ${esc(ctx.current_filing?.label ?? "-")}이며 납부세액은 ${fmtMoney(ctx.current_filing?.vat_due ?? null)}입니다.`,
      `매출 대비 매입 비율은 ${fmtPct(ctx.latest_metrics.purchase_ratio)}, 매출 대비 인건비 비율은 ${fmtPct(ctx.latest_metrics.labor_ratio)}입니다.`,
      `부가세자료 기준 예상손익은 ${fmtMoney(ctx.latest_metrics.profit)}이며, ${ctx.latest_metrics.previous_actual_year}년도 실제 손익은 ${fmtMoney(ctx.latest_metrics.previous_actual_profit)}입니다.`,
    ];

    const previousSection = ctx.previous_quarter_metrics ? (mobile
      ? `<section class="panel previous-quarter-section"><h2>${ctx.previous_quarter_metrics.year}년 분기별 수치</h2>${renderMetricCards(ctx.previous_quarter_metrics.metrics, ctx.previous_quarter_metrics.periods)}</section>`
      : `<section class="panel previous-quarter-section"><h2>${ctx.previous_quarter_metrics.year}년 분기별 수치</h2>${renderMetricTable(ctx.previous_quarter_metrics.metrics, ctx.previous_quarter_metrics.periods)}<div class="note">전년도는 전년도 실제 입력 형태에 맞춰 분기를 표시했습니다.</div></section>`)
      : "";

    return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(ctx.company)} ${ctx.latest_year}</title><style>
      :root{--bg:#0d1825;--card:rgba(20,34,50,.96);--card2:rgba(16,29,43,.98);--line:rgba(153,184,216,.14);--line2:rgba(153,184,216,.26);--text:#e7eff8;--muted:#9eb0c4;--accent:#8bbce8;--shadow:0 14px 34px rgba(0,0,0,.22)}
      body.theme-white{--bg:#fff;--card:#fff;--card2:#fff;--line:rgba(34,58,84,.12);--line2:rgba(34,58,84,.2);--text:#22384b;--muted:#647789;--accent:#365c84;--shadow:0 10px 24px rgba(26,46,68,.08)}
      *{box-sizing:border-box}body{margin:0;font-family:"Malgun Gothic","Apple SD Gothic Neo",sans-serif;background:radial-gradient(circle at top right,rgba(184,166,132,.14),transparent 28%),linear-gradient(180deg,#dfe4e8 0%,#cfd6dc 100%);color:var(--text)}body.theme-white{background:#fff}
      .wrap{max-width:${mobile ? "480px" : "1240px"};margin:0 auto;padding:${mobile ? "16px 14px 32px" : "28px 18px 56px"}}.topbar{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;gap:10px}.top-actions{display:flex;gap:8px;flex-wrap:wrap}.theme-switch{display:inline-flex;gap:6px;border:1px solid var(--line);border-radius:999px;padding:6px;background:rgba(18,31,46,.92)}body.theme-white .theme-switch{background:#f5f8fb}.theme-btn,.print-btn{border:0;background:transparent;color:var(--muted);padding:8px 12px;border-radius:999px;font:inherit;cursor:pointer}.theme-btn.active{background:rgba(139,188,232,.14);color:var(--text)}.print-btn{border:1px solid var(--line);background:rgba(18,31,46,.92);color:var(--text)}body.theme-white .print-btn{background:#f5f8fb}
      .hero,.panel{background:linear-gradient(180deg,var(--card) 0%,var(--card2) 100%);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow)}.hero{padding:${mobile ? "20px" : "30px"}}.panel{margin-top:14px;padding:${mobile ? "18px" : "24px"}}.eyebrow{color:var(--accent);font-size:11px;letter-spacing:.12em;text-transform:uppercase;font-weight:700}h1{margin:10px 0 ${mobile ? "14px" : "12px"};font-size:${mobile ? "21px" : "32px"};line-height:${mobile ? "1.24" : "1.25"};letter-spacing:${mobile ? "-0.03em" : "normal"};word-break:keep-all}.hero-company,.hero-report{display:block}.hero-report{margin-top:${mobile ? "8px" : "0"}}.sub,.note{color:var(--muted);line-height:1.8}.panel h2{margin:0 0 12px;font-size:${mobile ? "18px" : "22px"}}ul{margin:0;padding-left:18px;line-height:1.8}
      .kpis{display:grid;grid-template-columns:${mobile ? "1fr" : "repeat(4,minmax(0,1fr))"};gap:12px;margin-top:14px}.kpi{background:rgba(25,42,60,.92);border:1px solid var(--line);border-radius:8px;padding:16px;min-width:0;overflow:hidden}.label{color:var(--muted);font-size:11px;text-transform:uppercase}.value{margin-top:8px;font-size:${mobile ? "24px" : "27px"};font-weight:700;white-space:nowrap;word-break:keep-all;overflow-wrap:normal;letter-spacing:-0.02em;font-variant-numeric:tabular-nums;display:block;width:100%;max-width:100%;overflow:hidden}
      .toggle-wrap,.legend{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}.toggle,.legend-item{display:inline-flex;align-items:center;gap:8px;padding:8px 10px;border-radius:4px;border:1px solid var(--line);background:rgba(19,34,50,.88);color:var(--muted);font-size:12px}.toggle.active{background:rgba(139,188,232,.12);color:var(--text)}.toggle input{display:none}.dot,.legend-color{width:10px;height:10px;border-radius:50%}
      .chart-frame{border:1px solid var(--line);background:rgba(18,31,46,.92);border-radius:8px;padding:${mobile ? "16px 12px 14px 56px" : "22px 22px 20px 80px"}}.chart-title{font-size:${mobile ? "18px" : "22px"};font-weight:700}.chart-sub{color:var(--muted);font-size:12px;margin-top:4px}.plot{position:relative;height:${mobile ? "220px" : "320px"};margin-top:12px}.y-axis{position:absolute;left:${mobile ? "-52px" : "-78px"};top:0;bottom:0;width:${mobile ? "48px" : "72px"}}.y-tick{position:absolute;left:0;width:${mobile ? "48px" : "72px"};transform:translateY(-50%);text-align:right;color:var(--muted);font-size:${mobile ? "10px" : "12px"}}.grid-line{position:absolute;left:0;right:0;height:1px;background:rgba(153,184,216,.12)}svg{width:100%;height:${mobile ? "220px" : "320px"};display:block}.x-axis{display:grid;margin-top:8px;color:var(--muted);font-size:${mobile ? "11px" : "13px"};text-align:center}.x-axis div{white-space:pre-line}
      .tooltip{position:fixed;z-index:30;pointer-events:none;padding:8px 10px;border-radius:6px;border:1px solid var(--line2);background:rgba(12,22,34,.98);color:var(--text);font-size:12px;line-height:1.5;opacity:0}.tooltip strong{display:block;color:var(--accent);font-size:11px}
      table{width:100%;border-collapse:collapse;font-size:14px;table-layout:fixed}thead th{background:rgba(139,188,232,.06);color:var(--muted);font-size:12px;text-transform:uppercase}th,td{padding:12px 10px;border-bottom:1px solid rgba(153,184,216,.14);text-align:right}th:first-child,td:first-child{text-align:left;width:160px}.summary-table th:first-child{width:${mobile ? "160px" : "280px"};white-space:nowrap;font-size:${mobile ? "13px" : "14px"}}.summary-table td{font-variant-numeric:tabular-nums}.summary-table td.auto-fit{font-size:${mobile ? "clamp(11px, 3.5vw, 14px)" : "14px"};white-space:nowrap}.summary-table td.auto-fit.tight{font-size:${mobile ? "11px" : "13px"}}.summary-table td.auto-fit.xtight{font-size:${mobile ? "10px" : "12px"}}.report-table .auto-fit-print{white-space:nowrap}.report-table .auto-fit-print.tight{font-size:12px}.report-table .auto-fit-print.xtight{font-size:11px}.quarter-head,.quarter-cell{text-align:center}.quarter-head.ghost,.quarter-cell.ghost{background:transparent;color:transparent}
      .q-card{border:1px solid var(--line);border-radius:8px;padding:14px;background:rgba(18,31,46,.88);margin-top:10px}.q-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}.q-item{border:1px solid var(--line);border-radius:6px;padding:10px;background:rgba(139,188,232,.08)}.q-item span{display:block;color:var(--muted);font-size:11px}.q-item strong{display:block;margin-top:6px;font-size:13px}.q-item.period{text-align:center;display:flex;flex-direction:column;justify-content:center;min-height:72px}.q-item.period.span-2{grid-column:span 2}.q-item.ghost{background:transparent;border-style:dashed;border-color:transparent;min-height:72px}.q-item.total{grid-column:1 / -1}
      body.theme-white .hero,body.theme-white .panel,body.theme-white .kpi,body.theme-white .toggle,body.theme-white .legend-item,body.theme-white .chart-frame,body.theme-white .q-card,body.theme-white .q-item{background:#fff}body.theme-white .toggle.active,body.theme-white .legend-item,body.theme-white thead th{background:#f4f7fa}body.theme-white .tooltip{background:#fff}
      @media print{@page{size:A4 portrait;margin:12mm}body{--bg:#fff;--card:#fff;--card2:#fff;--line:rgba(34,58,84,.12);--line2:rgba(34,58,84,.2);--text:#22384b;--muted:#647789;--accent:#365c84;--shadow:none;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}.topbar,.tooltip{display:none !important}.wrap{max-width:none;padding:0}.hero,.panel,.kpi,.toggle,.legend-item,.chart-frame,.q-card,.q-item{background:#fff !important;box-shadow:none !important}.hero,.panel,.chart-frame,.q-card,table,tr{break-inside:avoid;page-break-inside:avoid}section{break-inside:avoid;page-break-inside:avoid}.plot,.chart-frame,svg{overflow:visible !important}.chart-frame{padding:18px 18px 18px 64px}.kpis{grid-template-columns:repeat(2,minmax(0,1fr)) !important;gap:10px}.kpi{padding:14px;min-height:92px}.kpi .label{font-size:10px}.kpi .value{font-size:23px !important;white-space:nowrap;overflow:visible}${mobile ? `.latest-quarter-section,.previous-quarter-section{break-before:page;page-break-before:always}.latest-quarter-section{padding:14px}.latest-quarter-section h2,.previous-quarter-section h2{margin-bottom:8px}.latest-quarter-section .q-card{margin-top:6px;padding:8px}.latest-quarter-section .q-grid{gap:6px}.latest-quarter-section .q-item{padding:7px;min-height:52px}.latest-quarter-section .q-item span{font-size:9px}.latest-quarter-section .q-item strong{font-size:11px}.latest-quarter-section .note{margin-top:6px;font-size:10px;line-height:1.5}.previous-quarter-section .q-card{margin-top:8px;padding:10px}.previous-quarter-section .q-grid{gap:6px}.previous-quarter-section .q-item{padding:8px;min-height:58px}.previous-quarter-section .q-item span{font-size:10px}.previous-quarter-section .q-item strong{font-size:12px}` : ""}}
    </style></head><body class="theme-dark"><div class="wrap"><div class="topbar"><button class="print-btn" id="printBtn">PDF 저장</button><div class="top-actions"><div class="theme-switch"><button class="theme-btn active" id="darkBtn">Dark</button><button class="theme-btn" id="whiteBtn">White</button></div></div></div><section class="hero"><div class="eyebrow">Tax & Accounting Brief</div><h1>${heroTitle}</h1><div class="sub">이번 신고 대상과 핵심 수치를 먼저 확인하고, 아래에서 매출 유형별 흐름을 비교하실 수 있습니다.</div><div class="kpis">${[
      ["이번 신고 대상", ctx.current_filing?.label ?? "-"],
      ["이번 납부세액", fmtMoney(ctx.current_filing?.vat_due ?? null)],
      ["매출 대비 매입 비율", fmtPct(ctx.latest_metrics.purchase_ratio)],
      ["매출 대비 인건비 비율", fmtPct(ctx.latest_metrics.labor_ratio)],
    ].map(([label, value]) => `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div></div>`).join("")}</div></section><section class="panel"><h2>요약 의견</h2><ul>${summaryItems.map(item => `<li>${item}</li>`).join("")}</ul></section><section class="panel"><h2>최신 연도 핵심 수치</h2><table class="summary-table"><tbody><tr><th>총매출</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.sales)}</td></tr><tr><th>총매입</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.purchase)}</td></tr><tr><th>총 인건비</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.labor)}</td></tr><tr><th>연간 부가세 납부세액 합계</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.vat_total)}</td></tr><tr><th>부가세자료 기준 예상손익</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.profit)}</td></tr><tr><th>${ctx.latest_metrics.previous_actual_year}년도 실제 손익</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.previous_actual_profit)}</td></tr>${ctx.latest_metrics.elapsed_months && ctx.latest_metrics.elapsed_months < 12 ? `<tr><th>${ctx.latest_year} 연환산 예상매출</th><td class="auto-fit">${fmtMoney(ctx.latest_metrics.projected_annual_sales)}</td></tr>` : ""}</tbody></table></section><section class="panel"><h2>매출 추이 비교</h2><div class="toggle-wrap">${Object.keys(ctx.sales_series.annual.datasets).map((name, idx) => `<label class="toggle ${idx === 0 ? "active" : ""}" data-series-toggle="${esc(name)}"><input type="checkbox" data-series="${esc(name)}" ${idx === 0 ? "checked" : ""}><span class="dot" style="background:${SERIES_COLORS[name]}"></span><span>${esc(name)}</span></label>`).join("")}</div><div class="chart-frame"><div class="chart-title">연도별 매출 추이</div><div class="chart-sub">${ctx.latest_metrics.elapsed_months && ctx.latest_metrics.elapsed_months < 12 ? `최신연도는 ${ctx.latest_metrics.elapsed_months}개월 기준 실적이며, 연환산 예상매출 점선을 함께 표시합니다.` : `연도별 총매출 기준으로 비교합니다.`}</div><div class="plot"><div class="y-axis" id="yAxis"></div><div class="grid-line" style="top:0%"></div><div class="grid-line" style="top:20%"></div><div class="grid-line" style="top:40%"></div><div class="grid-line" style="top:60%"></div><div class="grid-line" style="top:80%"></div><div class="grid-line" style="top:100%"></div><svg id="salesChart" viewBox="0 0 960 ${mobile ? 220 : 320}"></svg></div><div class="x-axis" id="xAxis"></div><div class="legend" id="chartLegend"></div></div></section><section class="panel latest-quarter-section"><h2>최신 연도 분기별 수치</h2>${latestQuarterSection}<div class="note">본 보고서의 손익은 부가세 신고자료를 기준으로 산출한 값입니다. 적격증빙 외 수익 또는 비용은 반영되어 있지 않으며, 4대보험, 각종 지원금, 이자비용 등에 따라 실제 손익과 차이가 발생할 수 있습니다.</div></section>${previousSection}</div><div class="tooltip" id="tooltip"></div><script>
      const DATA=${data};const COLORS=${JSON.stringify(SERIES_COLORS)};const body=document.body,svg=document.getElementById("salesChart"),legend=document.getElementById("chartLegend"),tooltip=document.getElementById("tooltip"),yAxis=document.getElementById("yAxis"),xAxis=document.getElementById("xAxis"),gridLines=[...document.querySelectorAll('.grid-line')];
      const chartHeight=${mobile ? 220 : 320}, padding={top:24,bottom:38}, plotH=chartHeight-padding.top-padding.bottom;
      const money=v=>Math.round(v||0).toLocaleString("ko-KR")+"원"; const short=v=>v>=100000000?(v/100000000).toFixed(1)+"억원":v>=10000?Math.round(v/10000)+"만원":money(v);
      function setTheme(t){body.classList.toggle("theme-white",t==="white");document.getElementById("darkBtn").classList.toggle("active",t!=="white");document.getElementById("whiteBtn").classList.toggle("active",t==="white")}
      document.getElementById("darkBtn").onclick=()=>setTheme("dark"); document.getElementById("whiteBtn").onclick=()=>setTheme("white");
      document.getElementById("printBtn").onclick=()=>window.print();
      function fitMobileSummaryValues(){
        if(!window.matchMedia('(max-width: 640px)').matches) return;
        document.querySelectorAll('.summary-table td.auto-fit').forEach(cell=>{
          cell.classList.remove('tight','xtight');
          if(cell.scrollWidth > cell.clientWidth) cell.classList.add('tight');
          if(cell.scrollWidth > cell.clientWidth) cell.classList.add('xtight');
        });
      }
      function fitTextToWidth(cell, maxSize, minSize){
        const container = cell.parentElement || cell;
        const cellStyle = window.getComputedStyle(cell);
        const containerStyle = window.getComputedStyle(container);
        const available = Math.max(
          0,
          container.clientWidth
            - parseFloat(containerStyle.paddingLeft || 0)
            - parseFloat(containerStyle.paddingRight || 0)
            - parseFloat(cellStyle.marginLeft || 0)
            - parseFloat(cellStyle.marginRight || 0)
        );
        if (!available) {
          cell.style.fontSize = maxSize + 'px';
          return;
        }

        const probe = document.createElement('span');
        probe.textContent = cell.textContent || '';
        probe.style.position = 'absolute';
        probe.style.visibility = 'hidden';
        probe.style.whiteSpace = 'nowrap';
        probe.style.pointerEvents = 'none';
        probe.style.fontFamily = cellStyle.fontFamily;
        probe.style.fontWeight = cellStyle.fontWeight;
        probe.style.letterSpacing = cellStyle.letterSpacing;
        probe.style.fontVariantNumeric = cellStyle.fontVariantNumeric;
        document.body.appendChild(probe);

        let low = minSize;
        let high = maxSize;
        let best = minSize;
        while (low <= high) {
          const mid = Math.floor((low + high) / 2);
          probe.style.fontSize = mid + 'px';
          if (probe.getBoundingClientRect().width <= available) {
            best = mid;
            low = mid + 1;
          } else {
            high = mid - 1;
          }
        }

        cell.style.fontSize = best + 'px';
        probe.remove();
      }
      function fitKpiValues(){
        document.querySelectorAll('.kpi .value').forEach(cell=>{
          const maxSize = window.matchMedia('(max-width: 640px)').matches ? 24 : 27;
          const minSize = window.matchMedia('(max-width: 640px)').matches ? 12 : 12;
          fitTextToWidth(cell, maxSize, minSize);
        });
      }
      function fitDesktopPrintTables(){
        document.querySelectorAll('.report-table .auto-fit-print').forEach(cell=>{
          cell.classList.remove('tight','xtight');
          if(cell.scrollWidth > cell.clientWidth) cell.classList.add('tight');
          if(cell.scrollWidth > cell.clientWidth) cell.classList.add('xtight');
        });
      }
      function currentSeries(){ return DATA.sales_series.annual; }
      function maxValue(){ const vals=Object.values(currentSeries().datasets).flat(); if(DATA.latest_metrics.elapsed_months && DATA.latest_metrics.elapsed_months < 12){ vals.push(...DATA.sales_series.projected_total_sales); } return Math.max(1,...vals); }
      function axisTopPct(i){ return ((padding.top + (plotH * i / 5)) / chartHeight) * 100; }
      function buildAxes(){ const max=maxValue(); yAxis.innerHTML=Array.from({length:6},(_,i)=>'<div class="y-tick" style="top:'+axisTopPct(i)+'%;">'+short(max*(5-i)/5)+'</div>').join(''); gridLines.forEach((line,i)=>line.style.top=axisTopPct(i)+'%'); xAxis.style.gridTemplateColumns='repeat('+currentSeries().labels.length+', minmax(0,1fr))'; xAxis.innerHTML=currentSeries().labels.map(v=>'<div>'+String(v)+'</div>').join(''); }
      svg.setAttribute('preserveAspectRatio','none');
      function xCenters(total){ return Array.from({length: total}, (_, i) => 960 * ((i + 0.5) / total)); }
      const point=(x,v,max)=>({x,y:padding.top+plotH-((v||0)/max)*plotH});
      function render(){const selected=[...document.querySelectorAll('input[data-series]:checked')].map(el=>el.dataset.series); svg.innerHTML=''; legend.innerHTML=''; buildAxes(); const series=currentSeries(); const max=maxValue(); const centers=xCenters(series.labels.length);
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
      document.querySelectorAll('input[data-series]').forEach(input=>input.addEventListener('change',()=>{input.closest('[data-series-toggle]').classList.toggle('active',input.checked); render();}));
      function runAllFits(){
        fitMobileSummaryValues();
        fitDesktopPrintTables();
        fitKpiValues();
      }
      render(); runAllFits(); requestAnimationFrame(runAllFits); setTimeout(runAllFits, 60);
      window.addEventListener('load', ()=>{ runAllFits(); requestAnimationFrame(runAllFits); setTimeout(runAllFits, 60); });
      window.addEventListener('resize', ()=>{ runAllFits(); });
      window.addEventListener('beforeprint', fitDesktopPrintTables);
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
    statusEl.textContent = selectedFile ? `선택된 파일: ${selectedFile.name}` : "업로드할 파일을 선택해 주세요.";
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
      if (!window.XLSX) throw new Error("엑셀 파일 처리 라이브러리를 불러오지 못했습니다.");
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
