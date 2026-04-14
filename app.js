/* global XLSX, Chart */

const I18N = {
  title: "\uD1B5\uD569 \uC2DC\uC7A5 \uBD84\uC11D",
  subtitle:
    "\uAC00\uACA9\u00B7\uAC70\uB798\uB7C9, \uAE30\uAD00\u00B7\uC678\uAD6D\uC778 \uC21C\uB9E4\uC218\uB97C \uC77C\uC790\uBCC4\uB85C \uACB0\uD569\uD574 \uAE30\uAC04 \uC218\uC775\uB960, \uC21C\uB9E4\uC218 \uD750\uB984, \uC774\uC288 \uD6C4\uBCF4 \uC885\uBAA9\uC744 \uD55C \uD654\uBA74\uC5D0\uC11C \uD655\uC778\uD569\uB2C8\uB2E4.",
  uploadHeading: "\uB370\uC774\uD130 \uBD88\uB7EC\uC624\uAE30",
  uploadHint:
    "\uD3F4\uB354 \uB370\uC774\uD1302\uc758 \uC138 \uD30C\uC77C(\uAC00\uACA9\u00B7\uAC70\uB798\uB7C9, \uAE30\uAD00, \uC678\uAD6D\uC778)\uC744 \uAC01\uAC01 \uC120\uD0DD\uD558\uAC70\uB098 \uC544\uB798\uC5D0 \uB4DC\uB798\uADF8\uC564\uB4DC\uB86D\uD558\uC138\uC694.",
  dropZone: "\uC138 \uAC1C\uC758 xlsx \uD30C\uC77C\uC744 \uC5EC\uAE30\uC5D0 \uB4DC\uB798\uADF8\uD558\uC138\uC694.",
  lblPrice: "\uAC00\uACA9\u00B7\uAC70\uB798\uB7C9 (.xlsx)",
  lblInst: "\uAE30\uAD00 (.xlsx)",
  lblFor: "\uC678\uAD6D\uC778 (.xlsx)",
  tabRank: "\uAE30\uAC04\uBCC4 \uC218\uC775\uB960",
  tabDetail: "\uC885\uBAA9 \uC0C1\uC138 \u00B7 \uCC28\uD2B8",
  tabAlerts: "\uC774\uC288 \uC885\uBAA9 \uC2A4\uCE94",
  tabReport: "\uC694\uC57D \uB9AC\uD3EC\uD2B8",
  lblAnchor: "\uAE30\uC900\uC77C",
  lblMarket: "\uC2DC\uC7A5",
  optAll: "\uC804\uCCB4",
  lblHorizon: "\uAE30\uAC04",
  h1d: "1\uC77C",
  h1w: "1\uC8FC (\uC601\uC5C5\uC77C)",
  h1m: "1\uAC1C\uC6D4 (\uC601\uC5C5\uC77C)",
  lblTopN: "\uD45C\uC2DC \uAC1C\uC218",
  btnRefresh: "\uC0C8\uB85C\uACE0\uCE68",
  h3Up: "\uC0C1\uC2B9\uB960 \uC0C1\uC704",
  h3Down: "\uD558\uB77D\uB960 \uC0C1\uC704",
  h3Broker:
    "\uAC70\uB798\uB300\uAE08\u00B7\uBCC0\uB3D9\uC131 \uAD00\uC810 \uD6C4\uBCF4 (\uBE0C\uB85C\uCEE4 \uAD00\uC810 \uC9C0\uD45C)",
  hintBroker:
    "\uCD5C\uADFC \uAD6C\uAC04\uC758 \uD3C9\uADE0 \uAC70\uB798\uB300\uAE08\uACFC \uC77C\uAC04 \uC218\uC775\uB960 \uD45C\uC900\uD3B8\uCC28\uB97C \uACB0\uD569\uD55C \uB2E8\uC21C \uC2A4\uCF54\uC5B4\uC785\uB2C8\uB2E4. \uC2E4\uC81C \uC218\uC218\uB8CC\uC640\uB294 \uB2E4\uB97C \uC218 \uC788\uC2B5\uB2C8\uB2E4.",
  lblSearch: "\uC885\uBAA9 \uAC80\uC0C9",
  phSearch: "\uCF54\uB4DC \uB610\uB294 \uC885\uBAA9\uBA85",
  btnChart: "\uCC28\uD2B8 \uBCF4\uAE30",
  h3Series: "\uC77C\uC790\uBCC4 \uC694\uC57D",
  lblAlertDate: "\uAE30\uC900\uC77C",
  lblAlertPct: "|\uB4F1\uB77D\uB960| \u2265 (%)",
  lblAlertVol: "\uAC70\uB798\uB7C9 \uBC30\uC218 (\uC911\uC559\uAC12 \uB300\uBE44)",
  chkInst: "\uAE30\uAD00 \uC21C\uB9E4\uC218 > 0",
  chkFor: "\uC678\uAD6D\uC778 \uC21C\uB9E4\uC218 > 0",
  btnScan: "\uC2A4\uCE94",
  lblRepStart: "\uC2DC\uC791\uC77C",
  lblRepEnd: "\uC885\uB8CC\uC77C",
  btnReport: "\uB9AC\uD3EC\uD2B8 \uC0DD\uC131",
  btnCopy: "\uD074\uB9BD\uBCF4\uB4DC\uC5D0 \uBCF5\uC0AC",
  footer:
    "\uB274\uC2A4\u00B7\uACF5\uC2DC \uC5F0\uB3D9 \uC790\uB3D9 \uC694\uC57D\uC740 \uBCC4\uB3C4 \uB370\uC774\uD130/API\uAC00 \uD544\uC694\uD569\uB2C8\uB2E4. \uD604\uC7AC \uBC84\uC804\uC740 \uC5C5\uB85C\uB4DC\uD55C \uC77C\uBCC4 \uC2A4\uB0B5\uC0F7\uC73C\uB85C \uC21C\uB9E4\uC218\u00B7\uAC00\uACA9 \uC0C1\uAD00\uC744 \uBD84\uC11D\uD569\uB2C8\uB2E4.",
  loadOk: (n) =>
    `\uB85C\uB4DC \uC644\uB8CC: ${n}\uC77C, \uC885\uBAA9 \uC2DC\uB9AC\uC988 \uAC00\uB2A5.`,
  loadNeed: "\uC138 \uD30C\uC77C\uC744 \uBAA8\uB450 \uC120\uD0DD\uD558\uC138\uC694.",
  errParse: (m) => `\uD30C\uC2F1 \uC624\uB958: ${m}`,
};

function applyI18n() {
  const set = (id, t) => {
    const el = document.getElementById(id);
    if (el) el.textContent = t;
  };
  set("uiTitle", I18N.title);
  set("uiSubtitle", I18N.subtitle);
  set("uiUploadHeading", I18N.uploadHeading);
  set("uiUploadHint", I18N.uploadHint);
  set("dropZone", I18N.dropZone);
  set("uiLblPrice", I18N.lblPrice);
  set("uiLblInst", I18N.lblInst);
  set("uiLblFor", I18N.lblFor);
  set("tabBtnRank", I18N.tabRank);
  set("tabBtnDetail", I18N.tabDetail);
  set("tabBtnAlerts", I18N.tabAlerts);
  set("tabBtnReport", I18N.tabReport);
  set("lblAnchor", I18N.lblAnchor);
  set("lblMarket", I18N.lblMarket);
  set("optMarketAll", I18N.optAll);
  set("lblHorizon", I18N.lblHorizon);
  set("optH1d", I18N.h1d);
  set("optH1w", I18N.h1w);
  set("optH1m", I18N.h1m);
  set("lblTopN", I18N.lblTopN);
  set("btnRank", I18N.btnRefresh);
  set("h3Up", I18N.h3Up);
  set("h3Down", I18N.h3Down);
  set("h3Broker", I18N.h3Broker);
  set("hintBroker", I18N.hintBroker);
  set("lblSearch", I18N.lblSearch);
  const search = document.getElementById("stockSearch");
  if (search) search.placeholder = I18N.phSearch;
  set("btnLoadStock", I18N.btnChart);
  set("h3Series", I18N.h3Series);
  set("lblAlertDate", I18N.lblAlertDate);
  set("lblAlertPct", I18N.lblAlertPct);
  set("lblAlertVol", I18N.lblAlertVol);
  set("lblChkInst", I18N.chkInst);
  set("lblChkFor", I18N.chkFor);
  set("btnScan", I18N.btnScan);
  set("lblRepStart", I18N.lblRepStart);
  set("lblRepEnd", I18N.lblRepEnd);
  set("btnReport", I18N.btnReport);
  set("btnCopy", I18N.btnCopy);
  set("uiFooter", I18N.footer);
  document.title = `${I18N.title} — \uAC00\uACA9\u00B7\uAE30\uAD00\u00B7\uC678\uAD6D\uC778`;
}

/** @type {{ sortedDates: string[], byDate: Map<string, any>, series: Map<string, Row[]> }} */
let state = {
  sortedDates: [],
  byDate: new Map(),
  series: new Map(),
};

let chartPrice = null;
let chartFlow = null;

function normStr(v) {
  if (v == null || v === "") return "";
  return String(v).trim();
}

function normCode(v) {
  const s = normStr(v);
  if (!s) return "";
  if (/^\d+$/.test(s)) return String(parseInt(s, 10));
  return s;
}

function sheetNameToIso(name) {
  const m = normStr(name).match(/^(\d{4})\.(\d{2})\.(\d{2})$/);
  if (!m) return null;
  return `${m[1]}-${m[2]}-${m[3]}`;
}

function parseNum(v) {
  if (v == null || v === "") return NaN;
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "" || s === "-") return NaN;
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function buildHeaderMap(row) {
  const map = {};
  row.forEach((cell, i) => {
    const k = normStr(cell).replace(/\s+/g, "");
    if (k) map[k] = i;
  });
  return map;
}

function findCol(map, predicates) {
  const keys = Object.keys(map);
  for (const p of predicates) {
    for (const k of keys) {
      if (p(k)) return map[k];
    }
  }
  return -1;
}

function parsePriceSheet(rows) {
  if (!rows.length) return [];
  const h = buildHeaderMap(rows[0]);
  let iCode = findCol(h, [(kk) => kk.includes("\uC885\uBAA9\uCF54\uB4DC") || kk.toLowerCase().includes("code")]);
  if (iCode < 0) iCode = 0;
  const iName = findCol(h, [(k) => k.includes("\uC885\uBAA9\uBA85")]) >= 0 ? findCol(h, [(k) => k.includes("\uC885\uBAA9\uBA85")]) : 1;
  const iMkt = findCol(h, [(k) => k.includes("\uC2DC\uC7A5")]) >= 0 ? findCol(h, [(k) => k.includes("\uC2DC\uC7A5")]) : 2;
  const iTheme =
    findCol(h, [(k) => k.includes("\uD14C\uB9C8") || k.includes("\uC5C5\uC885") || k.includes("\uC18D")]) >= 0
      ? findCol(h, [(k) => k.includes("\uD14C\uB9C8") || k.includes("\uC5C5\uC885") || k.includes("\uC18D")])
      : 3;
  const iClose = findCol(h, [(k) => k.includes("\uC885\uAC00")]) >= 0 ? findCol(h, [(k) => k.includes("\uC885\uAC00")]) : 4;
  const iChg = findCol(h, [(k) => k === "\uB4F1\uB77D" || k.includes("\uB300\uBE44") || k.includes("\uC804\uC77C\uBE44")]) >= 0
    ? findCol(h, [(k) => k === "\uB4F1\uB77D" || k.includes("\uB300\uBE44")])
    : 5;
  const iPct = findCol(h, [(k) => k.includes("\uB4F1\uB77D\uB960")]) >= 0 ? findCol(h, [(k) => k.includes("\uB4F1\uB77D\uB960")]) : 6;
  const iOpen = findCol(h, [(k) => k.includes("\uC2DC\uAC00")]) >= 0 ? findCol(h, [(k) => k.includes("\uC2DC\uAC00")]) : 7;
  const iHigh = findCol(h, [(k) => k.includes("\uACE0\uAC00")]) >= 0 ? findCol(h, [(k) => k.includes("\uACE0\uAC00")]) : 8;
  const iLow = findCol(h, [(k) => k.includes("\uC800\uAC00")]) >= 0 ? findCol(h, [(k) => k.includes("\uC800\uAC00")]) : 9;
  const iVol = findCol(h, [(k) => k.includes("\uAC70\uB798\uB7C9") && !k.includes("\uB9E4\uB3C4")]) >= 0
    ? findCol(h, [(k) => k.includes("\uAC70\uB798\uB7C9") && !k.includes("\uB9E4\uB3C4")])
    : 10;
  const iVal = findCol(h, [(k) => k.includes("\uAC70\uB798\uB300\uAE08") && !k.includes("\uB9E4\uB3C4")]) >= 0
    ? findCol(h, [(k) => k.includes("\uAC70\uB798\uB300\uAE08") && !k.includes("\uB9E4\uB3C4")])
    : 11;

  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;
    const code = normCode(row[iCode]);
    if (!code) continue;
    out.push({
      code,
      name: normStr(row[iName]),
      market: normStr(row[iMkt]),
      theme: normStr(row[iTheme]),
      close: parseNum(row[iClose]),
      chg: parseNum(row[iChg]),
      chgPct: parseNum(row[iPct]),
      open: parseNum(row[iOpen]),
      high: parseNum(row[iHigh]),
      low: parseNum(row[iLow]),
      volume: parseNum(row[iVol]),
      tradeValue: parseNum(row[iVal]),
    });
  }
  return out;
}

function parseFlowSheet(rows) {
  if (!rows.length) return [];
  const h = buildHeaderMap(rows[0]);
  const iCode = findCol(h, [(k) => k.includes("\uC885\uBAA9\uCF54\uB4DC")]) >= 0 ? findCol(h, [(k) => k.includes("\uC885\uBAA9\uCF54\uB4DC")]) : 0;
  const iName = findCol(h, [(k) => k.includes("\uC885\uBAA9\uBA85")]) >= 0 ? findCol(h, [(k) => k.includes("\uC885\uBAA9\uBA85")]) : 1;
  const iVn = findCol(h, [(k) => k.includes("\uAC70\uB798\uB7C9") && k.includes("\uC21C\uB9E4\uC218")]);
  const iValn = findCol(h, [(k) => k.includes("\uAC70\uB798\uB300\uAE08") && k.includes("\uC21C\uB9E4\uC218")]);
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;
    const code = normCode(row[iCode]);
    if (!code) continue;
    out.push({
      code,
      name: normStr(row[iName]),
      netVol: iVn >= 0 ? parseNum(row[iVn]) : NaN,
      netVal: iValn >= 0 ? parseNum(row[iValn]) : NaN,
    });
  }
  return out;
}

function readWorkbookDates(wb) {
  const dates = [];
  for (const sn of wb.SheetNames) {
    const iso = sheetNameToIso(sn);
    if (iso) dates.push(iso);
  }
  dates.sort();
  return dates;
}

function sheetRows(wb, iso) {
  const sn = iso.replace(/-/g, ".");
  if (!wb.Sheets[sn]) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
}

async function fileToWb(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: "array" });
}

function mergeData(priceWb, instWb, forWb) {
  const d1 = new Set(readWorkbookDates(priceWb));
  const d2 = new Set(readWorkbookDates(instWb));
  const d3 = new Set(readWorkbookDates(forWb));
  const common = [...d1].filter((d) => d2.has(d) && d3.has(d));
  common.sort();

  const byDate = new Map();
  const series = new Map();

  for (const iso of common) {
    const prow = sheetRows(priceWb, iso);
    const irow = sheetRows(instWb, iso);
    const frow = sheetRows(forWb, iso);
    const plist = parsePriceSheet(prow);
    const ilist = parseFlowSheet(irow);
    const flist = parseFlowSheet(frow);
    const imap = new Map(ilist.map((x) => [x.code, x]));
    const fmap = new Map(flist.map((x) => [x.code, x]));

    const merged = [];
    for (const p of plist) {
      const i = imap.get(p.code) || {};
      const f = fmap.get(p.code) || {};
      const row = {
        ...p,
        instNetVol: i.netVol,
        instNetVal: i.netVal,
        forNetVol: f.netVol,
        forNetVal: f.netVal,
      };
      merged.push(row);
      if (!series.has(p.code)) series.set(p.code, []);
      series.get(p.code).push({ date: iso, ...row });
    }
    byDate.set(iso, merged);
  }

  for (const arr of series.values()) {
    arr.sort((a, b) => (a.date < b.date ? -1 : 1));
  }

  return { sortedDates: common, byDate, series };
}

function setStatus(msg, isErr) {
  const el = document.getElementById("loadStatus");
  el.textContent = msg;
  el.classList.toggle("err", !!isErr);
}

let files = { price: null, inst: null, foreign: null };

async function tryLoadAll() {
  if (!files.price || !files.inst || !files.foreign) {
    setStatus(I18N.loadNeed, true);
    return;
  }
  try {
    const [pw, iw, fw] = await Promise.all([fileToWb(files.price), fileToWb(files.inst), fileToWb(files.foreign)]);
    state = mergeData(pw, iw, fw);
    if (!state.sortedDates.length) {
      setStatus("\uACF5\uD1B5 \uC77C\uC790\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4. \uC138 \uD30C\uC77C\uC758 \uC2DC\uD2B8\uBA85(\uC77C\uC790)\uAC00 \uC77C\uCE58\uD558\uB294\uC9C0 \uD655\uC778\uD558\uC138\uC694.", true);
      return;
    }
    setStatus(I18N.loadOk(state.sortedDates.length), false);
    fillDateSelects();
    buildStockList();
    runRanking();
  } catch (e) {
    setStatus(I18N.errParse(e.message || String(e)), true);
  }
}

function fillDateSelects() {
  const ids = ["anchorDate", "alertDate", "repStart", "repEnd"];
  for (const id of ids) {
    const sel = document.getElementById(id);
    sel.innerHTML = "";
    for (const d of state.sortedDates) {
      const o = document.createElement("option");
      o.value = d;
      o.textContent = d;
      sel.appendChild(o);
    }
    if (id === "anchorDate" || id === "alertDate") {
      sel.value = state.sortedDates[state.sortedDates.length - 1];
    }
    if (id === "repStart") sel.value = state.sortedDates[0];
    if (id === "repEnd") sel.value = state.sortedDates[state.sortedDates.length - 1];
  }
}

function buildStockList() {
  const dl = document.getElementById("stockList");
  dl.innerHTML = "";
  const last = state.sortedDates[state.sortedDates.length - 1];
  const rows = state.byDate.get(last) || [];
  const seen = new Set();
  for (const r of rows) {
    if (seen.has(r.code)) continue;
    seen.add(r.code);
    const opt = document.createElement("option");
    opt.value = `${r.code} ${r.name}`;
    dl.appendChild(opt);
  }
}

function tradingDaysEnding(endIso, maxDays) {
  const idx = state.sortedDates.indexOf(endIso);
  if (idx < 0) return [];
  const out = [];
  for (let i = idx; i >= 0 && out.length < maxDays; i--) {
    out.push(state.sortedDates[i]);
  }
  out.reverse();
  return out;
}

function horizonTradingCount(h) {
  if (h === "1d") return 1;
  if (h === "1w") return 5;
  if (h === "1m") return 22;
  return 1;
}

function periodReturn(code, endIso, h) {
  const days = tradingDaysEnding(endIso, horizonTradingCount(h));
  if (!days.length) return { ret: NaN, from: null, to: null, days: [] };
  const ser = state.series.get(code);
  if (!ser || !ser.length) return { ret: NaN, from: null, to: null, days: days };
  const map = new Map(ser.map((x) => [x.date, x]));
  const last = map.get(days[days.length - 1]);
  if (h === "1d") {
    const r = last && Number.isFinite(last.chgPct) ? last.chgPct : NaN;
    return { ret: r, from: days[days.length - 1], to: days[days.length - 1], days };
  }
  const first = map.get(days[0]);
  if (!first || !last || !Number.isFinite(first.close) || !Number.isFinite(last.close) || first.close === 0) {
    return { ret: NaN, from: days[0], to: days[days.length - 1], days };
  }
  const ret = ((last.close - first.close) / first.close) * 100;
  return { ret, from: days[0], to: days[days.length - 1], days };
}

function filterMarket(rows) {
  const m = document.getElementById("filterMarket").value;
  if (!m) return rows;
  return rows.filter((r) => normStr(r.market).toUpperCase() === m);
}

function runRanking() {
  const end = document.getElementById("anchorDate").value;
  const h = document.getElementById("horizon").value;
  const topN = Math.max(5, Math.min(200, parseInt(document.getElementById("topN").value || "30", 10)));
  const base = state.byDate.get(end) || [];
  const rows = filterMarket(base).filter((r) => Number.isFinite(r.close));

  const scored = rows.map((r) => {
    const pr = periodReturn(r.code, end, h);
    const avgTv =
      pr.days.length > 0
        ? pr.days.reduce((s, d) => {
            const row = (state.series.get(r.code) || []).find((x) => x.date === d);
            return s + (row && Number.isFinite(row.tradeValue) ? row.tradeValue : 0);
          }, 0) / pr.days.length
        : 0;
    const rets = pr.days
      .map((d) => {
        const row = (state.series.get(r.code) || []).find((x) => x.date === d);
        return row && Number.isFinite(row.chgPct) ? row.chgPct : null;
      })
      .filter((x) => x != null);
    const mean = rets.length ? rets.reduce((a, b) => a + b, 0) / rets.length : 0;
    const varc = rets.length      ? rets.reduce((s, x) => s + (x - mean) * (x - mean), 0) / rets.length
      : 0;
    const std = Math.sqrt(varc);
    const brokerScore = Math.log1p(Math.max(0, avgTv) / 1e9) * (1 + std);
    return { ...r, periodRet: pr.ret, brokerScore, std, avgTv };
  });

  const valid = scored.filter((x) => Number.isFinite(x.periodRet));
  const up = [...valid].sort((a, b) => b.periodRet - a.periodRet).slice(0, topN);
  const down = [...valid].sort((a, b) => a.periodRet - b.periodRet).slice(0, topN);
  const brok = [...scored].filter((x) => Number.isFinite(x.brokerScore)).sort((a, b) => b.brokerScore - a.brokerScore).slice(0, topN);

  const hLabel =
    h === "1d" ? I18N.h1d : h === "1w" ? I18N.h1w : I18N.h1m;
  renderTable("tblUp", up, [
    { k: "code", t: "\uCF54\uB4DC" },
    { k: "name", t: "\uC885\uBAA9\uBA85" },
    { k: "market", t: "\uC2DC\uC7A5" },
    { k: "theme", t: "\uD14C\uB9C8" },
    { k: "periodRet", t: `\uC218\uC775\uB960 (${hLabel}) %`, fmt: (v) => v.toFixed(2) },
    { k: "close", t: `\uC885\uAC00 (${end})`, fmt: (v) => (Number.isFinite(v) ? v.toLocaleString() : "-") },
    { k: "volume", t: `\uAC70\uB798\uB7C9`, fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-") },
  ]);
  renderTable("tblDown", down, [
    { k: "code", t: "\uCF54\uB4DC" },
    { k: "name", t: "\uC885\uBAA9\uBA85" },
    { k: "market", t: "\uC2DC\uC7A5" },
    { k: "theme", t: "\uD14C\uB9C8" },
    { k: "periodRet", t: `\uC218\uC775\uB960 (${hLabel}) %`, fmt: (v) => v.toFixed(2) },
    { k: "close", t: `\uC885\uAC00 (${end})`, fmt: (v) => (Number.isFinite(v) ? v.toLocaleString() : "-") },
    { k: "volume", t: `\uAC70\uB798\uB7C9`, fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-") },
  ]);
  renderTable("tblBroker", brok, [
    { k: "code", t: "\uCF54\uB4DC" },
    { k: "name", t: "\uC885\uBAA9\uBA85" },
    { k: "market", t: "\uC2DC\uC7A5" },
    { k: "brokerScore", t: "\uC2A4\uCF54\uC5B4", fmt: (v) => v.toFixed(3) },
    { k: "std", t: "\uC77C\uAC04%\uC2A4\uD2B8", fmt: (v) => v.toFixed(2) },
    { k: "avgTv", t: "\uD3C9\uADE0\uB300\uAE08(\uAD6C\uAC04)", fmt: (v) => (v / 1e8).toFixed(1) + "\uC5B5" },
    { k: "close", t: "\uC885\uAC00", fmt: (v) => (Number.isFinite(v) ? v.toLocaleString() : "-") },
  ]);
}

function renderTable(tableId, rows, cols) {
  const tbl = document.getElementById(tableId);
  const thead = tbl.querySelector("thead");
  const tbody = tbl.querySelector("tbody");
  thead.innerHTML = `<tr>${cols.map((c) => `<th>${c.t}</th>`).join("")}</tr>`;
  tbody.innerHTML = rows
    .map((r) => {
      const tds = cols.map((c) => {
        const raw = r[c.k];
        const txt = c.fmt ? c.fmt(raw, r) : raw == null ? "" : String(raw);
        let cls = "";
        if (c.k === "periodRet" && Number.isFinite(raw)) {
          cls = raw >= 0 ? "num-up" : "num-down";
        }
        return `<td class="${cls}">${txt}</td>`;
      });
      return `<tr>${tds.join("")}</tr>`;
    })
    .join("");
}

function resolveCodeFromSearch() {
  const q = normStr(document.getElementById("stockSearch").value);
  if (!q) return null;
  const token = q.split(/\s+/)[0];
  const c = normCode(token);
  if (state.series.has(c)) return c;
  const last = state.sortedDates[state.sortedDates.length - 1];
  const rows = state.byDate.get(last) || [];
  const hit = rows.find((r) => r.name === q || `${r.code} ${r.name}` === q);
  return hit ? hit.code : null;
}

function renderCharts(code) {
  const ser = state.series.get(code) || [];
  if (!ser.length) return;
  const labels = ser.map((x) => x.date);
  const closes = ser.map((x) => x.close);
  const vols = ser.map((x) => x.volume);
  const inst = ser.map((x) => (Number.isFinite(x.instNetVol) ? x.instNetVol / 1000 : null));
  const fr = ser.map((x) => (Number.isFinite(x.forNetVol) ? x.forNetVol / 1000 : null));

  const ctx1 = document.getElementById("chartPrice");
  if (chartPrice) chartPrice.destroy();
  chartPrice = new Chart(ctx1, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "\uC885\uAC00",
          data: closes,
          borderColor: "#3d8bfd",
          tension: 0.15,
          yAxisID: "y",
        },
        {
          label: "\uAC70\uB798\uB7C9",
          data: vols,
          type: "bar",
          backgroundColor: "rgba(139,156,179,0.35)",
          yAxisID: "y1",
        },
      ],
    },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      stacked: false,
      plugins: { legend: { labels: { color: "#e8edf4" } } },
      scales: {
        x: { ticks: { color: "#8b9cb3", maxRotation: 45 } },
        y: { type: "linear", position: "left", ticks: { color: "#8b9cb3" } },
        y1: { type: "linear", position: "right", grid: { drawOnChartArea: false }, ticks: { color: "#8b9cb3" } },
      },
    },
  });

  const ctx2 = document.getElementById("chartFlow");
  if (chartFlow) chartFlow.destroy();
  chartFlow = new Chart(ctx2, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "\uAE30\uAD00 \uC21C\uB9E4\uC218 (\uCC9C \uB2E8\uC704)",
          data: inst,
          backgroundColor: "rgba(255,107,107,0.6)",
        },
        {
          label: "\uC678\uAD6D\uC778 \uC21C\uB9E4\uC218 (\uCC9C \uB2E8\uC704)",
          data: fr,
          backgroundColor: "rgba(77,171,247,0.6)",
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { labels: { color: "#e8edf4" } } },
      scales: {
        x: { ticks: { color: "#8b9cb3", maxRotation: 45 } },
        y: { ticks: { color: "#8b9cb3" } },
      },
    },
  });

  renderTable("tblSeries", ser, [
    { k: "date", t: "\uC77C\uC790" },
    { k: "close", t: "\uC885\uAC00", fmt: (v) => (Number.isFinite(v) ? v.toLocaleString() : "-") },
    { k: "chgPct", t: "%", fmt: (v) => (Number.isFinite(v) ? v.toFixed(2) : "-") },
    { k: "volume", t: "\uAC70\uB798\uB7C9", fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-") },
    {
      k: "instNetVol",
      t: "\uAE30\uAD00\uC21C\uB9E4\uC218",
      fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-"),
    },
    {
      k: "forNetVol",
      t: "\uC678\uAD6D\uC778\uC21C\uB9E4\uC218",
      fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-"),
    },
  ]);
}

function median(arr) {
  const a = [...arr].sort((x, y) => x - y);
  if (!a.length) return NaN;
  const m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
}

function runAlerts() {
  const end = document.getElementById("alertDate").value;
  const pctTh = parseFloat(document.getElementById("alertPct").value || "5");
  const volMult = parseFloat(document.getElementById("alertVolMult").value || "2");
  const needInst = document.getElementById("alertInst").checked;
  const needFor = document.getElementById("alertForeign").checked;

  const idx = state.sortedDates.indexOf(end);
  const hist = idx >= 0 ? state.sortedDates.slice(Math.max(0, idx - 19), idx + 1) : [];
  const rows = state.byDate.get(end) || [];

  const out = [];
  for (const r of rows) {
    const ser = state.series.get(r.code) || [];
    const byD = new Map(ser.map((x) => [x.date, x]));
    const cur = byD.get(end);
    if (!cur) continue;
    const chg = Number.isFinite(cur.chgPct) ? Math.abs(cur.chgPct) : 0;
    if (chg < pctTh) continue;
    const vols = hist.map((d) => byD.get(d)?.volume).filter((v) => Number.isFinite(v));
    const med = median(vols);
    if (!Number.isFinite(med) || med <= 0) continue;
    if (cur.volume < med * volMult) continue;
    if (needInst && !(Number.isFinite(cur.instNetVol) && cur.instNetVol > 0)) continue;
    if (needFor && !(Number.isFinite(cur.forNetVol) && cur.forNetVol > 0)) continue;
    out.push({
      ...cur,
      volMedian: med,
      volRatio: cur.volume / med,
    });
  }
  out.sort((a, b) => b.volRatio - a.volRatio);
  renderTable("tblAlerts", out, [
    { k: "code", t: "\uCF54\uB4DC" },
    { k: "name", t: "\uC885\uBAA9\uBA85" },
    { k: "chgPct", t: "%", fmt: (v) => (Number.isFinite(v) ? v.toFixed(2) : "-") },
    { k: "volume", t: "\uAC70\uB798\uB7C9", fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-") },
    {
      k: "volRatio",
      t: "\uAC70\uB798\uB7C9/\uC911\uC559",
      fmt: (v) => (Number.isFinite(v) ? v.toFixed(2) : "-"),
    },
    {
      k: "instNetVol",
      t: "\uAE30\uAD00\uC21C",
      fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-"),
    },
    {
      k: "forNetVol",
      t: "\uC678\uAD6D\uC778\uC21C",
      fmt: (v) => (Number.isFinite(v) ? Math.round(v).toLocaleString() : "-"),
    },
  ]);
}

function buildReport() {
  const s = document.getElementById("repStart").value;
  const e = document.getElementById("repEnd").value;
  const si = state.sortedDates.indexOf(s);
  const ei = state.sortedDates.indexOf(e);
  if (si < 0 || ei < 0 || si > ei) {
    document.getElementById("reportOut").textContent = "\uC77C\uC790 \uBC94\uC704\uAC00 \uC62C\uBC14\uB974\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4.";
    return;
  }
  const span = state.sortedDates.slice(si, ei + 1);
  const last = state.sortedDates[ei];
  const base = state.byDate.get(last) || [];

  const movers = base
    .map((r) => {
      const pr = periodReturn(r.code, last, "1w");
      return { ...r, wRet: pr.ret };
    })
    .filter((x) => Number.isFinite(x.wRet))
    .sort((a, b) => Math.abs(b.wRet) - Math.abs(a.wRet))
    .slice(0, 15);

  const flows = [...base]
    .filter((r) => Number.isFinite(r.instNetVal) && Number.isFinite(r.forNetVal))
    .map((r) => ({
      ...r,
      score: (r.instNetVal > 0 ? 1 : 0) + (r.forNetVal > 0 ? 1 : 0),
      sumNet: (r.instNetVal || 0) + (r.forNetVal || 0),
    }))
    .sort((a, b) => b.sumNet - a.sumNet)
    .slice(0, 15);

  const lines = [];
  lines.push(`=== \uC2DC\uC7A5 \uC694\uC57D \uB9AC\uD3EC\uD2B8 ===`);
  lines.push(`\uAE30\uAC04: ${s} ~ ${e} (\uACF5\uD1B5 \uC601\uC5C5\uC77C ${span.length}\uC77C)`);
  lines.push("");
  lines.push(`[\uCD5C\uC885\uC77C ${last} \uAE30\uC900] \uCD5C\uADFC 1\uC8FC \uC218\uC775\uB960 \uC5ED\uB7C9 \uC885\uBAA9 (15)`);
  for (const r of movers) {
    lines.push(`- ${r.code} ${r.name} | 1\uC8FC ${r.wRet.toFixed(2)}% | \uC885\uAC00 ${r.close}`);
  }
  lines.push("");
  lines.push(`[\uCD5C\uC885\uC77C ${last}] \uAE30\uAD00+\uC678\uAD6D\uC778 \uC21C\uB9E4\uC218 \uD569\uC0B0 \uC0C1\uC704 (15, \uB300\uAE08 \uAE30\uC900)`);
  for (const r of flows) {
    lines.push(
      `- ${r.code} ${r.name} | \uAE30\uAD00\uB300\uAE08\uC21C ${(r.instNetVal / 1e8).toFixed(1)}\uC5B5 | \uC678\uAD6D\uC778\uB300\uAE08\uC21C ${(r.forNetVal / 1e8).toFixed(1)}\uC5B5`
    );
  }
  lines.push("");
  lines.push(
    `\uCC38\uACE0: \uB274\uC2A4/\uACF5\uC2DC \uC5F0\uACB0 \uC694\uC57D\uC740 \uBC84\uC804 \uBC94\uC704 \uBC16. \uC218\uC790 \uB370\uC774\uD130\uB9CC\uC73C\uB85C \uC0AC\uC720\uB97C \uC0C1\uC0C1\uD558\uC5EC \uC791\uC131\uD558\uC138\uC694.`
  );

  document.getElementById("reportOut").textContent = lines.join("\n");
}

function wireUi() {
  document.getElementById("filePrice").addEventListener("change", (e) => {
    files.price = e.target.files[0] || null;
    tryLoadAll();
  });
  document.getElementById("fileInst").addEventListener("change", (e) => {
    files.inst = e.target.files[0] || null;
    tryLoadAll();
  });
  document.getElementById("fileForeign").addEventListener("change", (e) => {
    files.foreign = e.target.files[0] || null;
    tryLoadAll();
  });

  const dz = document.getElementById("dropZone");
  dz.addEventListener("dragover", (e) => {
    e.preventDefault();
    dz.classList.add("drag");
  });
  dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
  dz.addEventListener("drop", (e) => {
    e.preventDefault();
    dz.classList.remove("drag");
    const fl = [...e.dataTransfer.files].filter((f) => /\.xlsx?$/i.test(f.name));
    for (const f of fl) {
      const n = f.name.toLowerCase();
      if (n.includes("\uac00\uaca9") || n.includes("price") || n.includes("volume")) files.price = f;
      else if (n.includes("\uae30\uad00") || n.includes("inst")) files.inst = f;
      else if (n.includes("\uc678\uad6d") || n.includes("foreign")) files.foreign = f;
    }
    tryLoadAll();
  });

  document.querySelectorAll("#mainTabs .tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll("#mainTabs .tab").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const id = btn.getAttribute("data-tab");
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === `tab-${id}`));
    });
  });

  document.getElementById("btnRank").addEventListener("click", runRanking);
  document.getElementById("anchorDate").addEventListener("change", runRanking);
  document.getElementById("horizon").addEventListener("change", runRanking);
  document.getElementById("filterMarket").addEventListener("change", runRanking);

  document.getElementById("btnLoadStock").addEventListener("click", () => {
    const code = resolveCodeFromSearch();
    if (!code) {
      alert("\uC885\uBAA9\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.");
      return;
    }
    renderCharts(code);
  });

  document.getElementById("btnScan").addEventListener("click", runAlerts);
  document.getElementById("btnReport").addEventListener("click", buildReport);
  document.getElementById("btnCopy").addEventListener("click", async () => {
    const t = document.getElementById("reportOut").textContent;
    try {
      await navigator.clipboard.writeText(t);
           alert("\uBCF5\uC0AC\uB418\uC5C8\uC2B5\uB2C8\uB2E4.");
    } catch {
      alert("\uBCF5\uC0AC\uC5D0 \uC2E4\uD328\uD588\uC2B5\uB2C8\uB2E4.");
    }
  });
}

applyI18n();
wireUi();
