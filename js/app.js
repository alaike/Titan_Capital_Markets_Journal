/* Titan Capital Markets — trading performance desk */
(() => {
  "use strict";

  const PAIRS = [
    "AUDCHF","AUDJPY","AUDCAD","AUDUSD","AUDNZD",
    "EURGBP","EURCHF","EURJPY","EURCAD","EURUSD","EURNZD","EURAUD",
    "GBPCHF","GBPJPY","GBPCAD","GBPUSD","GBPNZD","GBPAUD",
    "NZDCHF","NZDJPY","NZDCAD","NZDUSD",
    "CADJPY","CADCHF","USDCHF","USDJPY","USDCAD",
    "XAUUSD","CHFJPY"
  ];
  const SETUPS = ["Safety trade", "H1 5050 Bounce", "H4 50505 Bounce", "Other"];
  const SESSIONS = ["Asian", "Tokyo", "London", "New York"];
  const WEEKDAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const WEEKDAYS_MON = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];

  const STORAGE_TRADES = "aurum.trades.v1";
  const STORAGE_SETTINGS = "aurum.settings.v1";

  const DEFAULT_SETTINGS = {
    traderName: "",
    traderId: "",
    accountName: "",
    initialCapital: 10000,
    currency: "USD",
    defaultRiskPct: 1,
    beThreshold: 0.5
  };

  const TITLES = {
    dashboard: "Performance dashboard",
    trade: "Trade ticket",
    journal: "Trade log",
    analytics: "Interactive analysis",
    settings: "Trader & account",
    guide: "System blueprint"
  };

  /* ───────── state ───────── */
  let trades = [];
  let settings = { ...DEFAULT_SETTINGS };
  let filters = { pair: "ALL", day: "ALL", session: "ALL", setup: "ALL", from: "", to: "" };
  let charts = {};
  let pendingConfirm = null;

  /* ───────── helpers ───────── */
  const $ = (id) => document.getElementById(id);
  const uid = () => "t_" + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  const clamp = (n, a, b) => Math.min(b, Math.max(a, n));
  const isNum = (v) => v !== "" && v !== null && v !== undefined && !Number.isNaN(Number(v));
  const n = (v, d = 0) => (isNum(v) ? Number(v) : d);
  const round = (v, d = 2) => {
    if (!Number.isFinite(v)) return v;
    const p = 10 ** d;
    return Math.round(v * p) / p;
  };

  function money(v, ccy = settings.currency) {
    if (v === null || v === undefined || !Number.isFinite(Number(v))) return "—";
    const x = Number(v);
    const abs = Math.abs(x);
    const digits = abs >= 1000 ? 2 : abs >= 1 ? 2 : 2;
    try {
      return new Intl.NumberFormat("en-US", {
        style: "currency",
        currency: ccy || "USD",
        minimumFractionDigits: digits,
        maximumFractionDigits: digits
      }).format(x);
    } catch {
      return `${x.toFixed(2)} ${ccy}`;
    }
  }

  function fmt(v, d = 2) {
    if (v === null || v === undefined || !Number.isFinite(Number(v))) return "—";
    return Number(v).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
  }

  function pipSize(pair) {
    if (pair === "XAUUSD") return 0.1;
    if (pair && pair.includes("JPY")) return 0.01;
    return 0.0001;
  }
  function contractSize(pair) {
    return pair === "XAUUSD" ? 100 : 100000;
  }
  function priceDecimals(pair) {
    if (pair === "XAUUSD") return 2;
    if (pair && pair.includes("JPY")) return 3;
    return 5;
  }
  function quoteCcy(pair) {
    if (!pair || pair.length < 6) return "";
    return pair === "XAUUSD" ? "USD" : pair.slice(-3);
  }
  function baseCcy(pair) {
    if (!pair || pair.length < 6) return "";
    return pair === "XAUUSD" ? "XAU" : pair.slice(0, 3);
  }

  function parseDT(date, time) {
    if (!date) return null;
    const t = time && time.length ? time : "00:00:00";
    const iso = t.length === 5 ? `${date}T${t}:00` : `${date}T${t}`;
    const d = new Date(iso);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function weekdayFromDate(dateStr) {
    if (!dateStr) return "";
    const d = new Date(dateStr + "T12:00:00");
    if (Number.isNaN(d.getTime())) return "";
    return WEEKDAYS[d.getDay()];
  }

  function suggestSession(timeStr) {
    if (!timeStr) return "";
    const [hh, mm] = timeStr.split(":").map(Number);
    const minutes = hh * 60 + (mm || 0);
    // Africa/Lagos WAT = UTC+1. Sessions in WAT:
    // Tokyo 01:00–10:00, London 08:00–17:00, New York 14:00–23:00, Asian ~23:00–08:00
    if (minutes >= 14 * 60 && minutes < 23 * 60) return "New York";
    if (minutes >= 8 * 60 && minutes < 17 * 60) return "London";
    if (minutes >= 1 * 60 && minutes < 10 * 60) return "Tokyo";
    return "Asian";
  }

  function durationParts(ms) {
    if (!Number.isFinite(ms) || ms < 0) return null;
    const totalMin = Math.round(ms / 60000);
    const days = Math.floor(totalMin / 1440);
    const hours = Math.floor((totalMin % 1440) / 60);
    const mins = totalMin % 60;
    return { days, hours, mins, totalMin };
  }
  function durationLabel(ms) {
    const p = durationParts(ms);
    if (!p) return "—";
    const bits = [];
    if (p.days) bits.push(p.days + "d");
    if (p.hours || p.days) bits.push(p.hours + "h");
    bits.push(p.mins + "m");
    return bits.join(" ");
  }

  function direction(position) {
    return position === "Sell" ? -1 : 1;
  }

  function estimatePnl(pair, position, entry, exit, lots) {
    if (!pair || !isNum(entry) || !isNum(exit) || !isNum(lots)) return { pnl: null, method: "none", quote: "" };
    const dir = direction(position);
    const raw = (Number(exit) - Number(entry)) * dir * contractSize(pair) * Number(lots);
    const quote = quoteCcy(pair);
    const base = baseCcy(pair);
    if (quote === "USD") {
      return { pnl: raw, method: "direct-usd", quote };
    }
    if (base === "USD") {
      return { pnl: Number(exit) ? raw / Number(exit) : null, method: "usd-base", quote };
    }
    const pips = ((Number(exit) - Number(entry)) * dir) / pipSize(pair);
    return { pnl: pips * 10 * Number(lots), method: "cross-approx-10", quote };
  }

  function compute(trade, beThreshold = settings.beThreshold) {
    const pair = trade.pair;
    const dir = direction(trade.position);
    const ps = pipSize(pair);
    const entry = n(trade.entry, NaN);
    const sl = n(trade.sl, NaN);
    const tp = n(trade.tp, NaN);
    const exit = n(trade.exit, NaN);
    const lots = n(trade.lots, NaN);
    const commission = n(trade.commission, 0);

    const slPips = Number.isFinite(entry) && Number.isFinite(sl) ? Math.abs(entry - sl) / ps : null;
    const tpPips = Number.isFinite(entry) && Number.isFinite(tp) && trade.tp !== "" && trade.tp !== null
      ? Math.abs(tp - entry) / ps : null;
    const resultPips = Number.isFinite(entry) && Number.isFinite(exit)
      ? ((exit - entry) * dir) / ps : null;

    const plannedRR = slPips ? (tpPips === null ? null : tpPips / slPips) : null;
    const realizedRR = slPips && resultPips !== null ? resultPips / slPips : null;

    const est = estimatePnl(pair, trade.position, entry, exit, lots);
    const hasOverride = trade.pnlOverride !== "" && trade.pnlOverride !== null && trade.pnlOverride !== undefined
      && Number.isFinite(Number(trade.pnlOverride));
    const gross = hasOverride ? Number(trade.pnlOverride) : est.pnl;
    const net = gross === null ? null : gross - commission;

    const entryDT = parseDT(trade.entryDate, trade.entryTime);
    const exitDT = parseDT(trade.exitDate, trade.exitTime);
    const durMs = entryDT && exitDT ? exitDT - entryDT : null;

    const closed = Number.isFinite(exit) && !!trade.exitDate;
    let result = "Open";
    if (closed && net !== null) {
      const be = Math.abs(n(beThreshold, 0.5));
      if (Math.abs(net) <= be) result = "BE";
      else if (net > 0) result = "Win";
      else result = "Loss";
    }

    const setup = trade.setup === "Other" && trade.setupCustom ? trade.setupCustom : trade.setup;
    const dow = weekdayFromDate(trade.entryDate);

    const slSideOk = Number.isFinite(entry) && Number.isFinite(sl)
      ? (trade.position === "Sell" ? sl > entry : sl < entry)
      : null;
    const tpSideOk = Number.isFinite(entry) && Number.isFinite(tp) && trade.tp !== "" && trade.tp !== null
      ? (trade.position === "Sell" ? tp < entry : tp > entry)
      : null;

    const riskFromStop = slPips !== null && Number.isFinite(lots)
      ? estimatePnl(pair, trade.position, entry, sl, lots).pnl
      : null;
    // risk from stop is the loss if SL is hit — take absolute
    const riskCalc = riskFromStop === null ? null : Math.abs(riskFromStop);

    return {
      ...trade,
      setupLabel: setup,
      dow,
      slPips,
      tpPips,
      resultPips,
      plannedRR,
      realizedRR,
      estPnl: est.pnl,
      estMethod: est.method,
      quote: est.quote,
      hasOverride,
      grossPnl: gross,
      netPnl: net,
      durationMs: durMs,
      duration: durationLabel(durMs),
      closed,
      result,
      slSideOk,
      tpSideOk,
      riskCalc,
      decimals: priceDecimals(pair)
    };
  }

  function enrichAll() {
    return trades.map((t) => compute(t));
  }

  function withBalances(rows) {
    const closed = rows.filter((t) => t.closed && t.netPnl !== null)
      .slice()
      .sort((a, b) => {
        const da = parseDT(a.exitDate, a.exitTime) || parseDT(a.entryDate, a.entryTime);
        const db = parseDT(b.exitDate, b.exitTime) || parseDT(b.entryDate, b.entryTime);
        const ta = da ? da.getTime() : 0;
        const tb = db ? db.getTime() : 0;
        if (ta !== tb) return ta - tb;
        return (a.sn || 0) - (b.sn || 0);
      });
    let bal = n(settings.initialCapital, 0);
    const map = new Map();
    const equity = [{ label: "Start", balance: bal, pnl: 0, cum: 0, id: "start" }];
    let peak = bal;
    let maxDd = 0;
    let cum = 0;
    closed.forEach((t) => {
      cum += t.netPnl;
      bal += t.netPnl;
      map.set(t.id, bal);
      if (bal > peak) peak = bal;
      const dd = peak - bal;
      if (dd > maxDd) maxDd = dd;
      equity.push({
        label: `#${t.sn} ${t.pair}`,
        balance: bal,
        pnl: t.netPnl,
        cum,
        id: t.id,
        date: t.exitDate || t.entryDate
      });
    });
    return { map, equity, maxDd, lastBalance: bal, closedCount: closed.length };
  }

  function applyFilters(rows) {
    return rows.filter((t) => {
      if (filters.pair !== "ALL" && t.pair !== filters.pair) return false;
      if (filters.day !== "ALL" && t.dow !== filters.day) return false;
      if (filters.session !== "ALL" && t.session !== filters.session) return false;
      if (filters.setup !== "ALL") {
        const label = t.setupLabel || t.setup;
        if (label !== filters.setup && t.setup !== filters.setup) return false;
      }
      if (filters.from && t.entryDate && t.entryDate < filters.from) return false;
      if (filters.to && t.entryDate && t.entryDate > filters.to) return false;
      return true;
    });
  }

  function stats(rows) {
    const closed = rows.filter((t) => t.closed && t.netPnl !== null);
    const open = rows.filter((t) => !t.closed);
    const wins = closed.filter((t) => t.result === "Win");
    const losses = closed.filter((t) => t.result === "Loss");
    const bes = closed.filter((t) => t.result === "BE");
    const pnls = closed.map((t) => t.netPnl);
    const totalPnl = pnls.reduce((a, b) => a + b, 0);
    const grossProfit = pnls.filter((x) => x > 0).reduce((a, b) => a + b, 0);
    const grossLossAbs = Math.abs(pnls.filter((x) => x < 0).reduce((a, b) => a + b, 0));
    let profitFactor = null;
    if (grossLossAbs === 0 && grossProfit > 0) profitFactor = Infinity;
    else if (grossLossAbs === 0 && grossProfit === 0) profitFactor = null;
    else if (grossLossAbs === 0) profitFactor = null;
    else profitFactor = grossProfit / grossLossAbs;

    const avgWin = wins.length ? wins.reduce((a, t) => a + t.netPnl, 0) / wins.length : null;
    const avgLoss = losses.length ? losses.reduce((a, t) => a + t.netPnl, 0) / losses.length : null;
    const largestWin = wins.length ? Math.max(...wins.map((t) => t.netPnl)) : null;
    const largestLoss = losses.length ? Math.min(...losses.map((t) => t.netPnl)) : null;
    const winRate = closed.length ? (wins.length / closed.length) * 100 : null;
    const winRateExcl = (wins.length + losses.length)
      ? (wins.length / (wins.length + losses.length)) * 100 : null;
    const expectancy = closed.length ? totalPnl / closed.length : null;
    const avgRR = (() => {
      const xs = closed.map((t) => t.realizedRR).filter((x) => Number.isFinite(x));
      return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null;
    })();

    return {
      closed, open, wins, losses, bes, totalPnl, grossProfit, grossLossAbs,
      profitFactor, avgWin, avgLoss, largestWin, largestLoss, winRate, winRateExcl,
      expectancy, avgRR, nClosed: closed.length, nOpen: open.length,
      nWins: wins.length, nLosses: losses.length, nBE: bes.length
    };
  }

  function groupBy(rows, keyFn) {
    const map = new Map();
    rows.forEach((t) => {
      const k = keyFn(t) || "—";
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(t);
    });
    return map;
  }

  function groupStats(rows, keyFn, order) {
    const g = groupBy(rows.filter((t) => t.closed && t.netPnl !== null), keyFn);
    const keys = order ? order.filter((k) => g.has(k)).concat([...g.keys()].filter((k) => !order.includes(k)))
      : [...g.keys()];
    return keys.map((k) => {
      const s = stats(g.get(k));
      return { key: k, ...s };
    });
  }

  function bestWorst(groups) {
    if (!groups.length) return { best: null, worst: null };
    const withTrades = groups.filter((g) => g.nClosed);
    if (!withTrades.length) return { best: null, worst: null };
    const best = withTrades.reduce((a, b) => (b.totalPnl > a.totalPnl ? b : a));
    const worst = withTrades.reduce((a, b) => (b.totalPnl < a.totalPnl ? b : a));
    return { best, worst };
  }

  /* ───────── storage ───────── */
  function load() {
    try {
      const t = JSON.parse(localStorage.getItem(STORAGE_TRADES) || "[]");
      trades = Array.isArray(t) ? t : [];
    } catch { trades = []; }
    try {
      const s = JSON.parse(localStorage.getItem(STORAGE_SETTINGS) || "null");
      settings = { ...DEFAULT_SETTINGS, ...(s || {}) };
    } catch { settings = { ...DEFAULT_SETTINGS }; }
    reindex();
  }
  function persistTrades() {
    localStorage.setItem(STORAGE_TRADES, JSON.stringify(trades));
  }
  function persistSettings() {
    localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(settings));
  }
  function reindex() {
    const sorted = trades.slice().sort((a, b) => {
      const da = parseDT(a.entryDate, a.entryTime);
      const db = parseDT(b.entryDate, b.entryTime);
      const ta = da ? da.getTime() : 0;
      const tb = db ? db.getTime() : 0;
      if (ta !== tb) return ta - tb;
      return String(a.id).localeCompare(String(b.id));
    });
    sorted.forEach((t, i) => { t.sn = i + 1; });
    trades = sorted;
  }

  /* ───────── toast / modal ───────── */
  function toast(msg, kind = "ok") {
    const el = document.createElement("div");
    el.className = "toast " + kind;
    el.textContent = msg;
    $("toasts").appendChild(el);
    setTimeout(() => el.remove(), 3200);
  }
  function confirmModal(title, body) {
    return new Promise((resolve) => {
      $("modalTitle").textContent = title;
      $("modalBody").textContent = body;
      $("modal").classList.add("open");
      pendingConfirm = resolve;
    });
  }
  $("modalOk").addEventListener("click", () => {
    $("modal").classList.remove("open");
    if (pendingConfirm) pendingConfirm(true);
    pendingConfirm = null;
  });
  $("modalCancel").addEventListener("click", () => {
    $("modal").classList.remove("open");
    if (pendingConfirm) pendingConfirm(false);
    pendingConfirm = null;
  });

  /* ───────── navigation ───────── */
  function showView(name, extra) {
    document.querySelectorAll(".view").forEach((v) => v.classList.remove("active"));
    document.querySelectorAll(".nav-btn").forEach((b) => {
      b.classList.toggle("active", b.dataset.view === name);
    });
    const view = $("view-" + name);
    if (view) view.classList.add("active");
    $("pageTitle").textContent = TITLES[name] || name;
    if (name === "dashboard") renderDashboard();
    if (name === "analytics") renderAnalytics();
    if (name === "journal") renderJournal();
    if (name === "settings") renderSettings();
    if (name === "trade") {
      if (extra && extra.id) fillForm(trades.find((t) => t.id === extra.id));
      else if (!extra || !extra.keep) resetForm(false);
      updateTicket();
    }
    updateSide();
  }

  document.getElementById("nav").addEventListener("click", (e) => {
    const btn = e.target.closest(".nav-btn");
    if (btn) showView(btn.dataset.view);
  });
  $("quickNewBtn").addEventListener("click", () => showView("trade"));
  $("identityChip").addEventListener("click", () => showView("settings"));
  $("traderCard").addEventListener("click", () => showView("settings"));

  /* ───────── selects ───────── */
  function fillSelect(el, items, withAll, withBlank) {
    const cur = el.value;
    el.innerHTML = "";
    if (withAll) {
      const o = document.createElement("option");
      o.value = "ALL"; o.textContent = "All";
      el.appendChild(o);
    }
    if (withBlank) {
      const o = document.createElement("option");
      o.value = ""; o.textContent = "—";
      el.appendChild(o);
    }
    items.forEach((x) => {
      const o = document.createElement("option");
      o.value = x; o.textContent = x;
      el.appendChild(o);
    });
    if ([...el.options].some((o) => o.value === cur)) el.value = cur;
  }

  function usedSetups() {
    const extra = new Set(SETUPS);
    enrichAll().forEach((t) => { if (t.setupLabel) extra.add(t.setupLabel); });
    return [...extra];
  }

  function initFormSelects() {
    fillSelect($("f-pair"), [...PAIRS, "OTHER"]);
    fillSelect($("f-setup"), SETUPS);
  }

  /* ───────── form ───────── */
  function formValues() {
    return {
      id: $("f-id").value || uid(),
      pair: $("f-pair").value,
      position: $("f-position").value,
      setup: $("f-setup").value,
      setupCustom: $("f-setupCustom").value.trim(),
      entryDate: $("f-entryDate").value,
      entryTime: $("f-entryTime").value,
      session: $("f-session").value,
      lots: $("f-lots").value,
      riskAmt: $("f-riskAmt").value,
      riskPct: $("f-riskPct").value,
      entry: $("f-entry").value,
      sl: $("f-sl").value,
      tp: $("f-tp").value,
      exit: $("f-exit").value,
      exitDate: $("f-exitDate").value,
      exitTime: $("f-exitTime").value,
      exitReason: $("f-exitReason").value,
      commission: $("f-commission").value,
      pnlOverride: $("f-pnlOverride").value,
      plan: $("f-plan").value,
      market: $("f-market").value,
      notes: $("f-notes").value,
      tf: $("f-tf").value,
      emotion: $("f-emotion").value,
      ticket: $("f-ticket").value.trim(),
      chartUrl: $("f-chartUrl").value.trim(),
      mae: $("f-mae").value,
      mfe: $("f-mfe").value
    };
  }

  function fillForm(t) {
    if (!t) return;
    $("f-id").value = t.id;
    $("f-pair").value = PAIRS.includes(t.pair) ? t.pair : "OTHER";
    $("f-position").value = t.position || "Buy";
    $("f-setup").value = SETUPS.includes(t.setup) ? t.setup : "Other";
    $("f-setupCustom").value = t.setupCustom || (!SETUPS.includes(t.setup) ? t.setup : "");
    $("f-entryDate").value = t.entryDate || "";
    $("f-entryTime").value = t.entryTime || "";
    $("f-session").value = t.session || "London";
    $("f-lots").value = t.lots ?? "";
    $("f-riskAmt").value = t.riskAmt ?? "";
    $("f-riskPct").value = t.riskPct ?? "";
    $("f-entry").value = t.entry ?? "";
    $("f-sl").value = t.sl ?? "";
    $("f-tp").value = t.tp ?? "";
    $("f-exit").value = t.exit ?? "";
    $("f-exitDate").value = t.exitDate || "";
    $("f-exitTime").value = t.exitTime || "";
    $("f-exitReason").value = t.exitReason || "";
    $("f-commission").value = t.commission ?? 0;
    $("f-pnlOverride").value = t.pnlOverride ?? "";
    $("f-plan").value = t.plan || "";
    $("f-market").value = t.market || "";
    $("f-notes").value = t.notes || "";
    $("f-tf").value = t.tf || "";
    $("f-emotion").value = t.emotion || "";
    $("f-ticket").value = t.ticket || "";
    $("f-chartUrl").value = t.chartUrl || "";
    $("f-mae").value = t.mae ?? "";
    $("f-mfe").value = t.mfe ?? "";
    $("saveBtn").textContent = "Update trade";
    $("deleteBtn").hidden = false;
    $("f-dow").value = weekdayFromDate(t.entryDate);
    setPriceSteps();
    updateTicket();
  }

  function resetForm(keepDate = true) {
    $("tradeForm").reset();
    $("f-id").value = "";
    $("f-commission").value = 0;
    $("saveBtn").textContent = "Save trade";
    $("deleteBtn").hidden = true;
    if (keepDate) {
      const now = new Date();
      const pad = (x) => String(x).padStart(2, "0");
      $("f-entryDate").value = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
      $("f-entryTime").value = `${pad(now.getHours())}:${pad(now.getMinutes())}`;
      $("f-dow").value = weekdayFromDate($("f-entryDate").value);
      const sug = suggestSession($("f-entryTime").value);
      if (sug) $("f-session").value = sug;
    }
    if (settings.defaultRiskPct) $("f-riskPct").value = settings.defaultRiskPct;
    setPriceSteps();
    updateTicket();
  }

  function setPriceSteps() {
    const pair = $("f-pair").value;
    const step = pipSize(pair) / 10;
    ["f-entry", "f-sl", "f-tp", "f-exit"].forEach((id) => {
      $(id).step = String(step);
    });
  }

  function qcList(c) {
    const items = [];
    if (!c.pair || c.pair === "OTHER") items.push({ k: "warn", t: "Select a listed pair, or type a custom setup/pair in notes." });
    if (!c.entryDate || !c.entryTime) items.push({ k: "bad", t: "Entry date and time are required." });
    if (!isNum(c.entry) || Number(c.entry) <= 0) items.push({ k: "bad", t: "Entry price must be > 0." });
    if (!isNum(c.sl) || Number(c.sl) <= 0) items.push({ k: "bad", t: "Stop loss is required." });
    if (!isNum(c.lots) || Number(c.lots) <= 0) items.push({ k: "bad", t: "Lot size must be > 0." });
    if (c.slSideOk === false) items.push({ k: "bad", t: "Stop is on the wrong side of entry for this position." });
    else if (c.slSideOk) items.push({ k: "ok", t: "Stop is on the protective side." });
    if (c.tpSideOk === false) items.push({ k: "warn", t: "Take profit is on the unusual side of entry." });
    if (c.closed && c.durationMs !== null && c.durationMs < 0) items.push({ k: "bad", t: "Exit is earlier than entry." });
    if (c.riskCalc && isNum(c.riskAmt) && Number(c.riskAmt) > 0) {
      const drift = Math.abs(c.riskCalc - Number(c.riskAmt)) / Number(c.riskAmt);
      if (drift > 0.15) items.push({ k: "warn", t: `Risk from stop ≈ ${money(c.riskCalc)} vs typed risk ${money(Number(c.riskAmt))} (>15% apart).` });
      else items.push({ k: "ok", t: `Stop risk ≈ ${money(c.riskCalc)} matches typed risk.` });
    } else if (c.riskCalc) {
      items.push({ k: "ok", t: `Calculated risk at stop ≈ ${money(c.riskCalc)}.` });
    }
    if (c.estMethod === "cross-approx-10" && c.closed && !c.hasOverride) {
      items.push({ k: "warn", t: "Cross pair: PnL is estimated at ~$10/pip/lot. Enter broker PnL for accuracy." });
    }
    if (c.hasOverride) items.push({ k: "ok", t: "Using broker PnL override." });
    if (c.closed && c.durationMs !== null && c.durationMs >= 0) items.push({ k: "ok", t: "Duration " + c.duration });
    return items;
  }

  function updateTicket() {
    const raw = formValues();
    const c = compute(raw);
    $("f-dow").value = c.dow || "";
    $("tk-pair").textContent = c.pair && c.pair !== "OTHER" ? c.pair : "PAIR";
    $("tk-sn").textContent = c.id && trades.some((t) => t.id === c.id)
      ? "S/N " + (trades.find((t) => t.id === c.id).sn)
      : "New ticket";
    $("tk-side").innerHTML = `<span class="badge ${c.position === "Sell" ? "sell" : "buy"}">${c.position === "Sell" ? "Sell / Short" : "Buy / Long"}</span>`;

    const d = c.decimals;
    const row = (k, v) => `<div class="ticket-row"><span>${k}</span><b>${v}</b></div>`;
    const pips = (x) => (x === null || !Number.isFinite(x) ? "—" : fmt(x, 1));
    const rr = (x) => (x === null || !Number.isFinite(x) ? "—" : (x >= 0 ? "1 : " + fmt(x, 2) : fmt(x, 2)));

    let stamp = "Draft";
    let stampCls = "";
    if (c.closed) {
      stamp = c.result === "Win" ? "Win" : c.result === "Loss" ? "Loss" : "Breakeven";
      stampCls = c.result === "Win" ? "win" : c.result === "Loss" ? "loss" : "be";
    } else if (isNum(c.entry)) {
      stamp = "Open";
      stampCls = "open";
    }

    const riskShare = c.slPips && c.tpPips ? c.slPips / (c.slPips + c.tpPips) : 0.5;
    const rewardShare = 1 - riskShare;

    const qc = qcList(c).map((i) => `<div class="qc-item ${i.k}">${i.t}</div>`).join("");

    $("ticketBody").innerHTML = `
      ${row("Day", c.dow || "—")}
      ${row("Session", c.session || "—")}
      ${row("Setup", c.setupLabel || "—")}
      ${row("Lots", isNum(c.lots) ? fmt(Number(c.lots), 2) : "—")}
      ${row("Entry", isNum(c.entry) ? Number(c.entry).toFixed(d) : "—")}
      ${row("Stop · pips", `${isNum(c.sl) ? Number(c.sl).toFixed(d) : "—"} · ${pips(c.slPips)}`)}
      ${row("Target · pips", `${isNum(c.tp) && c.tp !== "" ? Number(c.tp).toFixed(d) : "—"} · ${pips(c.tpPips)}`)}
      ${row("Planned RR", rr(c.plannedRR))}
      <div class="rr-bar" title="Risk vs reward"><i class="risk" style="width:${(riskShare * 100).toFixed(1)}%"></i><i class="reward" style="width:${(rewardShare * 100).toFixed(1)}%"></i></div>
      ${row("Exit", isNum(c.exit) ? Number(c.exit).toFixed(d) : "—")}
      ${row("Result pips", pips(c.resultPips))}
      ${row("Realized RR", rr(c.realizedRR))}
      ${row("Est. PnL", c.estPnl === null ? "—" : money(c.estPnl))}
      ${row("Commission", money(n(c.commission, 0)))}
      ${row("Net PnL", c.netPnl === null ? "—" : money(c.netPnl))}
      ${row("Duration", c.duration)}
      <div class="ticket-stamp ${stampCls}">${stamp}</div>
      <div class="qc-list">${qc}</div>
    `;
  }

  ["f-pair","f-position","f-setup","f-entryDate","f-entryTime","f-session","f-lots","f-riskAmt",
   "f-entry","f-sl","f-tp","f-exit","f-exitDate","f-exitTime","f-commission","f-pnlOverride",
   "f-setupCustom"].forEach((id) => {
    const el = $(id);
    el.addEventListener("input", () => {
      if (id === "f-pair") setPriceSteps();
      if (id === "f-entryDate") $("f-dow").value = weekdayFromDate($("f-entryDate").value);
      updateTicket();
    });
    el.addEventListener("change", updateTicket);
  });

  $("advToggle").addEventListener("click", () => {
    const body = $("advBody");
    const open = body.classList.toggle("open");
    $("advToggle").textContent = open ? "Hide advanced fields" : "Show advanced fields";
  });

  $("resetBtn").addEventListener("click", () => resetForm(true));

  $("tradeForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const raw = formValues();
    const c = compute(raw);
    const hard = qcList(c).filter((i) => i.k === "bad");
    if (hard.length) {
      toast(hard[0].t, "bad");
      return;
    }
    const existing = trades.findIndex((t) => t.id === raw.id);
    const dup = trades.find((t) => t.id !== raw.id && t.pair === raw.pair && t.entryDate === raw.entryDate
      && t.entryTime === raw.entryTime && String(t.entry) === String(raw.entry));
    const save = () => {
      if (existing >= 0) trades[existing] = { ...trades[existing], ...raw };
      else trades.push(raw);
      reindex();
      persistTrades();
      toast(existing >= 0 ? "Trade updated" : "Trade saved");
      showView("journal");
    };
    if (dup) {
      confirmModal("Possible duplicate",
        `A ${dup.pair} trade already exists at ${dup.entryDate} ${dup.entryTime} (S/N ${dup.sn}). Save anyway?`
      ).then((ok) => { if (ok) save(); });
    } else save();
  });

  $("deleteBtn").addEventListener("click", async () => {
    const id = $("f-id").value;
    if (!id) return;
    const ok = await confirmModal("Delete trade", "This cannot be undone.");
    if (!ok) return;
    trades = trades.filter((t) => t.id !== id);
    reindex();
    persistTrades();
    toast("Trade deleted");
    resetForm(true);
    showView("journal");
  });

  /* ───────── filters UI ───────── */
  function filterBarHTML(prefix) {
    const setups = usedSetups().filter((s) => s !== "Other");
    const opt = (arr) => arr.map((x) => `<option value="${x}">${x}</option>`).join("");
    return `
      <div class="filter-bar" data-prefix="${prefix}">
        <div class="filter-group">
          <label>Currency pair</label>
          <select data-f="pair">
            <option value="ALL">All pairs</option>
            ${opt(PAIRS)}
          </select>
        </div>
        <div class="filter-group">
          <label>Day of week</label>
          <select data-f="day">
            <option value="ALL">All days</option>
            ${opt(WEEKDAYS_MON)}
          </select>
        </div>
        <div class="filter-group">
          <label>Session</label>
          <select data-f="session">
            <option value="ALL">All sessions</option>
            ${opt(SESSIONS)}
          </select>
        </div>
        <div class="filter-group">
          <label>Setup</label>
          <select data-f="setup">
            <option value="ALL">All setups</option>
            ${opt(setups)}
          </select>
        </div>
        <div class="filter-group">
          <label>From</label>
          <input type="date" data-f="from" />
        </div>
        <div class="filter-group">
          <label>To</label>
          <input type="date" data-f="to" />
        </div>
        <div class="filter-actions">
          <button type="button" class="btn btn-ghost btn-sm" data-reset>Reset</button>
        </div>
      </div>`;
  }

  function bindFilterBar(container) {
    container.querySelectorAll("[data-f]").forEach((el) => {
      el.value = filters[el.dataset.f];
      el.addEventListener("change", () => {
        filters[el.dataset.f] = el.value;
        syncFilterBars();
        renderDashboard();
        renderAnalytics();
      });
    });
    const reset = container.querySelector("[data-reset]");
    if (reset) reset.addEventListener("click", () => {
      filters = { pair: "ALL", day: "ALL", session: "ALL", setup: "ALL", from: "", to: "" };
      syncFilterBars();
      renderDashboard();
      renderAnalytics();
    });
  }
  function syncFilterBars() {
    document.querySelectorAll(".filter-bar [data-f]").forEach((el) => {
      el.value = filters[el.dataset.f];
    });
  }

  /* ───────── charts ───────── */
  const gold = "#c9a84c";
  const up = "#3ecf8e";
  const down = "#e85d6c";
  const muted = "#6d7a8c";
  const info = "#7ba3e3";

  Chart.defaults.color = "#9aa6b8";
  Chart.defaults.borderColor = "rgba(230,236,247,0.07)";
  Chart.defaults.font.family = "Manrope, sans-serif";
  Chart.defaults.font.size = 11;
  Chart.defaults.plugins.legend.labels.boxWidth = 10;
  Chart.defaults.plugins.legend.labels.boxHeight = 10;
  Chart.defaults.plugins.tooltip.backgroundColor = "#161d2a";
  Chart.defaults.plugins.tooltip.borderColor = "rgba(201,168,76,0.3)";
  Chart.defaults.plugins.tooltip.borderWidth = 1;
  Chart.defaults.plugins.tooltip.titleColor = "#e7d392";
  Chart.defaults.plugins.tooltip.bodyColor = "#e9eef6";

  function destroyChart(id) {
    if (charts[id]) {
      charts[id].destroy();
      delete charts[id];
    }
  }

  function emptyOrChart(canvasId, has, build) {
    const canvas = $(canvasId);
    if (!canvas) return;
    const wrap = canvas.parentElement;
    destroyChart(canvasId);
    const old = wrap.querySelector(".empty-chart");
    if (old) old.remove();
    if (!has) {
      canvas.style.display = "none";
      const d = document.createElement("div");
      d.className = "empty-chart";
      d.textContent = "Log closed trades to populate this chart";
      wrap.appendChild(d);
      return;
    }
    canvas.style.display = "block";
    charts[canvasId] = build(canvas);
  }

  function barColors(values) {
    return values.map((v) => (v > 0 ? "rgba(62,207,142,0.85)" : v < 0 ? "rgba(232,93,108,0.85)" : "rgba(139,151,168,0.5)"));
  }

  function lineChart(canvas, labels, data, color, fill) {
    return new Chart(canvas, {
      type: "line",
      data: {
        labels,
        datasets: [{
          data,
          borderColor: color,
          backgroundColor: fill || "transparent",
          fill: !!fill,
          tension: 0.25,
          pointRadius: labels.length > 40 ? 0 : 3,
          pointHoverRadius: 5,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 10 }, grid: { display: false } },
          y: { grid: { color: "rgba(230,236,247,0.05)" } }
        }
      }
    });
  }

  function barChart(canvas, labels, data, extra = {}) {
    return new Chart(canvas, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: extra.colors || barColors(data),
          borderRadius: 4,
          maxBarThickness: 42
        }]
      },
      options: {
        indexAxis: extra.horizontal ? "y" : "x",
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { display: extra.horizontal }, ticks: extra.horizontal ? {} : { maxRotation: 40, autoSkip: true } },
          y: { grid: { color: "rgba(230,236,247,0.05)" }, beginAtZero: extra.beginAtZero !== false }
        }
      }
    });
  }

  /* ───────── KPI render ───────── */
  function kpiHTML({ label, value, sub, cls }) {
    return `<article class="kpi ${cls || ""}"><div class="kpi-label">${label}</div><div class="kpi-value num">${value}</div>${sub ? `<div class="kpi-sub">${sub}</div>` : ""}</article>`;
  }
  function pfLabel(pf) {
    if (pf === null) return "—";
    if (!Number.isFinite(pf)) return "∞";
    return fmt(pf, 2);
  }
  function pnlClass(v) {
    if (v === null || v === undefined) return "muted";
    if (v > 0) return "up";
    if (v < 0) return "down";
    return "muted";
  }

  function insightText(s, rows) {
    if (!s.nClosed) {
      return `<h3>Insight</h3><p>No closed trades in this slice. Log a complete ticket (entry and exit) to see performance. Open positions never enter win rate, profit factor, or equity.</p>`;
    }
    const pairG = groupStats(rows, (t) => t.pair);
    const dayG = groupStats(rows, (t) => t.dow, WEEKDAYS_MON);
    const sesG = groupStats(rows, (t) => t.session, SESSIONS);
    const setG = groupStats(rows, (t) => t.setupLabel);
    const bp = bestWorst(pairG);
    const bd = bestWorst(dayG);
    const bs = bestWorst(sesG);
    const bt = bestWorst(setG);
    const bits = [];
    if (bp.best) bits.push(`Best pair <strong>${bp.best.key}</strong> (${money(bp.best.totalPnl)})`);
    if (bp.worst && bp.worst.key !== bp.best.key) bits.push(`least <strong>${bp.worst.key}</strong> (${money(bp.worst.totalPnl)})`);
    if (bd.best) bits.push(`best weekday <strong>${bd.best.key}</strong>`);
    if (bd.worst && bd.worst.key !== bd.best.key) bits.push(`least <strong>${bd.worst.key}</strong>`);
    if (bs.best) bits.push(`best session <strong>${bs.best.key}</strong>`);
    if (bt.best) bits.push(`best setup <strong>${bt.best.key}</strong>`);
    const slice = [];
    if (filters.pair !== "ALL") slice.push(filters.pair);
    if (filters.day !== "ALL") slice.push(filters.day);
    if (filters.session !== "ALL") slice.push(filters.session);
    if (filters.setup !== "ALL") slice.push(filters.setup);
    const head = slice.length ? `Filtered to ${slice.join(" · ")}. ` : "All trades. ";
    return `<h3>Insight</h3><p>${head}${bits.join(" · ")}.</p>
      <div class="pill-row">
        <div class="pill">Expectancy <strong>${s.expectancy === null ? "—" : money(s.expectancy)}</strong></div>
        <div class="pill">Avg RR <strong>${s.avgRR === null ? "—" : fmt(s.avgRR, 2)}</strong></div>
        <div class="pill">Win rate excl. BE <strong>${s.winRateExcl === null ? "—" : fmt(s.winRateExcl, 1) + "%"}</strong></div>
        <div class="pill">Gross profit <strong>${money(s.grossProfit)}</strong></div>
        <div class="pill">Gross loss <strong>${money(-s.grossLossAbs)}</strong></div>
      </div>`;
  }

  function updateIdentity() {
    const name = (settings.traderName || "").trim();
    const tid = (settings.traderId || "").trim();
    const set = !!(name || tid);
    const card = $("traderCard");
    const chip = $("identityChip");
    if (card) card.classList.toggle("empty", !set);
    if (chip) chip.classList.toggle("unset", !set);
    $("sideTraderName").textContent = name || "Set name in Settings";
    $("sideTraderId").textContent = tid ? "ID  " + tid : "Trader ID —";
    $("idName").textContent = name || "Trader not set";
    $("idCode").textContent = tid || (name ? "Add Trader ID in Settings" : "Add name & ID in Settings");
    const tk = $("tk-trader");
    if (tk) tk.textContent = [name, tid].filter(Boolean).join(" · ");
  }

  function updateSide() {
    const all = enrichAll();
    const { lastBalance, maxDd } = withBalances(all);
    $("sideBalance").textContent = money(lastBalance);
    const s = stats(all);
    const acct = settings.accountName ? settings.accountName + " · " : "";
    $("sideMeta").textContent = `${acct}${s.nClosed} closed · DD ${money(maxDd)}`;
    updateIdentity();
  }

  /* ───────── dashboard ───────── */
  function renderDashboard() {
    $("dashFilters").innerHTML = filterBarHTML("dash");
    bindFilterBar($("dashFilters"));

    const all = enrichAll();
    const rows = applyFilters(all);
    const s = stats(rows);
    const bal = withBalances(rows);

    const openN = all.filter((t) => !t.closed).length;
    const banner = $("openBanner");
    if (openN) {
      banner.classList.add("show");
      banner.textContent = `${openN} open trade${openN > 1 ? "s" : ""} excluded from performance until an exit is logged.`;
    } else banner.classList.remove("show");

    $("kpiPrimary").innerHTML = [
      kpiHTML({ label: "Closed trades", value: s.nClosed, sub: `${all.length} total in journal`, cls: "muted" }),
      kpiHTML({ label: "Win rate", value: s.winRate === null ? "—" : fmt(s.winRate, 1) + "%", sub: `${s.nWins}W / ${s.nLosses}L / ${s.nBE} BE`, cls: "gold" }),
      kpiHTML({ label: "Total PnL", value: money(s.totalPnl), cls: pnlClass(s.totalPnl) }),
      kpiHTML({ label: "Profit factor", value: pfLabel(s.profitFactor), sub: "Gross profit ÷ gross loss", cls: "gold" }),
      kpiHTML({ label: "Equity", value: money(bal.lastBalance), sub: `Start ${money(settings.initialCapital)}`, cls: "gold" })
    ].join("");

    $("kpiSecondary").innerHTML = [
      kpiHTML({ label: "Wins", value: s.nWins, cls: "up" }),
      kpiHTML({ label: "Losses", value: s.nLosses, cls: "down" }),
      kpiHTML({ label: "Breakeven", value: s.nBE, cls: "muted" }),
      kpiHTML({ label: "Largest win", value: money(s.largestWin), cls: "up" }),
      kpiHTML({ label: "Largest loss", value: money(s.largestLoss), cls: "down" }),
      kpiHTML({ label: "Average win", value: money(s.avgWin), cls: "up" }),
      kpiHTML({ label: "Average loss", value: money(s.avgLoss), cls: "down" }),
      kpiHTML({ label: "Max drawdown", value: money(-bal.maxDd), cls: "down" })
    ].join("");

    $("dashInsight").innerHTML = insightText(s, rows);

    const eq = bal.equity;
    emptyOrChart("chEquity", eq.length > 1, (c) =>
      lineChart(c, eq.map((x) => x.label), eq.map((x) => round(x.balance, 2)), gold, "rgba(201,168,76,0.12)")
    );
    emptyOrChart("chCumPnl", eq.length > 1, (c) =>
      lineChart(c, eq.map((x) => x.label), eq.map((x) => round(x.cum, 2)), up, "rgba(62,207,142,0.10)")
    );
    emptyOrChart("chOutcome", s.nClosed > 0, (c) => new Chart(c, {
      type: "doughnut",
      data: {
        labels: ["Win", "Loss", "Breakeven"],
        datasets: [{ data: [s.nWins, s.nLosses, s.nBE], backgroundColor: [up, down, muted], borderWidth: 0 }]
      },
      options: { responsive: true, maintainAspectRatio: false, cutout: "62%", plugins: { legend: { position: "bottom" } } }
    }));

    const ses = groupStats(rows, (t) => t.session, SESSIONS);
    emptyOrChart("chSession", ses.some((x) => x.nClosed), (c) =>
      barChart(c, ses.map((x) => x.key), ses.map((x) => round(x.totalPnl, 2)))
    );
    const pairs = groupStats(rows, (t) => t.pair).sort((a, b) => b.totalPnl - a.totalPnl);
    emptyOrChart("chPairPnl", pairs.length, (c) =>
      barChart(c, pairs.map((x) => x.key), pairs.map((x) => round(x.totalPnl, 2)), { horizontal: pairs.length > 8 })
    );
    const days = groupStats(rows, (t) => t.dow, WEEKDAYS_MON);
    emptyOrChart("chDayPnl", days.some((x) => x.nClosed), (c) =>
      barChart(c, days.map((x) => x.key.slice(0, 3)), days.map((x) => round(x.totalPnl, 2)))
    );
    const sets = groupStats(rows, (t) => t.setupLabel);
    emptyOrChart("chSetupPnl", sets.length, (c) =>
      barChart(c, sets.map((x) => x.key), sets.map((x) => round(x.totalPnl, 2)))
    );

    const recent = s.closed.slice().reverse().slice(0, 10);
    $("recentTape").innerHTML = recent.length
      ? recent.map((t) => {
        const cls = t.netPnl > 0 ? "pnl-pos" : t.netPnl < 0 ? "pnl-neg" : "pnl-zero";
        return `<div class="tape-row"><span>#${t.sn}</span><span class="pair">${t.pair}</span><span>${t.position}</span><span>${t.setupLabel || ""}</span><span>${t.exitDate || ""}</span><span class="${cls}">${money(t.netPnl)}</span></div>`;
      }).join("")
      : `<div class="empty-chart" style="min-height:80px">No closed trades yet</div>`;
  }

  /* ───────── analytics ───────── */
  function breakdownTable(title, groups) {
    if (!groups.length) return "";
    const body = groups.map((g) => {
      const cls = g.totalPnl > 0 ? "pnl-pos" : g.totalPnl < 0 ? "pnl-neg" : "pnl-zero";
      return `<tr>
        <td class="pair">${g.key}</td>
        <td>${g.nClosed}</td>
        <td>${g.nWins}</td>
        <td>${g.nLosses}</td>
        <td>${g.nBE}</td>
        <td>${g.winRate === null ? "—" : fmt(g.winRate, 1) + "%"}</td>
        <td class="${cls}">${money(g.totalPnl)}</td>
        <td>${pfLabel(g.profitFactor)}</td>
        <td>${money(g.avgWin)}</td>
        <td>${money(g.avgLoss)}</td>
        <td>${money(g.largestWin)}</td>
        <td>${money(g.largestLoss)}</td>
        <td>${g.expectancy === null ? "—" : money(g.expectancy)}</td>
      </tr>`;
    }).join("");
    return `<div class="table-card">
      <div class="table-toolbar"><div class="chart-title">${title}</div></div>
      <div class="table-scroll"><table class="data">
        <thead><tr><th>Slice</th><th>Trades</th><th>W</th><th>L</th><th>BE</th><th>Win %</th><th>PnL</th><th>PF</th><th>Avg W</th><th>Avg L</th><th>Max W</th><th>Max L</th><th>Exp.</th></tr></thead>
        <tbody>${body}</tbody>
      </table></div>
    </div>`;
  }

  function monthKey(t) {
    const d = t.exitDate || t.entryDate;
    return d ? d.slice(0, 7) : "—";
  }

  function renderAnalytics() {
    $("anFilters").innerHTML = filterBarHTML("an");
    bindFilterBar($("anFilters"));
    const rows = applyFilters(enrichAll());
    const s = stats(rows);
    $("anInsight").innerHTML = insightText(s, rows);

    $("anKpi").innerHTML = [
      kpiHTML({ label: "Trades in slice", value: s.nClosed, cls: "muted" }),
      kpiHTML({ label: "Win rate", value: s.winRate === null ? "—" : fmt(s.winRate, 1) + "%", cls: "gold" }),
      kpiHTML({ label: "PnL", value: money(s.totalPnl), cls: pnlClass(s.totalPnl) }),
      kpiHTML({ label: "Profit factor", value: pfLabel(s.profitFactor), cls: "gold" }),
      kpiHTML({ label: "Avg trade", value: money(s.expectancy), cls: pnlClass(s.expectancy) })
    ].join("");

    const pairs = groupStats(rows, (t) => t.pair);
    emptyOrChart("chWrPair", pairs.length, (c) =>
      barChart(c, pairs.map((x) => x.key), pairs.map((x) => round(x.winRate || 0, 1)), {
        colors: pairs.map(() => info), beginAtZero: true
      })
    );
    const sets = groupStats(rows, (t) => t.setupLabel);
    emptyOrChart("chWrSetup", sets.length, (c) =>
      barChart(c, sets.map((x) => x.key), sets.map((x) => round(x.winRate || 0, 1)), {
        colors: sets.map(() => gold), beginAtZero: true
      })
    );
    const days = groupStats(rows, (t) => t.dow, WEEKDAYS_MON);
    emptyOrChart("chWrDay", days.some((x) => x.nClosed), (c) =>
      barChart(c, days.map((x) => x.key.slice(0, 3)), days.map((x) => round(x.winRate || 0, 1)), {
        colors: days.map(() => up), beginAtZero: true
      })
    );
    const months = groupStats(rows, monthKey).sort((a, b) => a.key.localeCompare(b.key));
    emptyOrChart("chMonthly", months.length, (c) =>
      barChart(c, months.map((x) => x.key), months.map((x) => round(x.totalPnl, 2)))
    );

    const rrs = s.closed.map((t) => t.realizedRR).filter((x) => Number.isFinite(x));
    const buckets = ["<0", "0–0.5", "0.5–1", "1–1.5", "1.5–2", "2–3", ">3"];
    const counts = [0, 0, 0, 0, 0, 0, 0];
    rrs.forEach((x) => {
      if (x < 0) counts[0]++;
      else if (x < 0.5) counts[1]++;
      else if (x < 1) counts[2]++;
      else if (x < 1.5) counts[3]++;
      else if (x < 2) counts[4]++;
      else if (x < 3) counts[5]++;
      else counts[6]++;
    });
    emptyOrChart("chRr", rrs.length, (c) =>
      barChart(c, buckets, counts, { colors: buckets.map((_, i) => (i === 0 ? down : gold)), beginAtZero: true })
    );
    emptyOrChart("chOutcome2", s.nClosed > 0, (c) => new Chart(c, {
      type: "doughnut",
      data: {
        labels: ["Win", "Loss", "Breakeven"],
        datasets: [{ data: [s.nWins, s.nLosses, s.nBE], backgroundColor: [up, down, muted], borderWidth: 0 }]
      },
      options: { responsive: true, maintainAspectRatio: false, cutout: "62%", plugins: { legend: { position: "bottom" } } }
    }));

    $("breakTables").innerHTML =
      breakdownTable("By currency pair", pairs) +
      breakdownTable("By weekday", days) +
      breakdownTable("By session", groupStats(rows, (t) => t.session, SESSIONS)) +
      breakdownTable("By setup", sets);
  }

  /* ───────── journal ───────── */
  function renderJournal() {
    const q = ($("logSearch").value || "").toLowerCase();
    const st = $("logStatus").value;
    const all = enrichAll();
    const bal = withBalances(all);
    let rows = all.slice().reverse();
    if (st === "Closed") rows = rows.filter((t) => t.closed);
    if (st === "Open") rows = rows.filter((t) => !t.closed);
    if (q) {
      rows = rows.filter((t) =>
        [t.pair, t.setupLabel, t.notes, t.ticket, t.session, t.position, t.sn]
          .join(" ").toLowerCase().includes(q)
      );
    }
    const head = `<thead><tr>
      <th>S/N</th><th>Day</th><th>Entry</th><th>Pair</th><th>Pos</th><th>Session</th>
      <th>Lots</th><th>Risk</th><th>Entry px</th><th>SL</th><th>TP</th><th>Setup</th>
      <th>Exit px</th><th>Exit</th><th>Dur</th><th>RR</th><th>Pips</th>
      <th>PnL</th><th>Balance</th><th>Result</th><th></th>
    </tr></thead>`;
    const body = rows.map((t) => {
      const cls = t.netPnl > 0 ? "pnl-pos" : t.netPnl < 0 ? "pnl-neg" : "pnl-zero";
      const res = t.result === "Win" ? "win" : t.result === "Loss" ? "loss" : t.result === "BE" ? "be" : "open";
      const balAfter = t.closed ? bal.map.get(t.id) : null;
      return `<tr>
        <td>${t.sn}</td>
        <td>${t.dow ? t.dow.slice(0, 3) : ""}</td>
        <td>${t.entryDate || ""} ${t.entryTime || ""}</td>
        <td class="pair">${t.pair}</td>
        <td><span class="badge ${t.position === "Sell" ? "sell" : "buy"}">${t.position}</span></td>
        <td><span class="badge session">${t.session || ""}</span></td>
        <td>${isNum(t.lots) ? fmt(Number(t.lots), 2) : ""}</td>
        <td>${isNum(t.riskAmt) ? money(Number(t.riskAmt)) : ""}</td>
        <td>${isNum(t.entry) ? Number(t.entry).toFixed(t.decimals) : ""}</td>
        <td>${isNum(t.sl) ? Number(t.sl).toFixed(t.decimals) : ""}</td>
        <td>${isNum(t.tp) && t.tp !== "" ? Number(t.tp).toFixed(t.decimals) : ""}</td>
        <td>${t.setupLabel || ""}</td>
        <td>${isNum(t.exit) ? Number(t.exit).toFixed(t.decimals) : ""}</td>
        <td>${t.exitDate || ""} ${t.exitTime || ""}</td>
        <td>${t.duration}</td>
        <td>${t.realizedRR === null ? (t.plannedRR === null ? "" : fmt(t.plannedRR, 2) + "p") : fmt(t.realizedRR, 2)}</td>
        <td class="${t.resultPips > 0 ? "pnl-pos" : t.resultPips < 0 ? "pnl-neg" : ""}">${t.resultPips === null ? "" : fmt(t.resultPips, 1)}</td>
        <td class="${cls}">${t.netPnl === null ? "" : money(t.netPnl)}</td>
        <td>${balAfter === undefined || balAfter === null ? "" : money(balAfter)}</td>
        <td><span class="badge ${res}">${t.result}</span></td>
        <td><div class="row-actions">
          <button class="icon-btn" data-edit="${t.id}" title="Edit">✎</button>
          <button class="icon-btn danger" data-del="${t.id}" title="Delete">✕</button>
        </div></td>
      </tr>`;
    }).join("");
    $("logTable").innerHTML = head + `<tbody>${body || `<tr><td colspan="21" style="padding:28px;color:var(--text-3);font-family:var(--font)">No trades in the log. Use New trade to write the first ticket.</td></tr>`}</tbody>`;
  }

  $("logSearch").addEventListener("input", renderJournal);
  $("logStatus").addEventListener("change", renderJournal);
  $("logTable").addEventListener("click", async (e) => {
    const edit = e.target.closest("[data-edit]");
    const del = e.target.closest("[data-del]");
    if (edit) {
      showView("trade", { id: edit.dataset.edit });
    }
    if (del) {
      const ok = await confirmModal("Delete trade", "This cannot be undone.");
      if (!ok) return;
      trades = trades.filter((t) => t.id !== del.dataset.del);
      reindex();
      persistTrades();
      toast("Trade deleted");
      renderJournal();
      updateSide();
    }
  });

  /* ───────── settings ───────── */
  function renderSettings() {
    $("s-trader").value = settings.traderName || "";
    $("s-traderId").value = settings.traderId || "";
    $("s-name").value = settings.accountName || "";
    $("s-capital").value = settings.initialCapital;
    $("s-ccy").value = settings.currency;
    $("s-defRisk").value = settings.defaultRiskPct;
    $("s-be").value = settings.beThreshold;
    $("dataMeta").textContent = `${trades.length} trade(s) stored in this browser.`;
  }
  function collectSettingsFromForm() {
    settings.traderName = $("s-trader").value.trim();
    settings.traderId = $("s-traderId").value.trim();
    settings.accountName = $("s-name").value.trim();
    settings.initialCapital = n($("s-capital").value, 0);
    settings.currency = $("s-ccy").value;
    settings.defaultRiskPct = n($("s-defRisk").value, 1);
    settings.beThreshold = n($("s-be").value, 0.5);
  }
  $("saveSettings").addEventListener("click", () => {
    collectSettingsFromForm();
    persistSettings();
    toast("Settings saved");
    updateSide();
  });

  /* ───────── import / export ───────── */
  function csvEscape(v) {
    const s = v === null || v === undefined ? "" : String(v);
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  const CSV_COLS = [
    "sn","entryDate","entryTime","dow","pair","position","session","lots","riskAmt","riskPct",
    "entry","sl","tp","setup","setupCustom","exit","exitDate","exitTime","duration","plannedRR",
    "realizedRR","resultPips","commission","netPnl","result","plan","market","notes","tf",
    "emotion","ticket","chartUrl","mae","mfe","pnlOverride"
  ];

  function exportCsv() {
    const all = enrichAll();
    const bal = withBalances(all);
    const header = [...CSV_COLS, "balanceAfter"].join(",");
    const lines = all.map((t) => {
      const row = CSV_COLS.map((k) => csvEscape(t[k]));
      row.push(csvEscape(t.closed ? bal.map.get(t.id) : ""));
      return row.join(",");
    });
    const who = [settings.traderName, settings.traderId].filter(Boolean).join(" | ") || "unassigned";
    const meta = [
      "# Titan Capital Markets — Trading Journal",
      `# Trader: ${who}`,
      `# Account: ${settings.accountName || ""}`,
      `# Currency: ${settings.currency}  Initial capital: ${settings.initialCapital}`,
      `# Exported: ${new Date().toISOString()}`
    ].join("\n");
    const slug = (settings.traderId || "trades").replace(/[^\w.-]+/g, "_");
    downloadFile(`TCM-${slug}-journal.csv`, [meta, header, ...lines].join("\n"), "text/csv");
  }

  function exportJson() {
    const payload = {
      firm: "Titan Capital Markets",
      version: 1,
      exportedAt: new Date().toISOString(),
      settings,
      trades
    };
    const slug = (settings.traderId || "backup").replace(/[^\w.-]+/g, "_");
    downloadFile(`TCM-${slug}-backup.json`, JSON.stringify(payload, null, 2), "application/json");
  }

  function downloadFile(name, content, mime) {
    const blob = new Blob([content], { type: mime });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  $("exportCsvBtn").addEventListener("click", exportCsv);
  $("exportJsonBtn").addEventListener("click", exportJson);

  $("importJsonBtn").addEventListener("click", () => $("importJsonFile").click());
  $("importJsonFile").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (data.settings) {
          settings = { ...DEFAULT_SETTINGS, ...data.settings };
          persistSettings();
        }
        if (Array.isArray(data.trades)) {
          trades = data.trades;
          reindex();
          persistTrades();
        }
        toast("Backup restored");
        renderSettings();
        updateSide();
      } catch {
        toast("Could not read that JSON file", "bad");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  });

  $("importCsvBtn").addEventListener("click", () => $("importCsvFile").click());
  $("importCsvFile").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = parseCsv(reader.result);
        parsed.forEach((row) => {
          if (!row.pair) return;
          trades.push({
            id: uid(),
            pair: row.pair,
            position: row.position === "Sell" || row.position === "Sell/Short" ? "Sell" : "Buy",
            setup: row.setup || "Safety trade",
            setupCustom: row.setupCustom || "",
            entryDate: row.entryDate || row["entry date"] || "",
            entryTime: row.entryTime || "",
            session: row.session || "London",
            lots: row.lots || row["lot size"] || "",
            riskAmt: row.riskAmt || row["risk per trade"] || "",
            riskPct: row.riskPct || "",
            entry: row.entry || row["entry price"] || "",
            sl: row.sl || row["stop loss"] || "",
            tp: row.tp || row["take profit"] || "",
            exit: row.exit || row["exit price"] || "",
            exitDate: row.exitDate || "",
            exitTime: row.exitTime || "",
            exitReason: row.exitReason || "",
            commission: row.commission || 0,
            pnlOverride: row.pnlOverride || row.netPnl || row["profit/loss"] || "",
            notes: row.notes || "",
            plan: row.plan || "",
            market: row.market || "",
            tf: row.tf || "",
            emotion: row.emotion || "",
            ticket: row.ticket || "",
            chartUrl: row.chartUrl || "",
            mae: row.mae || "",
            mfe: row.mfe || ""
          });
        });
        reindex();
        persistTrades();
        toast(`Imported ${parsed.length} row(s)`);
        renderSettings();
      } catch {
        toast("CSV import failed", "bad");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  });

  function parseCsv(text) {
    const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter((l) => l.trim() && !l.trim().startsWith("#"));
    if (!lines.length) return [];
    const split = (line) => {
      const out = []; let cur = ""; let q = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (q) {
          if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
          else if (ch === '"') q = false;
          else cur += ch;
        } else if (ch === '"') q = true;
        else if (ch === ",") { out.push(cur); cur = ""; }
        else cur += ch;
      }
      out.push(cur);
      return out;
    };
    const headers = split(lines[0]).map((h) => h.trim());
    return lines.slice(1).map((line) => {
      const cols = split(line);
      const row = {};
      headers.forEach((h, i) => { row[h] = cols[i]; });
      return row;
    });
  }

  $("wipeBtn").addEventListener("click", async () => {
    const ok = await confirmModal("Wipe all trades", "Settings are kept. Trades are permanently removed from this browser.");
    if (!ok) return;
    trades = [];
    persistTrades();
    toast("Journal emptied");
    renderSettings();
    updateSide();
  });

  /* Labeled EXAMPLE trades — opt-in only, never shown as live stats by default */
  $("demoBtn").addEventListener("click", async () => {
    const ok = await confirmModal("Load example trades",
      "Eight rows tagged EXAMPLE in notes will be appended. They are synthetic and must not be treated as your results. Delete them from the log when done.");
    if (!ok) return;
    const demo = exampleTrades();
    demo.forEach((t) => trades.push(t));
    reindex();
    persistTrades();
    toast("Example trades appended — notes say EXAMPLE");
    showView("dashboard");
  });

  function exampleTrades() {
    const note = "EXAMPLE DATA — synthetic, not a live result. Delete after preview.";
    return [
      { id: uid(), pair: "EURUSD", position: "Buy", setup: "H1 5050 Bounce", entryDate: "2026-01-06", entryTime: "09:15", session: "London", lots: 0.5, riskAmt: 100, entry: 1.16520, sl: 1.16320, tp: 1.16920, exit: 1.16920, exitDate: "2026-01-06", exitTime: "11:40", commission: 3, notes: note, plan: "Yes", market: "Trending" },
      { id: uid(), pair: "GBPUSD", position: "Sell", setup: "Safety trade", entryDate: "2026-01-07", entryTime: "14:05", session: "New York", lots: 0.4, riskAmt: 80, entry: 1.27240, sl: 1.27440, tp: 1.26840, exit: 1.27440, exitDate: "2026-01-07", exitTime: "15:10", commission: 2.4, notes: note, plan: "Yes", market: "Ranging" },
      { id: uid(), pair: "USDJPY", position: "Buy", setup: "H4 50505 Bounce", entryDate: "2026-01-08", entryTime: "03:20", session: "Tokyo", lots: 0.3, riskAmt: 90, entry: 157.420, sl: 157.120, tp: 158.020, exit: 157.880, exitDate: "2026-01-08", exitTime: "08:45", commission: 2, notes: note, plan: "Yes" },
      { id: uid(), pair: "XAUUSD", position: "Buy", setup: "H1 5050 Bounce", entryDate: "2026-01-09", entryTime: "10:00", session: "London", lots: 0.10, riskAmt: 120, entry: 2645.20, sl: 2633.20, tp: 2669.20, exit: 2669.20, exitDate: "2026-01-09", exitTime: "14:22", commission: 4, notes: note, plan: "Yes", market: "Trending" },
      { id: uid(), pair: "EURUSD", position: "Sell", setup: "Safety trade", entryDate: "2026-01-10", entryTime: "15:30", session: "New York", lots: 0.5, riskAmt: 100, entry: 1.17110, sl: 1.17310, tp: 1.16710, exit: 1.17105, exitDate: "2026-01-10", exitTime: "16:05", commission: 3, notes: note, plan: "Yes" },
      { id: uid(), pair: "GBPJPY", position: "Buy", setup: "H1 5050 Bounce", entryDate: "2026-01-13", entryTime: "08:40", session: "London", lots: 0.2, riskAmt: 85, entry: 199.540, sl: 199.140, tp: 200.340, exit: 199.210, exitDate: "2026-01-13", exitTime: "10:15", commission: 2.5, pnlOverride: -72.4, notes: note, plan: "No", market: "Volatile / news" },
      { id: uid(), pair: "AUDUSD", position: "Sell", setup: "H4 50505 Bounce", entryDate: "2026-01-14", entryTime: "01:10", session: "Asian", lots: 0.6, riskAmt: 90, entry: 0.66240, sl: 0.66440, tp: 0.65840, exit: 0.65910, exitDate: "2026-01-14", exitTime: "06:50", commission: 2.2, notes: note, plan: "Yes" },
      { id: uid(), pair: "USDCAD", position: "Buy", setup: "Safety trade", entryDate: "2026-01-15", entryTime: "14:50", session: "New York", lots: 0.4, riskAmt: 70, entry: 1.43210, sl: 1.42960, tp: 1.43710, exit: 1.43580, exitDate: "2026-01-15", exitTime: "18:05", commission: 2, notes: note, plan: "Yes" }
    ];
  }

  document.addEventListener("keydown", (e) => {
    if (e.key === "n" && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      showView("trade");
    }
  });

  /* ───────── boot ───────── */
  load();
  initFormSelects();
  resetForm(true);
  updateSide();
  renderDashboard();
  updateIdentity();
})();
