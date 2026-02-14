/**
 * KPIs GT — Leitura via SheetID (GVIZ) — SEM WebApp
 * - Lê a aba (unidade)
 * - Seletor de mês base (default = último mês)
 * - Cards com setas e cores (vs mês anterior e vs mesmo mês ano anterior)
 * - Tabela abaixo
 */

const CONFIG = {
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  // abas/unidades (você vai ajustar depois)
  UNIDADES: [
    { label: "Cajazeiras", tab: "Números Cajazeiras" },
    { label: "Camaçari", tab: "Números Camaçari" },
    { label: "São Cristóvão", tab: "Números São Cristóvão" },
  ],
  PREVIEW_ROWS: 9999,
  DEBUG: false,
};

// Indicadores para os cards (ajuste nomes conforme sua planilha)
const KPI_CONFIG = [
  { key: "FATURAMENTO T (R$)", label: "Faturamento", better: "up", fmt: "brl" },
  { key: "CUSTO TOTAL T (R$)", label: "Custo Total", better: "down", fmt: "brl" },
  { key: "ALUNOS ATIVOS T", label: "Ativos", better: "up", fmt: "int" },
  { key: "META DE MATRÍCULA", label: "Meta", better: "up", fmt: "int" },
  { key: "MATRÍCULAS REALIZADAS", label: "Matrículas", better: "up", fmt: "int" },
  { key: "CONVERSÃO (%)", label: "Conversão", better: "up", fmt: "pct" },
  { key: "INADIMPLÊNCIA (META 9,3%)", label: "Inadimplência", better: "down", fmt: "pct" },
  { key: "EVASÃO REAL (META 4,8%)", label: "Evasão", better: "down", fmt: "pct" },
];

let LAST_RAW = null;     // { header, rows } da aba selecionada
let LAST_SERIES = null;  // series por mês

// -----------------------
// DOM
// -----------------------
function $(id) { return document.getElementById(id); }

function setStatus(msg, type = "info") {
  const el = $("status");
  if (el) {
    el.textContent = msg;
    el.dataset.type = type;
  }
  console.log(`[${type}] ${msg}`);
}

function safeText(v) { return (v ?? "").toString(); }

function escapeHtml(str) {
  return safeText(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// -----------------------
// Meses (padrão: janeiro/23, março/24...)
// -----------------------
function isMonthLabelBR(s) {
  const v = safeText(s).trim().toLowerCase();
  return /^([a-zçãáâàéêíóôõúûü]{3,12})\/(\d{2}|\d{4})$/.test(v);
}

function parseMonthKeyFromHeader(label) {
  // "janeiro/23" -> "2023-01"
  const v = safeText(label).trim().toLowerCase();
  const m = v.match(/^([a-zçãáâàéêíóôõúûü]{3,12})\/(\d{2}|\d{4})$/);
  if (!m) return null;

  const nome = m[1];
  let ano = m[2];
  if (ano.length === 2) ano = "20" + ano;

  const map = {
    janeiro: 1, fevereiro: 2, "março": 3, marco: 3,
    abril: 4, maio: 5, junho: 6, julho: 7,
    agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12
  };

  const mes = map[nome];
  if (!mes) return null;

  return `${ano}-${String(mes).padStart(2, "0")}`;
}

function monthKeyToDate(monthKey) {
  const [y, m] = monthKey.split("-").map(Number);
  if (!y || !m) return null;
  return new Date(y, m - 1, 1);
}

function addMonths(monthKey, delta) {
  const dt = monthKeyToDate(monthKey);
  if (!dt) return null;
  dt.setMonth(dt.getMonth() + delta);
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function getYear(monthKey) {
  return Number(monthKey.split("-")[0]);
}

function formatMonthKeyBR(monthKey) {
  // "2025-12" -> "dez/25"
  const [y, m] = monthKey.split("-").map(Number);
  const map = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
  return `${map[m - 1]}/${String(y).slice(2)}`;
}

function toNumberSmart(v) {
  let s = safeText(v).trim();
  if (!s) return null;

  s = s.replaceAll("R$", "").replaceAll(/\s+/g, "");
  s = s.replaceAll("%", "");
  s = s.replaceAll(".", "").replaceAll(",", ".");

  const n = Number(s);
  if (Number.isNaN(n)) return null;
  return n;
}

// -----------------------
// GVIZ
// -----------------------
function gvizUrl(sheetId, tabName) {
  const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq`;
  const params = new URLSearchParams({
    sheet: tabName,
    headers: "1",
    tq: "select *",
    tqx: "out:json",
  });
  return `${base}?${params.toString()}`;
}

async function fetchGvizTable(sheetId, tabName) {
  const url = gvizUrl(sheetId, tabName);
  setStatus("Carregando dados...", "info");

  const resp = await fetch(url);
  const text = await resp.text();

  const match = text.match(/setResponse\(([\s\S]+)\);\s*$/);
  if (!match) throw new Error("Não achei setResponse(). Verifique permissão/aba.");

  const json = JSON.parse(match[1]);
  if (json.status !== "ok") {
    throw new Error(json.errors?.[0]?.detailed_message || "Erro na leitura.");
  }
  return json.table;
}

function tableToMatrix(table) {
  const header = table.cols.map((c) => safeText(c.label));
  const rows = table.rows.map((r) => {
    const cells = r.c || [];
    return cells.map((cell) => safeText(cell?.f ?? cell?.v ?? ""));
  });
  return { header, rows };
}

// -----------------------
// Série por mês + Fechamento
// -----------------------
function matrixToSeriesByMonth(header, rows) {
  const monthCols = [];
  for (let i = 0; i < header.length; i++) {
    const key = parseMonthKeyFromHeader(header[i]);
    if (key) monthCols.push({ idx: i, key, label: header[i] });
  }

  const series = {};
  monthCols.forEach(c => { series[c.key] = {}; });

  rows.forEach(r => {
    const indicador = safeText(r[0]).trim();
    if (!indicador) return;

    monthCols.forEach(c => {
      const raw = r[c.idx];
      const num = toNumberSmart(raw);
      series[c.key][indicador] = (num !== null ? num : safeText(raw).trim());
    });
  });

  return { series, monthCols };
}

function pickCurrentMonthKey(series) {
  const keys = Object.keys(series).sort();
  for (let i = keys.length - 1; i >= 0; i--) {
    const k = keys[i];
    const obj = series[k];
    const hasSomething = obj && Object.values(obj).some(v => v !== "" && v !== null && v !== undefined);
    if (hasSomething) return k;
  }
  return keys[keys.length - 1] || null;
}

function buildFechamento(series, monthKeyCurrent = null) {
  const mesAtual = monthKeyCurrent || pickCurrentMonthKey(series);
  if (!mesAtual) throw new Error("Não consegui determinar o mês atual.");

  const ano = getYear(mesAtual);

  const m1 = addMonths(mesAtual, -1);
  const m2 = addMonths(mesAtual, -2);
  const m3 = addMonths(mesAtual, -3);

  const ultimos3AnoCorrente = [m1, m2, m3].filter(k => k && getYear(k) === ano && series[k]);

  const mesAnoAnterior = addMonths(mesAtual, -12);
  const anoAnteriorExiste = mesAnoAnterior && series[mesAnoAnterior] ? mesAnoAnterior : null;

  const variacoes = {};
  const indicadores = new Set(Object.keys(series[mesAtual] || {}));

  indicadores.forEach(ind => {
    const atual = series[mesAtual]?.[ind] ?? null;
    const ant = (m1 && series[m1]) ? (series[m1][ind] ?? null) : null;
    const aa = anoAnteriorExiste ? (series[anoAnteriorExiste][ind] ?? null) : null;

    const isNumA = typeof atual === "number";
    const isNumB = typeof ant === "number";
    const isNumC = typeof aa === "number";

    variacoes[ind] = {
      atual,
      mesAnterior: ant,
      anoAnterior: aa,
      deltaMesAnterior: (isNumA && isNumB) ? (atual - ant) : null,
      deltaAnoAnterior: (isNumA && isNumC) ? (atual - aa) : null,
    };
  });

  return { mesAtual, ultimos3AnoCorrente, mesAnoAnterior: anoAnteriorExiste, variacoes };
}

// -----------------------
// UI: mês base + unidade
// -----------------------
function fillUnidadeOptions() {
  const sel = $("unidade");
  if (!sel) return;

  sel.innerHTML = "";
  CONFIG.UNIDADES.forEach(u => {
    const opt = document.createElement("option");
    opt.value = u.tab;
    opt.textContent = u.label;
    sel.appendChild(opt);
  });

  // default: Cajazeiras
  sel.value = CONFIG.UNIDADES[0]?.tab || CONFIG.TAB_NAME;
}

function fillMesBaseOptions(series) {
  const sel = $("mesBase");
  if (!sel) return;

  const keys = Object.keys(series).sort();
  sel.innerHTML = "";
  keys.forEach(k => {
    const opt = document.createElement("option");
    opt.value = k;
    opt.textContent = formatMonthKeyBR(k);
    sel.appendChild(opt);
  });

  const current = pickCurrentMonthKey(series) || keys[keys.length - 1];
  if (current) sel.value = current;
}

// -----------------------
// Render: cards setas/cores
// -----------------------
function fmtValue(val, fmt) {
  if (val === null || val === undefined || val === "") return "—";
  if (typeof val === "number") {
    if (fmt === "brl") return val.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
    if (fmt === "pct") return `${val.toLocaleString("pt-BR", { maximumFractionDigits: 2 })}%`;
    if (fmt === "int") return val.toLocaleString("pt-BR", { maximumFractionDigits: 0 });
    return val.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
  }
  return safeText(val);
}

function trendMeta(delta, better) {
  if (delta === null || delta === undefined) return { arrow: "→", cls: "neutral", sign: "" };
  if (delta === 0) return { arrow: "→", cls: "neutral", sign: "" };

  const isUp = delta > 0;
  const good = (better === "up") ? isUp : !isUp;

  return { arrow: isUp ? "▲" : "▼", cls: good ? "good" : "bad", sign: delta > 0 ? "+" : "" };
}

function renderFechamentoText(fechamento) {
  const box = $("fechamento");
  if (!box) return;

  const ult3 = fechamento.ultimos3AnoCorrente.map(formatMonthKeyBR).join(" | ");
  box.textContent =
    `Mês base: ${formatMonthKeyBR(fechamento.mesAtual)}\n` +
    `Últimos 3 (ano corrente): ${ult3 || "—"}\n` +
    `Mesmo mês ano anterior: ${fechamento.mesAnoAnterior ? formatMonthKeyBR(fechamento.mesAnoAnterior) : "—"}\n`;
}

function renderKPICards(fechamento) {
  const wrap = $("cardsKPI");
  if (!wrap) return;

  wrap.innerHTML = "";
  wrap.style.display = "grid";
  wrap.style.gridTemplateColumns = "repeat(auto-fit, minmax(220px, 1fr))";
  wrap.style.gap = "12px";
  wrap.style.margin = "12px 0";

  const mBase = fechamento.mesAtual;
  const m1 = addMonths(mBase, -1);
  const m12 = fechamento.mesAnoAnterior;

  KPI_CONFIG.forEach(kpi => {
    const v = fechamento.variacoes?.[kpi.key] || {};
    const atual = v.atual;
    const dMes = v.deltaMesAnterior;
    const dAno = v.deltaAnoAnterior;

    const tMes = trendMeta(dMes, kpi.better);
    const tAno = trendMeta(dAno, kpi.better);

    const card = document.createElement("div");
    card.style.border = "1px solid rgba(0,0,0,.12)";
    card.style.borderRadius = "12px";
    card.style.padding = "12px";
    card.style.background = "white";

    const title = document.createElement("div");
    title.style.fontWeight = "800";
    title.style.marginBottom = "8px";
    title.textContent = kpi.label;

    const value = document.createElement("div");
    value.style.fontSize = "20px";
    value.style.fontWeight = "900";
    value.style.marginBottom = "8px";
    value.textContent = fmtValue(atual, kpi.fmt);

    function line(label, t, delta, refTxt) {
      const row = document.createElement("div");
      row.style.display = "flex";
      row.style.justifyContent = "space-between";
      row.style.gap = "10px";
      row.style.fontSize = "13px";

      const left = document.createElement("span");
      left.textContent = `${label} ${refTxt}`;

      const right = document.createElement("span");
      right.textContent = `${t.arrow} ${t.sign}${fmtValue(delta, kpi.fmt === "pct" ? "pct" : "num")}`;
      right.style.fontWeight = "800";
      right.style.color = t.cls === "good" ? "#16794C" : t.cls === "bad" ? "#B42318" : "#555";

      row.appendChild(left);
      row.appendChild(right);
      return row;
    }

    card.appendChild(title);
    card.appendChild(value);
    card.appendChild(line("vs", tMes, dMes, m1 ? formatMonthKeyBR(m1) : "M-1"));
    card.appendChild(line("vs", tAno, dAno, m12 ? formatMonthKeyBR(m12) : "ano ant."));

    wrap.appendChild(card);
  });
}

// -----------------------
// Render: tabela
// -----------------------
function renderTable(header, rows) {
  const el = $("preview");
  if (!el) return;

  const n = Math.min(CONFIG.PREVIEW_ROWS, rows.length);

  let html = `<table border="1" cellspacing="0" cellpadding="6"><thead><tr>`;
  header.forEach((h) => (html += `<th>${escapeHtml(h)}</th>`));
  html += `</tr></thead><tbody>`;

  for (let i = 0; i < n; i++) {
    html += `<tr>`;
    (rows[i] || []).forEach((v) => (html += `<td>${escapeHtml(v)}</td>`));
    html += `</tr>`;
  }

  html += `</tbody></table>`;
  el.innerHTML = html;
}

function renderJSON(obj) {
  const el = $("json");
  if (!el) return;
  if (!CONFIG.DEBUG) return;
  el.style.display = "block";
  el.textContent = JSON.stringify(obj, null, 2);
}

// -----------------------
// Actions
// -----------------------
async function carregarDados() {
  try {
    const selUnidade = $("unidade");
    const tab = selUnidade && selUnidade.value ? selUnidade.value : (CONFIG.UNIDADES[0]?.tab || "Números Cajazeiras");

    const table = await fetchGvizTable(CONFIG.SHEET_ID, tab);
    const raw = tableToMatrix(table);

    LAST_RAW = raw;

    const { series } = matrixToSeriesByMonth(raw.header, raw.rows);
    LAST_SERIES = series;

    fillMesBaseOptions(series);

    setStatus("Dados carregados.", "success");

    aplicar();
  } catch (err) {
    console.error(err);
    setStatus(`ERRO: ${err.message}`, "error");
  }
}

function aplicar() {
  if (!LAST_RAW || !LAST_SERIES) {
    setStatus("Clique em Carregar primeiro.", "warn");
    return;
  }

  const selMes = $("mesBase");
  const mesEscolhido = selMes && selMes.value ? selMes.value : null;

  const fechamento = buildFechamento(LAST_SERIES, mesEscolhido);
  console.log("FECHAMENTO:", fechamento);

  renderFechamentoText(fechamento);
  renderKPICards(fechamento);

  renderTable(LAST_RAW.header, LAST_RAW.rows);

  renderJSON({ fechamento });
}

// -----------------------
// Init
// -----------------------
document.addEventListener("DOMContentLoaded", () => {
  fillUnidadeOptions();

  const btnCarregar = $("btnCarregar");
  if (btnCarregar) btnCarregar.addEventListener("click", carregarDados);

  const btnAplicar = $("btnAplicar");
  if (btnAplicar) btnAplicar.addEventListener("click", aplicar);

  const selMes = $("mesBase");
  if (selMes) selMes.addEventListener("change", aplicar);

  const selUnidade = $("unidade");
  if (selUnidade) selUnidade.addEventListener("change", carregarDados);

  setStatus("Pronto. Clique em Carregar.", "info");
});

// Expor (opcional)
window.KPIsGT = { carregarDados, aplicar };
