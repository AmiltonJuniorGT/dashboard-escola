/**
 * KPIs GT — Leitura via SheetID (GVIZ) — SEM WebApp
 * - Lê a aba (unidade)
 * - Renderiza tabela
 * - Filtro de período (início/fim) filtrando colunas de meses
 * - Motor Fechamento: mês atual + últimos 3 meses do ano + mesmo mês ano anterior
 * - Não mostra meses abaixo do botão
 * - Não mostra JSON (a não ser DEBUG=true)
 *
 * IDs esperados no HTML:
 *  - btnCarregar
 *  - status
 *  - preview
 *  - meses (opcional, será ocultado)
 *  - json  (opcional, será ocultado por padrão)
 *  - periodoInicio (opcional, input type="date")
 *  - periodoFim    (opcional, input type="date")
 *  - btnAplicar    (opcional)
 */

// =======================
// CONFIG
// =======================
const CONFIG = {
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  TAB_NAME: "Números Cajazeiras",
  PREVIEW_ROWS: 9999,
  DEBUG: false,
};

let LAST_RAW = null;      // { header, rows }
let LAST_RENDER = null;   // filtrado

// -----------------------
// Helpers DOM
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
// Filtro início/fim (filtra colunas de mês)
// -----------------------
function parseDateInput(id) {
  const el = $(id);
  if (!el || !el.value) return null;
  const [y, m, d] = el.value.split("-").map(Number);
  if (!y || !m) return null;
  return new Date(y, m - 1, d || 1);
}

function endOfMonth(dt) {
  return new Date(dt.getFullYear(), dt.getMonth() + 1, 0);
}

function filterByPeriod(raw, inicio, fim) {
  if (!inicio && !fim) return raw;

  const start = inicio ? new Date(inicio.getFullYear(), inicio.getMonth(), 1) : null;
  const end = fim ? endOfMonth(fim) : null;

  const keepIdx = [0]; // mantém indicador

  for (let i = 1; i < raw.header.length; i++) {
    const h = raw.header[i];
    if (!isMonthLabelBR(h)) continue;

    const key = parseMonthKeyFromHeader(h);
    if (!key) continue;
    const colDate = monthKeyToDate(key);
    if (!colDate) continue;

    const colStart = new Date(colDate.getFullYear(), colDate.getMonth(), 1);
    const colEnd = endOfMonth(colDate);

    const okStart = start ? (colEnd >= start) : true;
    const okEnd = end ? (colStart <= end) : true;

    if (okStart && okEnd) keepIdx.push(i);
  }

  const header = keepIdx.map(i => raw.header[i]);
  const rows = raw.rows.map(r => keepIdx.map(i => r[i] ?? ""));
  return { header, rows };
}

// -----------------------
// Motor Fechamento
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

  // variações por indicador (opcional, já fica pronto)
  const indicadores = new Set(Object.keys(series[mesAtual] || {}));
  const variacoes = {};
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
// Render
// -----------------------
function hideMesesArea() {
  const el = $("meses");
  if (!el) return;
  el.textContent = "";
  el.style.display = "none";
}

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

  if (!CONFIG.DEBUG) {
    el.textContent = "";
    el.style.display = "none";
    return;
  }
  el.style.display = "block";
  el.textContent = JSON.stringify(obj, null, 2);
}

// -----------------------
// Actions
// -----------------------
async function carregarDados() {
  hideMesesArea();

  try {
    const table = await fetchGvizTable(CONFIG.SHEET_ID, CONFIG.TAB_NAME);
    const raw = tableToMatrix(table);
    LAST_RAW = raw;

    aplicarFiltroPeriodo(); // render + motor
  } catch (err) {
    console.error(err);
    setStatus(`ERRO: ${err.message}`, "error");
  }
}

function aplicarFiltroPeriodo() {
  if (!LAST_RAW) {
    setStatus("Carregue os dados primeiro.", "warn");
    return;
  }

  const inicio = parseDateInput("periodoInicio");
  const fim = parseDateInput("periodoFim");

  const filtered = filterByPeriod(LAST_RAW, inicio, fim);
  LAST_RENDER = filtered;

  renderTable(filtered.header, filtered.rows);
  setStatus("Dados carregados.", "success");

  // MOTOR FECHAMENTO (usa a tabela completa, não a filtrada)
  const { series } = matrixToSeriesByMonth(LAST_RAW.header, LAST_RAW.rows);
  const fechamento = buildFechamento(series);

  console.log("FECHAMENTO:", fechamento);

  renderJSON({
    tab: CONFIG.TAB_NAME,
    filtro: {
      inicio: inicio ? $("periodoInicio").value : null,
      fim: fim ? $("periodoFim").value : null,
    },
    fechamento,
    meta: {
      linhas: filtered.rows.length,
      colunas: filtered.header.length,
    },
  });
}

// -----------------------
// Bind UI
// -----------------------
document.addEventListener("DOMContentLoaded", () => {
  hideMesesArea();

  const btnCarregar = $("btnCarregar");
  if (btnCarregar) btnCarregar.addEventListener("click", carregarDados);

  const btnAplicar = $("btnAplicar");
  if (btnAplicar) btnAplicar.addEventListener("click", aplicarFiltroPeriodo);

  const pi = $("periodoInicio");
  const pf = $("periodoFim");
  if (pi) pi.addEventListener("change", () => LAST_RAW && aplicarFiltroPeriodo());
  if (pf) pf.addEventListener("change", () => LAST_RAW && aplicarFiltroPeriodo());

  setStatus("Pronto.", "info");
});

// Expor opcional
window.KPIsGT = { carregarDados, aplicarFiltroPeriodo };
