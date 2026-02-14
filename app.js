/**
 * KPIs GT — Leitura da planilha via SheetID (GVIZ) — SEM WebApp
 * Lê a aba "Números Cajazeiras", usa a linha 1 como cabeçalho (ex: janeiro/23 ...),
 * detecta meses e normaliza para "jan 2023", renderiza meses, preview e JSON.
 *
 * Requer no index.html:
 * <button id="btnCarregar">Carregar</button>
 * <div id="status"></div>
 * <div id="meses"></div>
 * <div id="preview"></div>
 * <pre id="json"></pre>
 * <script src="./app.js"></script>
 */

// =======================
// CONFIG (troque só isso)
// =======================
const CONFIG = {
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  TAB_NAME: "Números Cajazeiras", // nome EXATO da aba
  PREVIEW_ROWS: 12,              // linhas no preview
};

let LAST_RESULT = null;

// -----------------------
// Helpers DOM
// -----------------------
function $(id) { return document.getElementById(id); }

function setStatus(msg, type = "info") {
  const el = $("status");
  if (el) {
    el.textContent = msg;
    el.dataset.type = type; // opcional (pra CSS)
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
// Detecta meses no seu padrão real: "janeiro/23", "março/24", etc.
// -----------------------
function isProbablyMonthLabel(s) {
  const v = safeText(s).trim().toLowerCase();

  // Aceita: janeiro/23, fevereiro/23, março/23, abril/2024, jan/23, etc
  // (com ou sem acento)
  return /^([a-zçãáâàéêíóôõúûü]{3,12})\/(\d{2}|\d{4})$/.test(v);
}

function normalizeMonthLabel(label) {
  const v = safeText(label).trim().toLowerCase();
  const m = v.match(/^([a-zçãáâàéêíóôõúûü]{3,12})\/(\d{2}|\d{4})$/);
  if (!m) return label;

  const nome = m[1];      // "janeiro" ou "março" etc
  let ano = m[2];         // "23" ou "2023"
  if (ano.length === 2) ano = "20" + ano;

  const map = {
    janeiro: "jan",
    fevereiro: "fev",
    "março": "mar",
    marco: "mar",
    abril: "abr",
    maio: "mai",
    junho: "jun",
    julho: "jul",
    agosto: "ago",
    setembro: "set",
    outubro: "out",
    novembro: "nov",
    dezembro: "dez",
  };

  const abrev = map[nome] || nome.slice(0, 3);
  return `${abrev} ${ano}`;
}

// -----------------------
// Google Sheets via gviz (sem WebApp)
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
  setStatus("Carregando planilha (GVIZ)...", "info");

  const resp = await fetch(url);
  const text = await resp.text();

  // Retorna algo como: google.visualization.Query.setResponse({...});
  const match = text.match(/setResponse\(([\s\S]+)\);\s*$/);
  if (!match) {
    throw new Error(
      "Resposta GVIZ inesperada (não achei setResponse). Verifique permissão da planilha/aba."
    );
  }

  const json = JSON.parse(match[1]);
  if (json.status !== "ok") {
    throw new Error(json.errors?.[0]?.detailed_message || "Erro no GVIZ.");
  }

  return json.table; // { cols: [...], rows: [...] }
}

function tableToMatrix(table) {
  const header = table.cols.map((c) => safeText(c.label));

  const rows = table.rows.map((r) => {
    const cells = r.c || [];
    return cells.map((cell) => {
      if (!cell) return "";
      return safeText(cell.f ?? cell.v ?? "");
    });
  });

  return { header, rows };
}

// ========= MOTOR FECHAMENTO (M-1/M-2/M-3 do ano corrente + M-12) =========

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

  const mm = String(mes).padStart(2, "0");
  return `${ano}-${mm}`;
}

function monthKeyToDate(monthKey) {
  // "2025-12" -> Date(2025, 11, 1)
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
  // converte "R$ 337.232,12" / "16,07" / "1.616" / "12,44%" etc
  let s = safeText(v).trim();
  if (!s) return null;

  // remove moeda e espaços
  s = s.replaceAll("R$", "").replaceAll(/\s+/g, "");

  // percentual
  const isPct = s.includes("%");
  s = s.replaceAll("%", "");

  // se tem "." e "," no padrão BR: "." milhar, "," decimal
  // remove "." e troca "," por "."
  s = s.replaceAll(".", "").replaceAll(",", ".");

  const n = Number(s);
  if (Number.isNaN(n)) return null;

  return n; // mantenha como número puro; % você trata na renderização se quiser
}

function matrixToSeriesByMonth(header, rows) {
  // header: ["Indicador", "janeiro/23", "fevereiro/23", ...]
  // rows: [ ["Faturamento (R$)", "337.232,12", ...], ... ]
  // retorna: { "2025-12": { "Faturamento (R$)": 337232.12, ... }, ... }

  // 1) mapear colunas de mês
  const monthCols = []; // [{ idx, key, label }]
  for (let i = 0; i < header.length; i++) {
    const key = parseMonthKeyFromHeader(header[i]);
    if (key) monthCols.push({ idx: i, key, label: header[i] });
  }

  // 2) iniciar série vazia
  const series = {};
  monthCols.forEach(c => { series[c.key] = {}; });

  // 3) preencher por indicador
  rows.forEach(r => {
    const indicador = safeText(r[0]).trim();
    if (!indicador) return;

    monthCols.forEach(c => {
      const raw = r[c.idx];
      const num = toNumberSmart(raw);
      // guarda número se conseguir, senão guarda texto original
      series[c.key][indicador] = (num !== null ? num : safeText(raw).trim());
    });
  });

  return { series, monthCols };
}

function pickCurrentMonthKey(series) {
  // último mês com pelo menos algum valor numérico preenchido
  const keys = Object.keys(series).sort(); // YYYY-MM ordena lexicograficamente ok
  for (let i = keys.length - 1; i >= 0; i--) {
    const k = keys[i];
    const obj = series[k];
    if (!obj) continue;
    const hasSomething = Object.values(obj).some(v => v !== "" && v !== null && v !== undefined);
    if (hasSomething) return k;
  }
  return keys[keys.length - 1] || null;
}

// =======Motor FECHAMENTO  ....

function buildFechamento(series, monthKeyCurrent = null) {
  const mesAtual = monthKeyCurrent || pickCurrentMonthKey(series);
  if (!mesAtual) throw new Error("Não consegui determinar o mês atual.");

  const ano = getYear(mesAtual);

  const m1 = addMonths(mesAtual, -1);
  const m2 = addMonths(mesAtual, -2);
  const m3 = addMonths(mesAtual, -3);

  // últimos 3 meses do ano corrente (só mantém se estiver no mesmo ano)
  const ultimos3AnoCorrente = [m1, m2, m3].filter(k => k && getYear(k) === ano && series[k]);

  // mesmo mês do ano anterior
  const mesAnoAnterior = addMonths(mesAtual, -12);
  const anoAnteriorExiste = mesAnoAnterior && series[mesAnoAnterior] ? mesAnoAnterior : null;

  // conjunto de indicadores (união)
  const indicadores = new Set();
  Object.keys(series[mesAtual] || {}).forEach(k => indicadores.add(k));
  if (m1 && series[m1]) Object.keys(series[m1]).forEach(k => indicadores.add(k));
  if (anoAnteriorExiste) Object.keys(series[anoAnteriorExiste]).forEach(k => indicadores.add(k));

  // variações por indicador
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
      // percentuais (opcional)
      pctMesAnterior: (isNumA && isNumB && ant !== 0) ? ((atual - ant) / ant) : null,
      pctAnoAnterior: (isNumA && isNumC && aa !== 0) ? ((atual - aa) / aa) : null,
    };
  });

  return {
    mesAtual,
    ultimos3AnoCorrente,   // [M-1, M-2, M-3] dentro do mesmo ano
    mesAnoAnterior: anoAnteriorExiste,
    variacoes
  };
}

const { header, rows } = tableToMatrix(table);

// Converte tabela para série por mês
const { series } = matrixToSeriesByMonth(header, rows);

// Monta pacote de fechamento
const fechamento = buildFechamento(series);

// Apenas para teste agora:
console.log("FECHAMENTO:", fechamento);


// -----------------------
// Render
// -----------------------
function renderMeses(meses) {
  const el = $("meses");
  if (!el) return;

  if (!meses?.length) {
    el.textContent = "Nenhum mês detectado no cabeçalho.";
    return;
  }
  el.textContent = meses.join(" | ");
}

function renderPreview(header, rows) {
  const el = $("preview");
  if (!el) return;

  const n = Math.min(CONFIG.PREVIEW_ROWS, rows.length);

  let html = `<table border="1" cellspacing="0" cellpadding="6"><thead><tr>`;
  header.forEach((h) => (html += `<th>${escapeHtml(h)}</th>`));
  html += `</tr></thead><tbody>`;

  for (let i = 0; i < n; i++) {
    html += `<tr>`;
    rows[i].forEach((v) => (html += `<td>${escapeHtml(v)}</td>`));
    html += `</tr>`;
  }

  html += `</tbody></table>`;
  el.innerHTML = html;
}

function renderJSON(obj) {
  const el = $("json");
  if (!el) return;
  el.textContent = JSON.stringify(obj, null, 2);
}

// -----------------------
// Action principal
// -----------------------
async function carregarCajazeiras() {
  try {
    const table = await fetchGvizTable(CONFIG.SHEET_ID, CONFIG.TAB_NAME);
    setStatus("Tabela recebida. Processando...", "info");

    const { header, rows } = tableToMatrix(table);

    // Detecta e normaliza meses no formato "jan 2023"
    const meses = header
      .filter(isProbablyMonthLabel)
      .map(normalizeMonthLabel);

    const records = rowsToObjects(header, rows);

    const result = {
      unidade: "Cajazeiras",
      meses,
      records,
      meta: {
        sheetId: CONFIG.SHEET_ID,
        tab: CONFIG.TAB_NAME,
        totalLinhas: rows.length,
        totalColunas: header.length,
      },
      raw: { header, rows },
    };

    LAST_RESULT = result;

    renderMeses(result.meses);
    renderPreview(result.raw.header, result.raw.rows);
    renderJSON(result);

    setStatus(`OK: ${result.meta.totalLinhas} linhas, ${result.meta.totalColunas} colunas`, "success");
  } catch (err) {
    console.error(err);
    setStatus(`ERRO: ${err.message}`, "error");
  }
}

// -----------------------
// Bind UI
// -----------------------
document.addEventListener("DOMContentLoaded", () => {
  const btn = $("btnCarregar");
  if (!btn) {
    setStatus("ERRO: não achei o botão #btnCarregar no HTML.", "error");
    return;
  }

  btn.addEventListener("click", carregarCajazeiras);
  setStatus("Pronto. Clique em Carregar.", "info");

  // Se quiser auto-carregar ao abrir:
  // carregarCajazeiras();
});

// Expor (opcional)
window.KPIsGT = { carregarCajazeiras };
