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

function rowsToObjects(header, rows) {
  return rows.map((row) => {
    const obj = {};
    header.forEach((h, i) => (obj[h] = row[i] ?? ""));
    return obj;
  });
}

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
