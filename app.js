// =======================
// CONFIG (troque só isso)
// =======================
const CONFIG = {
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  TAB_NAME: "Números Cajazeiras", // nome EXATO da aba
  PREVIEW_ROWS: 12
};

let LAST_RESULT = null;

// -----------------------
// Helpers DOM
// -----------------------
function $(id) { return document.getElementById(id); }

function setStatus(msg, type = "info") {
  const el = $("status");
  if (!el) return console.log(`[${type}] ${msg}`);
  el.textContent = msg;
  el.dataset.type = type;
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

function isProbablyMonthLabel(s) {
  // "jan 2023" / "dez 2025"
  const v = safeText(s).trim().toLowerCase();
  return /^[a-zç]{3}\s\d{4}$/.test(v);
}

// -----------------------
// Google Sheets via gviz (JSONP)
// -----------------------
function gvizUrl(sheetId, tabName) {
  // output=gviz => retorna JS com um JSON dentro
  // headers=1 => primeira linha como cabeçalho
  const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq`;
  const params = new URLSearchParams({
    sheet: tabName,
    headers: "1",
    tq: "select *",
    tqx: "out:json"
  });
  return `${base}?${params.toString()}`;
}

async function fetchGvizTable(sheetId, tabName) {
  const url = gvizUrl(sheetId, tabName);
  const resp = await fetch(url);
  const text = await resp.text();

  // O retorno vem tipo: google.visualization.Query.setResponse({...});
  const match = text.match(/setResponse\(([\s\S]+)\);\s*$/);
  if (!match) throw new Error("Resposta gviz inesperada (não achei setResponse).");

  const json = JSON.parse(match[1]);
  if (json.status !== "ok") {
    throw new Error(json.errors?.[0]?.detailed_message || "Erro no gviz.");
  }

  return json.table; // { cols: [...], rows: [...] }
}

function tableToMatrix(table) {
  // header vem de cols[].label
  const header = table.cols.map(c => safeText(c.label));

  // rows vem de table.rows[].c[] com {v, f}
  const rows = table.rows.map(r => {
    const cells = r.c || [];
    return cells.map(cell => {
      if (!cell) return "";
      // preferir formatted (f) quando existir
      return safeText(cell.f ?? cell.v ?? "");
    });
  });

  return { header, rows };
}

function rowsToObjects(header, rows) {
  return rows.map(row => {
    const obj = {};
    header.forEach((h, i) => obj[h] = row[i] ?? "");
    return obj;
  });
}

function buildResult(unidade, header, rows) {
  const meses = header.filter(isProbablyMonthLabel);
  const records = rowsToObjects(header, rows);

  return {
    unidade,
    meses,
    records,
    raw: { header, rows },
    meta: {
      sheetId: CONFIG.SHEET_ID,
      tab: CONFIG.TAB_NAME,
      totalLinhas: rows.length,
      totalColunas: header.length
    }
  };
}

// -----------------------
// Render
// -----------------------
function renderMeses(meses) {
  const el = $("meses");
  if (!el) return;
  el.innerHTML = meses?.length ? `<div>${meses.join(" | ")}</div>` : "Nenhum mês detectado.";
}

function renderPreview(header, rows) {
  const el = $("preview");
  if (!el) return;

  const n = Math.min(CONFIG.PREVIEW_ROWS, rows.length);
  let html = `<table border="1" cellspacing="0" cellpadding="6"><thead><tr>`;
  header.forEach(h => html += `<th>${escapeHtml(h)}</th>`);
  html += `</tr></thead><tbody>`;

  for (let i = 0; i < n; i++) {
    html += `<tr>`;
    rows[i].forEach(v => html += `<td>${escapeHtml(v)}</td>`);
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
// Actions
// -----------------------
async function carregarCajazeiras() {
  setStatus("Lendo planilha (SheetID via gviz)...", "info");

  try {
    const table = await fetchGvizTable(CONFIG.SHEET_ID, CONFIG.TAB_NAME);
    const { header, rows } = tableToMatrix(table);

    const result = buildResult("Cajazeiras", header, rows);
    LAST_RESULT = result;

    renderMeses(result.meses);
    renderPreview(result.raw.header, result.raw.rows);
    renderJSON(result);

    setStatus(`OK: ${result.meta.totalLinhas} linhas, ${result.meta.totalColunas} colunas`, "success");
  } catch (err) {
    console.error(err);
    setStatus(`Erro: ${err.message}`, "error");
  }
}

function bindUI() {
  const btn = $("btnCarregar");
  if (btn) btn.addEventListener("click", carregarCajazeiras);
}

// Expor para onclick se quiser
window.KPIsGT = { carregarCajazeiras };

document.addEventListener("DOMContentLoaded", () => {
  bindUI();
  // auto-carregar se quiser:
  // carregarCajazeiras();
});
