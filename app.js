/**
 * KPIs GT — app.js (Opção 2: Apps Script WebApp)
 * Objetivo: ler a planilha base e estruturar meses (linha 1) + dados.
 *
 * Planilha base:
 * - Sheet ID: 1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go
 * - Aba: Números Cajazeiras
 * - Meses na linha 1 como texto (jan 2023 ... dez 2025)
 *
 * Requisitos no HTML:
 * - elementos com IDs (opcional, mas recomendado):
 *   - status
 *   - meses
 *   - preview
 *   - json
 *   - btnCarregar
 *   - btnBaixarJSON
 */

// =======================
// CONFIG (troque só isso)
// =======================
const CONFIG = {
  WEBAPP_URL: "COLE_AQUI_A_URL_DO_SEU_WEBAPP", // ex: https://script.google.com/macros/s/XXXX/exec
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  TAB_NAME: "Números Cajazeiras",

  // se você quiser limitar o range, pode usar "A1:ZZ200"
  // caso deixe vazio, o backend pode usar getDataRange()
  RANGE_A1: "",

  // quantas linhas mostrar no preview
  PREVIEW_ROWS: 10,
};

// =======================
// State
// =======================
let LAST_RESULT = null;

// =======================
// Utils DOM
// =======================
function $(id) {
  return document.getElementById(id);
}

function setStatus(msg, type = "info") {
  const el = $("status");
  if (!el) return;

  el.textContent = msg;
  el.dataset.type = type; // se quiser estilizar via CSS
}

function safeText(str) {
  return (str ?? "").toString();
}

function isProbablyMonthLabel(s) {
  // exemplo: "jan 2023" / "dez 2025"
  // critério simples: 3 letras + espaço + 4 dígitos
  const v = safeText(s).trim().toLowerCase();
  return /^[a-zç]{3}\s\d{4}$/.test(v);
}

// =======================
// Networking
// =======================
async function callWebApp(payload) {
  const url = CONFIG.WEBAPP_URL;

  if (!url || url.includes("COLE_AQUI")) {
    throw new Error("CONFIG.WEBAPP_URL não definido. Cole a URL do WebApp.");
  }

  // Opção A (recomendado): POST JSON
  // (Seu Apps Script deve implementar doPost(e) e retornar JSON)
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });

  if (!resp.ok) {
    const txt = await resp.text().catch(() => "");
    throw new Error(`Erro HTTP ${resp.status}: ${txt || "sem detalhes"}`);
  }

  const data = await resp.json();
  if (data?.ok === false) {
    throw new Error(data?.error || "Erro retornado pelo WebApp.");
  }
  return data;
}

// =======================
// Transformações
// =======================
function normalizeMatrix(values) {
  // Garante que todas as linhas tenham o mesmo tamanho do cabeçalho
  const headerLen = (values?.[0]?.length || 0);
  return values.map((row) => {
    const r = Array.isArray(row) ? row.slice() : [];
    while (r.length < headerLen) r.push("");
    return r;
  });
}

function parseSheetResponse(apiResponse) {
  // Esperado do WebApp:
  // {
  //   ok: true,
  //   sheetId: "...",
  //   tab: "...",
  //   values: [ [..], [..], ... ]
  // }
  const values = apiResponse?.values;
  if (!Array.isArray(values) || values.length === 0) {
    throw new Error("Resposta sem 'values' ou matriz vazia.");
  }

  const matrix = normalizeMatrix(values);
  const header = matrix[0].map(safeText);

  // meses = todas as colunas do header que parecem mês (jan 2023 etc)
  const meses = header.filter(isProbablyMonthLabel);

  const rows = matrix.slice(1);

  return {
    unidade: "Cajazeiras",
    sheetId: apiResponse.sheetId || CONFIG.SHEET_ID,
    tab: apiResponse.tab || CONFIG.TAB_NAME,
    range: apiResponse.range || CONFIG.RANGE_A1 || "DATA_RANGE",
    header,
    meses,
    rows,
  };
}

function rowsToObjects(header, rows) {
  return rows.map((row) => {
    const obj = {};
    header.forEach((col, idx) => {
      obj[col] = row[idx] ?? "";
    });
    return obj;
  });
}

function buildStructured(result) {
  // Estrutura “ideal” pro KPIs GT:
  // {
  //   unidade, meses, registros (linhas como objeto),
  //   raw: { header, rows, ... }
  // }
  const registros = rowsToObjects(result.header, result.rows);

  return {
    unidade: result.unidade,
    meses: result.meses,
    registros,
    meta: {
      sheetId: result.sheetId,
      tab: result.tab,
      range: result.range,
      totalLinhas: result.rows.length,
      totalColunas: result.header.length,
    },
    raw: {
      header: result.header,
      rows: result.rows,
    },
  };
}

// =======================
// Render
// =======================
function renderMeses(meses) {
  const el = $("meses");
  if (!el) return;

  el.innerHTML = "";
  if (!meses?.length) {
    el.textContent = "Nenhum mês detectado na linha 1.";
    return;
  }

  const ul = document.createElement("ul");
  meses.forEach((m) => {
    const li = document.createElement("li");
    li.textContent = m;
    ul.appendChild(li);
  });

  el.appendChild(ul);
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
    rows[i].forEach((v) => (html += `<td>${escapeHtml(safeText(v))}</td>`));
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

function escapeHtml(str) {
  return safeText(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// =======================
// Actions
// =======================
async function carregarNumerosCajazeiras() {
  setStatus("Carregando planilha (WebApp)...", "info");

  try {
    const payload = {
      action: "read",
      sheetId: CONFIG.SHEET_ID,
      tab: CONFIG.TAB_NAME,
      rangeA1: CONFIG.RANGE_A1, // pode vir vazio
    };

    const api = await callWebApp(payload);
    const parsed = parseSheetResponse(api);
    const structured = buildStructured(parsed);

    LAST_RESULT = structured;

    // Render
    renderMeses(structured.meses);
    renderPreview(structured.raw.header, structured.raw.rows);
    renderJSON(structured);

    setStatus(
      `OK: ${structured.meta.totalLinhas} linhas, ${structured.meta.totalColunas} colunas — ${structured.meta.tab}`,
      "success"
    );
  } catch (err) {
    console.error(err);
    setStatus(`Erro: ${err.message}`, "error");
  }
}

function baixarJSON() {
  if (!LAST_RESULT) {
    setStatus("Nada para baixar ainda. Clique em Carregar primeiro.", "warn");
    return;
  }

  const blob = new Blob([JSON.stringify(LAST_RESULT, null, 2)], {
    type: "application/json",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `kpisgt_${LAST_RESULT.unidade.toLowerCase()}_${new Date()
    .toISOString()
    .slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

// =======================
// Bind UI
// =======================
function bindUI() {
  const btn = $("btnCarregar");
  if (btn) btn.addEventListener("click", carregarNumerosCajazeiras);

  const btnDown = $("btnBaixarJSON");
  if (btnDown) btnDown.addEventListener("click", baixarJSON);
}

// Auto-init
document.addEventListener("DOMContentLoaded", () => {
  bindUI();

  // Se quiser auto-carregar ao abrir, descomente:
  // carregarNumerosCajazeiras();
});
