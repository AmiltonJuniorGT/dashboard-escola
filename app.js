/**
 * KPIs GT — app.js (Opção 2: Apps Script WebApp)
 * Frontend no GitHub Pages chamando WebApp para ler Google Sheets.
 *
 * Planilha base:
 * - Sheet ID: 1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go
 * - Aba: Números Cajazeiras
 * - Meses na linha 1 como texto (jan 2023 ... dez 2025)
 *
 * Requisitos no HTML (recomendado):
 * - IDs: status, meses, preview, json, btnCarregar, btnBaixarJSON
 */

// =======================
// CONFIG (troque só isso)
// =======================
const CONFIG = {
  // ⚠️ COLE A URL DO SEU WEBAPP (Apps Script Deploy > Web app)
  WEBAPP_URL: "COLE_AQUI_A_URL_DO_WEBAPP",

  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  TAB_NAME: "Números Cajazeiras",

  // Se quiser limitar range: "A1:ZZ200" | vazio = backend usa getDataRange()
  RANGE_A1: "",

  // Preview
  PREVIEW_ROWS: 12,

  // JSONP fallback timeout (ms)
  JSONP_TIMEOUT: 12000,
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
  if (!el) {
    console.log(`[STATUS/${type}]`, msg);
    return;
  }
  el.textContent = msg;
  el.dataset.type = type; // opcional p/ CSS
}

function safeText(v) {
  return (v ?? "").toString();
}

function escapeHtml(str) {
  return safeText(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function isProbablyMonthLabel(s) {
  // "jan 2023", "fev 2024", "dez 2025" (critério simples)
  const v = safeText(s).trim().toLowerCase();
  return /^[a-zç]{3}\s\d{4}$/.test(v);
}

// =======================
// Networking (POST first, fallback JSONP)
// =======================
async function callWebApp(payload) {
  const url = CONFIG.WEBAPP_URL;

  if (!url || url.includes("COLE_AQUI")) {
    throw new Error("CONFIG.WEBAPP_URL não definido. Cole a URL do WebApp do Apps Script.");
  }

  // 1) Tenta POST (fetch). Se CORS bloquear, cai no JSONP automaticamente.
  try {
    const resp = await fetch(url, {
      method: "POST",
      // Apps Script lê body como texto; usar text/plain evita preflight em alguns cenários
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });

    if (!resp.ok) {
      const txt = await resp.text().catch(() => "");
      throw new Error(`Erro HTTP ${resp.status}: ${txt || "sem detalhes"}`);
    }

    const data = await resp.json();
    if (data?.ok === false) throw new Error(data?.error || "Erro retornado pelo WebApp.");
    return data;
  } catch (err) {
    // 2) Fallback JSONP (GET)
    console.warn("POST falhou (provável CORS). Tentando JSONP...", err);
    return await callWebAppJSONP(payload);
  }
}

function buildQuery(params) {
  const usp = new URLSearchParams();
  Object.entries(params).forEach(([k, v]) => {
    if (v === undefined || v === null) return;
    if (typeof v === "string" && v.trim() === "") return;
    usp.set(k, String(v));
  });
  return usp.toString();
}

function callWebAppJSONP(payload) {
  return new Promise((resolve, reject) => {
    const url = CONFIG.WEBAPP_URL;
    const cbName = `__kpisgt_cb_${Date.now()}_${Math.floor(Math.random() * 1e6)}`;

    const cleanup = () => {
      try {
        delete window[cbName];
      } catch (_) {}
      if (script && script.parentNode) script.parentNode.removeChild(script);
      if (timer) clearTimeout(timer);
    };

    window[cbName] = (data) => {
      cleanup();
      if (data?.ok === false) return reject(new Error(data?.error || "Erro retornado pelo WebApp (JSONP)."));
      resolve(data);
    };

    const params = {
      // seu Code.gs deve aceitar action/sheetId/tab/rangeA1 e callback
      action: payload.action,
      sheetId: payload.sheetId,
      tab: payload.tab,
      rangeA1: payload.rangeA1 || "",
      callback: cbName,
    };

    const qs = buildQuery(params);
    const fullUrl = `${url}?${qs}`;

    const script = document.createElement("script");
    script.src = fullUrl;
    script.async = true;

    script.onerror = () => {
      cleanup();
      reject(new Error("Falha ao carregar JSONP (script.onerror). Verifique deploy do WebApp e parâmetros."));
    };

    const timer = setTimeout(() => {
      cleanup();
      reject(new Error("Timeout no JSONP. Verifique se o WebApp está publicado e acessível."));
    }, CONFIG.JSONP_TIMEOUT);

    document.head.appendChild(script);
  });
}

// =======================
// Transformações
// =======================
function normalizeMatrix(values) {
  const headerLen = (values?.[0]?.length || 0);
  return values.map((row) => {
    const r = Array.isArray(row) ? row.slice() : [];
    while (r.length < headerLen) r.push("");
    return r;
  });
}

function parseSheetResponse(apiResponse) {
  const values = apiResponse?.values;
  if (!Array.isArray(values) || values.length === 0) {
    throw new Error("Resposta sem 'values' ou matriz vazia.");
  }

  const matrix = normalizeMatrix(values);
  const header = matrix[0].map(safeText);

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

function buildStructured(parsed) {
  const registros = rowsToObjects(parsed.header, parsed.rows);

  return {
    unidade: parsed.unidade,
    meses: parsed.meses,
    registros,
    meta: {
      sheetId: parsed.sheetId,
      tab: parsed.tab,
      range: parsed.range,
      totalLinhas: parsed.rows.length,
      totalColunas: parsed.header.length,
    },
    raw: {
      header: parsed.header,
      rows: parsed.rows,
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

// =======================
// Actions
// =======================
async function carregarNumerosCajazeiras() {
  setStatus("Carregando (WebApp)...", "info");

  try {
    const payload = {
      action: "read",
      sheetId: CONFIG.SHEET_ID,
      tab: CONFIG.TAB_NAME,
      rangeA1: CONFIG.RANGE_A1,
    };

    const api = await callWebApp(payload);
    const parsed = parseSheetResponse(api);
    const structured = buildStructured(parsed);

    LAST_RESULT = structured;

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

document.addEventListener("DOMContentLoaded", () => {
  bindUI();
  // Se quiser auto-carregar:
  // carregarNumerosCajazeiras();
});

// Expor funções (opcional) para usar via onclick no HTML
window.KPIsGT = {
  carregarNumerosCajazeiras,
  baixarJSON,
};
