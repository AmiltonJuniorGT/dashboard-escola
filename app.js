// Dashboard – Gestão da Escola (Fechamento 3 Unidades)
// Versão: 2026-02-11 (corrige meses como DATE no cabeçalho)

const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS_UNIDADES = {
  Cajazeiras: "Números Cajazeiras",
  Camaçari: "Números Camaçari",
  "São Cristóvão": "Números São Cristóvão",
};

const state = {
  unidade: "TODAS",      // TODAS | Cajazeiras | Camaçari | São Cristóvão
  marca: "C",            // C (Consolidado) | T (Grau Técnico) | P (Profissionalizante)
  monthKey: null,        // {y,m}
  months: [],            // [{y,m,label, colIndex}]
};

const $ = (id) => document.getElementById(id);

// ---------- GViz helper ----------
async function fetchSheetTable(sheetId, tabName) {
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    encodeURIComponent(sheetId) +
    "/gviz/tq?gid=&sheet=" +
    encodeURIComponent(tabName) +
    "&tq=" +
    encodeURIComponent("select *");
  const res = await fetch(url, { cache: "no-store" });
  const txt = await res.text();
  const m = txt.match(/google\.visualization\.Query\.setResponse\(([\s\S]+)\);/);
  if (!m) throw new Error("GViz inválido (planilha privada ou aba inexistente). Aba: " + tabName);
  const obj = JSON.parse(m[1]);
  if (obj.status !== "ok") throw new Error("GViz status != ok: " + JSON.stringify(obj));
  const cols = obj.table.cols.map((c) => c.label || c.id || "");
  const rows = obj.table.rows.map((r) => r.c.map((cell) => (cell ? cell.v : null)));
  return { cols, rows, raw: obj };
}

function isDateLike(v) {
  return v instanceof Date ||
    (typeof v === "string" && /^Date\(\d+,\d+,\d+\)$/.test(v)) ||
    (typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) ||
    (typeof v === "string" && /^\d{2}\/\d{2}\/\d{4}$/.test(v));
}

function parseMonthKey(v) {
  if (v == null) return null;

  // Date object (rare in browser; GViz usually returns "Date(YYYY,M,D)")
  if (v instanceof Date) {
    return { y: v.getFullYear(), m: v.getMonth() + 1 };
  }

  if (typeof v === "string") {
    let s = v.trim();

    // GViz date string: "Date(2025,0,23)" => Jan/2025
    let md = s.match(/^Date\((\d{4}),(\d{1,2}),(\d{1,2})\)$/);
    if (md) return { y: Number(md[1]), m: Number(md[2]) + 1 };

    // ISO date "2025-01-23"
    md = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (md) return { y: Number(md[1]), m: Number(md[2]) };

    // BR date "23/01/2025"
    md = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (md) return { y: Number(md[3]), m: Number(md[2]) };

    // Legacy: "janeiro.23" / "jan 23" / etc.
    const sm = stripAccents(s.toLowerCase());
    const mm = sm.match(/^([a-z]+)[\.\s\/-]?(\d{2}|\d{4})$/i);
    if (mm) {
      const mon = monthNameToNumber(mm[1]);
      if (!mon) return null;
      const yy = mm[2].length === 2 ? 2000 + Number(mm[2]) : Number(mm[2]);
      return { y: yy, m: mon };
    }
  }

  return null;
}

function monthNameToNumber(name) {
  const n = stripAccents(String(name || "").toLowerCase());
  const map = {
    janeiro: 1, jan: 1,
    fevereiro: 2, fev: 2,
    marco: 3, mar: 3,
    abril: 4, abr: 4,
    maio: 5,
    junho: 6, jun: 6,
    julho: 7, jul: 7,
    agosto: 8, ago: 8,
    setembro: 9, set: 9,
    outubro: 10, out: 10,
    novembro: 11, nov: 11,
    dezembro: 12, dez: 12,
  };
  return map[n] || null;
}

function stripAccents(s) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function keyToId(k) {
  return `${k.y}-${String(k.m).padStart(2, "0")}`;
}

function keyToLabel(k) {
  return `${String(k.m).padStart(2, "0")}/${k.y}`;
}

function addMonths(k, delta) {
  let y = k.y;
  let m = k.m + delta;
  while (m <= 0) { m += 12; y -= 1; }
  while (m > 12) { m -= 12; y += 1; }
  return { y, m };
}

// ---------- Parsing sections ----------
function findSections(rows) {
  // returns { tecnico: {start,end}, prof: {start,end}, grauEdu: {start,end} }
  const sec = {};
  const hdrRe = /^PLANILHA DE INDICADORES/i;
  const geRe = /^GRAU EDUCACIONAL/i;

  const starts = [];
  for (let i = 0; i < rows.length; i++) {
    const a = rows[i]?.[0];
    if (typeof a === "string" && hdrRe.test(a)) starts.push({ i, t: a });
    if (typeof a === "string" && geRe.test(a)) starts.push({ i, t: a });
  }
  starts.sort((x, y) => x.i - y.i);

  function endOf(idx) {
    const pos = starts.findIndex((s) => s.i === idx);
    const next = starts[pos + 1];
    return next ? next.i - 1 : rows.length - 1;
  }

  for (const s of starts) {
    const t = stripAccents(String(s.t).toLowerCase());
    if (t.includes("tecnico")) sec.tecnico = { start: s.i + 1, end: endOf(s.i) };
    else if (t.includes("profissionalizante")) sec.prof = { start: s.i + 1, end: endOf(s.i) };
    else if (t.startsWith("grau educacional")) sec.grauEdu = { start: s.i + 1, end: endOf(s.i) };
  }
  return sec;
}

function findRowValue(rows, start, end, labelRegex, colIndex) {
  for (let i = start; i <= end; i++) {
    const a = rows[i]?.[0];
    if (typeof a !== "string") continue;
    if (labelRegex.test(stripAccents(a).toUpperCase())) {
      const v = rows[i]?.[colIndex];
      return v == null ? null : v;
    }
  }
  return null;
}

function asNumber(v) {
  if (v == null) return null;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const s = v.replace(/\./g, "").replace(",", ".").replace(/[^\d\.\-]/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function fmt(v, kind) {
  if (v == null || !Number.isFinite(v)) return "—";
  if (kind === "int") return Math.round(v).toLocaleString("pt-BR");
  if (kind === "pct") return (v * 1).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
  // money
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function deltaCell(delta, kind) {
  if (delta == null || !Number.isFinite(delta)) return `<span class="muted">—</span>`;
  const up = delta > 0;
  const cls = up ? "up" : (delta < 0 ? "down" : "flat");
  const arrow = up ? "▲" : (delta < 0 ? "▼" : "•");
  let txt = "";
  if (kind === "int") txt = Math.round(delta).toLocaleString("pt-BR");
  else if (kind === "pct") txt = delta.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " p.p.";
  else txt = delta.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  return `<span class="${cls}">${arrow} ${txt}</span>`;
}

// ---------- Month discovery ----------
function discoverMonths(headerRow) {
  const months = [];
  for (let c = 1; c < headerRow.length; c++) {
    const k = parseMonthKey(headerRow[c]);
    if (!k) continue;
    months.push({ ...k, label: keyToLabel(k), colIndex: c });
  }
  // unique by y-m (some sheets repeat)
  const map = new Map();
  for (const m of months) map.set(keyToId(m), m);
  const arr = Array.from(map.values());
  arr.sort((a, b) => (a.y - b.y) || (a.m - b.m));
  return arr;
}

function pickDefaultMonth(months) {
  if (!months.length) return null;
  return { y: months[months.length - 1].y, m: months[months.length - 1].m };
}

function monthKeyFromInputs() {
  // using start date input (YYYY-MM-DD)
  const d = $("f_start").value;
  if (!d) return null;
  const m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return { y: Number(m[1]), m: Number(m[2]) };
}

function setInputsFromMonthKey(k) {
  if (!k) return;
  const start = new Date(Date.UTC(k.y, k.m - 1, 1));
  const end = new Date(Date.UTC(k.y, k.m, 0));
  $("f_start").value = start.toISOString().slice(0, 10);
  $("f_end").value = end.toISOString().slice(0, 10);
}

// ---------- Rendering ----------
function buildMetricRows(kind) {
  // kind: "T" or "P" or "C"
  if (kind === "T") {
    return [
      { label: "Faturamento (R$)", re: /FATURAMENTO/i, kind: "money" },
      { label: "Custo Operacional (R$)", re: /^CUSTO OPERACIONAL R\$/i, kind: "money" },
      { label: "Custo Total (R$)", re: /^CUSTO TOTAL/i, kind: "money" },
      { label: "Ticket Médio (R$)", re: /TICKET/i, kind: "money" },
      { label: "Número de Ativos (Un.)", re: /ALUNOS ATIVOS/i, kind: "int" },
      { label: "Matrículas (Un.)", re: /MATRICULAS REALIZADAS/i, kind: "int" },
      { label: "Meta (Un.)", re: /META DE MATRICULA/i, kind: "int" },
      { label: "Evasão Real (Un.)", re: /EVASAO REAL/i, kind: "int" },
      { label: "Inadimplência (%)", re: /INADIMPLENCIA/i, kind: "pct" },
    ];
  }
  if (kind === "P") {
    return [
      { label: "Faturamento (R$)", re: /FATURAMENTO/i, kind: "money" },
      { label: "Custo Operacional (R$)", re: /^CUSTO OPERACIONAL R\$/i, kind: "money" },
      { label: "Custo Total (R$)", re: /^CUSTO TOTAL/i, kind: "money" },
      { label: "Ticket Médio (R$)", re: /TICKET/i, kind: "money" },
      { label: "Número de Ativos (Un.)", re: /ALUNOS ATIVOS/i, kind: "int" },
      { label: "Matrículas (Un.)", re: /MATRICULAS REALIZADAS/i, kind: "int" },
      { label: "Meta (Un.)", re: /META DE MATRICULA/i, kind: "int" },
      { label: "Evasão Real (Un.)", re: /EVASAO REAL/i, kind: "int" },
      { label: "Inadimplência (%)", re: /INADIMPLENCIA/i, kind: "pct" },
    ];
  }
  // Consolidado = soma GT+GP para dinheiro e int; ticket médio fica média ponderada simples (não perfeito) e % média simples
  return [
    { label: "Faturamento (R$)", kind: "money", calc: "sum", key: "faturamento" },
    { label: "Custo Operacional (R$)", kind: "money", calc: "sum", key: "custo_op" },
    { label: "Custo Total (R$)", kind: "money", calc: "sum", key: "custo_total" },
    { label: "Ticket Médio (R$)", kind: "money", calc: "avg", key: "ticket" },
    { label: "Número de Ativos (Un.)", kind: "int", calc: "sum", key: "ativos" },
    { label: "Matrículas (Un.)", kind: "int", calc: "sum", key: "matriculas" },
    { label: "Meta (Un.)", kind: "int", calc: "sum", key: "meta" },
    { label: "Evasão Real (Un.)", kind: "int", calc: "sum", key: "evasao" },
    { label: "Inadimplência (%)", kind: "pct", calc: "avg", key: "inad" },
  ];
}

function extractBlock(rows, sec, monthCol, kind) {
  const defs = buildMetricRows(kind);
  const out = {};
  for (const d of defs) {
    const v = findRowValue(rows, sec.start, sec.end, d.re || /$^/, monthCol);
    out[d.label] = asNumber(v);
  }
  return out;
}

function mergeConsolidado(gt, gp) {
  const o = {};
  const map = {
    "Faturamento (R$)": "sum",
    "Custo Operacional (R$)": "sum",
    "Custo Total (R$)": "sum",
    "Ticket Médio (R$)": "avg",
    "Número de Ativos (Un.)": "sum",
    "Matrículas (Un.)": "sum",
    "Meta (Un.)": "sum",
    "Evasão Real (Un.)": "sum",
    "Inadimplência (%)": "avg",
  };
  for (const k of Object.keys(map)) {
    const a = gt[k], b = gp[k];
    if (a == null && b == null) o[k] = null;
    else if (map[k] === "sum") o[k] = (a || 0) + (b || 0);
    else { // avg
      const vals = [a, b].filter((x) => x != null && Number.isFinite(x));
      o[k] = vals.length ? (vals.reduce((s, x) => s + x, 0) / vals.length) : null;
    }
  }
  return o;
}

function renderComparativo(title, current, prev, yoy, prev3, kind) {
  const defs = buildMetricRows(kind === "C" ? "C" : "T"); // labels same
  const rowsHtml = defs.map((d) => {
    const lab = d.label;
    const v0 = current[lab];
    const vPrev = prev ? prev[lab] : null;
    const vYoy = yoy ? yoy[lab] : null;
    const deltaPrev = (v0 != null && vPrev != null) ? (v0 - vPrev) : null;
    const deltaYoy = (v0 != null && vYoy != null) ? (v0 - vYoy) : null;

    const cellsPrev3 = prev3.map((blk) => `<td class="num">${fmt(blk[lab], d.kind)}</td>`).join("");
    return `
      <tr>
        <td class="metric">${lab}</td>
        <td class="num strong">${fmt(v0, d.kind)}</td>
        <td class="delta">${deltaCell(deltaPrev, d.kind)}</td>
        <td class="delta">${deltaCell(deltaYoy, d.kind)}</td>
        ${cellsPrev3}
        <td class="num">${fmt(vYoy, d.kind)}</td>
      </tr>`;
  }).join("");

  const colsPrev3 = prev3.map((m) => `<th>${m.__label}</th>`).join("");

  return `
  <div class="card">
    <h3>${title}</h3>
    <div class="hint">Mês selecionado: <b>${current.__label}</b></div>
    <div class="tableWrap">
      <table class="tbl">
        <thead>
          <tr>
            <th style="min-width:240px">Indicador</th>
            <th>${current.__label}</th>
            <th>Variação (mês anterior)</th>
            <th>Variação (ano anterior)</th>
            ${colsPrev3}
            <th>${yoy ? yoy.__label : "Ano anterior"}</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  </div>`;
}

async function loadUnit(unidadeNome) {
  const tab = TABS_UNIDADES[unidadeNome];
  const { rows } = await fetchSheetTable(DATA_SHEET_ID, tab);

  if (!rows.length) throw new Error("Aba vazia: " + tab);

  // months are in first row (row 0), columns 1+
  const months = discoverMonths(rows[0] || []);
  if (!months.length) {
    throw new Error("Não encontrei colunas de mês no cabeçalho. Verifique se os meses estão na 1ª linha (ex: janeiro.23 ou data).");
  }

  // store global months once
  if (!state.months.length) state.months = months;

  // choose current monthKey
  if (!state.monthKey) state.monthKey = pickDefaultMonth(months);

  // find month column index
  function colFor(k) {
    const found = months.find((m) => m.y === k.y && m.m === k.m);
    return found ? found.colIndex : null;
  }

  const mk = state.monthKey;
  const col0 = colFor(mk);
  const colPrev = colFor(addMonths(mk, -1));
  const colYoy = colFor({ y: mk.y - 1, m: mk.m });

  const prev3Keys = [addMonths(mk, -1), addMonths(mk, -2), addMonths(mk, -3)];
  const prev3Cols = prev3Keys.map(colFor);

  const sec = findSections(rows);

  function blockFor(kind, col) {
    if (!col) return null;
    if (kind === "T") return { ...extractBlock(rows, sec.tecnico, col, "T") };
    if (kind === "P") return { ...extractBlock(rows, sec.prof, col, "P") };
    return null;
  }

  function withLabel(obj, k) {
    if (!obj) return null;
    obj.__label = keyToLabel(k);
    return obj;
  }

  // current
  const gt0 = withLabel(blockFor("T", col0), mk);
  const gp0 = withLabel(blockFor("P", col0), mk);
  const c0 = withLabel(mergeConsolidado(gt0 || {}, gp0 || {}), mk);

  // prev / yoy
  const prevKey = addMonths(mk, -1);
  const yoyKey = { y: mk.y - 1, m: mk.m };

  const gtPrev = colPrev ? withLabel(blockFor("T", colPrev), prevKey) : null;
  const gpPrev = colPrev ? withLabel(blockFor("P", colPrev), prevKey) : null;
  const cPrev = colPrev ? withLabel(mergeConsolidado(gtPrev || {}, gpPrev || {}), prevKey) : null;

  const gtYoy = colYoy ? withLabel(blockFor("T", colYoy), yoyKey) : null;
  const gpYoy = colYoy ? withLabel(blockFor("P", colYoy), yoyKey) : null;
  const cYoy = colYoy ? withLabel(mergeConsolidado(gtYoy || {}, gpYoy || {}), yoyKey) : null;

  // prev3
  const prev3 = [];
  for (let i = 0; i < 3; i++) {
    const k = prev3Keys[i];
    const col = prev3Cols[i];
    if (!col) {
      prev3.push({ __label: keyToLabel(k) });
      continue;
    }
    const gt = blockFor("T", col);
    const gp = blockFor("P", col);
    const cc = mergeConsolidado(gt || {}, gp || {});
    cc.__label = keyToLabel(k);
    prev3.push(cc);
  }

  return { unidadeNome, gt0, gp0, c0, gtPrev, gpPrev, cPrev, gtYoy, gpYoy, cYoy, prev3 };
}

function renderFechamento(resultList) {
  const wrap = $("tab_fechamento");
  wrap.innerHTML = "";

  for (const r of resultList) {
    const header = `<h2>Fechamento – ${r.unidadeNome}</h2>`;
    const kind = state.marca; // T/P/C

    let html = header;

    if (kind === "T") {
      html += renderComparativo("Grau Técnico", r.gt0, r.gtPrev, r.gtYoy, r.prev3.map((x) => x), "T");
    } else if (kind === "P") {
      html += renderComparativo("Profissionalizante", r.gp0, r.gpPrev, r.gpYoy, r.prev3.map((x) => x), "P");
    } else {
      html += renderComparativo("Consolidado da Unidade (T + P)", r.c0, r.cPrev, r.cYoy, r.prev3.map((x) => x), "C");
      html += `<div class="grid2">
        ${renderComparativo("Grau Técnico", r.gt0, r.gtPrev, r.gtYoy, r.prev3.map((x) => x), "T")}
        ${renderComparativo("Profissionalizante", r.gp0, r.gpPrev, r.gpYoy, r.prev3.map((x) => x), "P")}
      </div>`;
    }

    const block = document.createElement("div");
    block.className = "unitBlock";
    block.innerHTML = html;
    wrap.appendChild(block);
  }
}

function renderStatus(msg, ok = true) {
  const el = $("status");
  el.textContent = msg;
  el.className = ok ? "status ok" : "status err";
}

function showTab(tabId) {
  document.querySelectorAll(".tabPanel").forEach((el) => (el.style.display = "none"));
  $(tabId).style.display = "block";

  document.querySelectorAll(".tabBtn").forEach((b) => b.classList.remove("active"));
  document.querySelector(`[data-tab="${tabId}"]`)?.classList.add("active");
}

// ---------- Init ----------
function fillUnitSelect() {
  const sel = $("f_unidade");
  sel.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "TODAS";
  optAll.textContent = "Todas";
  sel.appendChild(optAll);

  for (const k of Object.keys(TABS_UNIDADES)) {
    const o = document.createElement("option");
    o.value = k;
    o.textContent = k;
    sel.appendChild(o);
  }
  sel.value = state.unidade;
}

function fillMarcaSelect() {
  const sel = $("f_marca");
  sel.innerHTML = "";
  const opts = [
    { v: "C", t: "Consolidado" },
    { v: "T", t: "Grau Técnico (T)" },
    { v: "P", t: "Profissionalizante (P)" },
  ];
  for (const x of opts) {
    const o = document.createElement("option");
    o.value = x.v;
    o.textContent = x.t;
    sel.appendChild(o);
  }
  sel.value = state.marca;
}

async function applyFilters() {
  try {
    renderStatus("Carregando dados...", true);

    // monthKey from inputs if user changed
    const mk = monthKeyFromInputs();
    if (mk) state.monthKey = mk;

    state.unidade = $("f_unidade").value;
    state.marca = $("f_marca").value;

    const units = state.unidade === "TODAS" ? Object.keys(TABS_UNIDADES) : [state.unidade];
    const results = [];
    for (const u of units) results.push(await loadUnit(u));

    // if monthKey was null, set inputs
    if (state.monthKey) setInputsFromMonthKey(state.monthKey);

    renderFechamento(results);
    renderStatus("Atualizado com sucesso ✅", true);
  } catch (e) {
    console.error(e);
    $("tab_fechamento").innerHTML = `<div class="card errBox">${String(e.message || e)}</div>`;
    renderStatus("Erro ao atualizar ❌", false);
  }
}

function bindUI() {
  document.querySelectorAll(".tabBtn").forEach((btn) => {
    btn.addEventListener("click", () => showTab(btn.getAttribute("data-tab")));
  });

  $("btn_apply").addEventListener("click", applyFilters);
  $("btn_default").addEventListener("click", async () => {
    state.months = [];
    state.monthKey = null;
    await applyFilters();
  });
}

async function boot() {
  fillUnitSelect();
  fillMarcaSelect();
  bindUI();

  // default tab: Fechamento
  showTab("tab_fechamento");
  await applyFilters();
}

document.addEventListener("DOMContentLoaded", boot);
