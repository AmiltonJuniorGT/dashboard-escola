// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS_UNIDADES = {
  Cajazeiras: "Números Cajazeiras",
  Camaçari: "Números Camaçari",
  "São Cristóvão": "Números Sao Cristovao",
};

const state = {
  unidade: "TODAS",
  monthKey: null, // {y: 2025, m: 12}
};

const $ = (id) => document.getElementById(id);

// ================= Helpers =================
function stripAccents(s) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}
function norm(s) {
  return stripAccents(String(s ?? "")).trim().toUpperCase();
}

function parseBR(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  const s0 = String(v).trim();
  if (!s0) return null;

  const s = s0
    .replace(/R\$\s?/i, "")
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(",", ".")
    .replace("%", "");

  const n = Number(s);
  return Number.isNaN(n) ? null : n;
}

function fmtCurrency(n) {
  if (n === null || n === undefined) return "—";
  try {
    return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  } catch {
    return `R$ ${n}`;
  }
}
function fmtInt(n) {
  if (n === null || n === undefined) return "—";
  return Math.round(n).toLocaleString("pt-BR");
}
function fmtPct(n) {
  if (n === null || n === undefined) return "—";
  return `${n.toLocaleString("pt-BR", { maximumFractionDigits: 2 })}%`;
}
function fmtValue(n, type) {
  if (type === "currency") return fmtCurrency(n);
  if (type === "int") return fmtInt(n);
  if (type === "pct") return fmtPct(n);
  return n ?? "—";
}
function fmtDelta(n, type) {
  if (n === null || n === undefined) return "—";
  const up = n >= 0;
  const arrow = up ? "↑" : "↓";
  const cls = up ? "delta up" : "delta down";
  return `<span class="${cls}">${arrow} ${fmtValue(n, type)}</span>`;
}

// ================= GViz fetch =================
async function fetchGviz(sheetId, tab) {
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?sheet=${encodeURIComponent(tab)}`;
  const res = await fetch(url);
  const txt = await res.text();
  const json = JSON.parse(txt.substring(txt.indexOf("{"), txt.lastIndexOf("}") + 1));

  const cols = json.table.cols.map((c) => (c.label ?? "").trim());
  const rows = json.table.rows.map((r) => r.c.map((cell) => (cell ? cell.v : null)));

  return { cols, rows };
}

// ================= Month label parsing =================
const PT_MONTHS = {
  JANEIRO: 1, JAN: 1,
  FEVEREIRO: 2, FEV: 2,
  MARCO: 3, MARÇO: 3, MAR: 3,
  ABRIL: 4, ABR: 4,
  MAIO: 5, MAI: 5,
  JUNHO: 6, JUN: 6,
  JULHO: 7, JUL: 7,
  AGOSTO: 8, AGO: 8,
  SETEMBRO: 9, SET: 9,
  OUTUBRO: 10, OUT: 10,
  NOVEMBRO: 11, NOV: 11,
  DEZEMBRO: 12, DEZ: 12,
};

function parseMonthLabel(label) {
  const raw = String(label ?? "").trim();
  if (!raw) return null;

  const cleaned = stripAccents(raw).toUpperCase();
  // aceita: "janeiro.23" "jan.23" "janeiro 23" "janeiro/23"
  const m = cleaned.match(/^([A-ZÇ]+)[\.\s\/-]?(\d{2}|\d{4})$/);
  if (!m) return null;

  const name = m[1];
  const yy = m[2];
  const month = PT_MONTHS[name];
  if (!month) return null;

  const year = yy.length === 4 ? Number(yy) : Number("20" + yy);
  if (!year || year < 2000 || year > 2100) return null;

  return { y: year, m: month };
}

function monthToKeyNum(mk) {
  return mk.y * 100 + mk.m;
}
function monthKeyToLabel(mk) {
  const names = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  return `${names[mk.m]}.${String(mk.y).slice(-2)}`;
}
function shiftMonth(mk, delta) {
  let y = mk.y;
  let m = mk.m + delta;
  while (m <= 0) { m += 12; y -= 1; }
  while (m > 12) { m -= 12; y += 1; }
  return { y, m };
}

// ================= Section ranges =================
function findSectionRange(rows, headerNeedle) {
  const needle = norm(headerNeedle);
  let start = -1;

  for (let i = 0; i < rows.length; i++) {
    const a = norm(rows[i][0]);
    if (a.includes(needle)) { start = i + 1; break; }
  }
  if (start < 0) return null;

  let end = rows.length;
  for (let i = start; i < rows.length; i++) {
    const a = norm(rows[i][0]);
    // próximo bloco grande / quebra
    if (a.includes("PLANILHA DE INDICADORES") && i > start) { end = i; break; }
    if (a.includes("FACULDADE")) { end = i; break; }
    if (a.includes("GRAU EDUCACIONAL") && i > start) { end = i; break; }
  }
  return { start, end };
}

function findRowIndexInRange(rows, range, needles) {
  const ns = needles.map(norm);
  for (let i = range.start; i < range.end; i++) {
    const a = norm(rows[i][0]);
    if (!a) continue;
    for (const n of ns) {
      if (a === n || a.includes(n)) return i;
    }
  }
  return -1;
}

// ================= Métricas =================
const METRICS = [
  { label: "Faturamento", type: "currency", needles: ["FATURAMENTO"] },
  { label: "Custo Operacional", type: "currency", needles: ["CUSTO OPERACIONAL"] },
  { label: "Custo Total", type: "currency", needles: ["CUSTO TOTAL"] },
  { label: "Ticket Médio", type: "currency", needles: ["TICKET MEDIO", "TICKET MÉDIO"] },
  { label: "Ativos", type: "int", needles: ["ALUNOS ATIVOS", "ATIVOS", "NUMERO DE ATIVOS", "NÚMERO DE ATIVOS"] },
  { label: "Matrículas", type: "int", needles: ["MATRICULAS REALIZADAS", "MATRÍCULAS REALIZADAS", "MATRICULAS"] },
  { label: "Meta", type: "int", needles: ["META DE MATRICULA", "META DE MATRÍCULA", "META"] },
  { label: "Evasão Real", type: "int", needles: ["EVASAO REAL", "EVASÃO REAL"] },
  { label: "Inadimplência (%)", type: "pct", needles: ["INADIMPLENCIA", "INADIMPLÊNCIA"] },
];

const COLS = [
  { id: "cur", label: "Mês atual" },
  { id: "vm", label: "Δ M-1" },
  { id: "vy", label: "Δ Ano-1" },
  { id: "m1", label: "M-1" },
  { id: "m2", label: "M-2" },
  { id: "m3", label: "M-3" },
  { id: "y1", label: "Ano-1" },
];

// ================= Month col mapping =================
function buildMonthColumns(cols) {
  const monthCols = [];
  for (let i = 0; i < cols.length; i++) {
    const mk = parseMonthLabel(cols[i]);
    if (mk) monthCols.push({ idx: i, mk, keyNum: monthToKeyNum(mk), label: cols[i] });
  }
  monthCols.sort((a, b) => a.keyNum - b.keyNum);
  const byKeyNum = new Map(monthCols.map((c) => [c.keyNum, c.idx]));
  return { monthCols, byKeyNum };
}

function hasAnyDataInColumn(rows, range, colIdx) {
  for (let i = range.start; i < range.end; i++) {
    const v = rows[i][colIdx];
    if (v !== null && v !== undefined && String(v).trim() !== "") return true;
  }
  return false;
}

function pickLatestMonthKey(monthCols, rows, rangesToCheck) {
  // pega o mês mais recente que tenha algum dado em pelo menos 1 seção
  for (let i = monthCols.length - 1; i >= 0; i--) {
    const colIdx = monthCols[i].idx;
    for (const rg of rangesToCheck) {
      if (rg && hasAnyDataInColumn(rows, rg, colIdx)) return monthCols[i].mk;
    }
  }
  // fallback: último header
  return monthCols.length ? monthCols[monthCols.length - 1].mk : null;
}

// ================= Extract values =================
function getMetricValue(rows, range, metric, colIdx) {
  const rowIdx = findRowIndexInRange(rows, range, metric.needles);
  if (rowIdx < 0) return null;
  return parseBR(rows[rowIdx][colIdx]);
}

function buildTableRows(rows, range, colMap, mkCur) {
  const mkM1 = shiftMonth(mkCur, -1);
  const mkM2 = shiftMonth(mkCur, -2);
  const mkM3 = shiftMonth(mkCur, -3);
  const mkY1 = shiftMonth(mkCur, -12);

  const idxCur = colMap.get(monthToKeyNum(mkCur));
  const idxM1  = colMap.get(monthToKeyNum(mkM1));
  const idxM2  = colMap.get(monthToKeyNum(mkM2));
  const idxM3  = colMap.get(monthToKeyNum(mkM3));
  const idxY1  = colMap.get(monthToKeyNum(mkY1));

  return METRICS.map((m) => {
    const cur = idxCur != null ? getMetricValue(rows, range, m, idxCur) : null;
    const m1  = idxM1  != null ? getMetricValue(rows, range, m, idxM1)  : null;
    const m2  = idxM2  != null ? getMetricValue(rows, range, m, idxM2)  : null;
    const m3  = idxM3  != null ? getMetricValue(rows, range, m, idxM3)  : null;
    const y1  = idxY1  != null ? getMetricValue(rows, range, m, idxY1)  : null;

    const vm = (cur != null && m1 != null) ? (cur - m1) : null;
    const vy = (cur != null && y1 != null) ? (cur - y1) : null;

    return { metric: m, values: { cur, vm, vy, m1, m2, m3, y1 } };
  });
}

// ================= Render HTML =================
function ensureStylesOnce() {
  if (document.getElementById("fechamento-inline-style")) return;
  const style = document.createElement("style");
  style.id = "fechamento-inline-style";
  style.innerHTML = `
    .fech-topbar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:10px 0 14px 0}
    .fech-topbar select{padding:6px 8px;border-radius:8px;border:1px solid #cfd8cf}
    .fech-card{background:#fff;border:1px solid #dfe7df;border-radius:14px;padding:12px 14px;margin:12px 0}
    .fech-h2{margin:0 0 6px 0}
    .fech-muted{opacity:.75;font-size:12px;margin:0 0 10px 0}
    .fech-table-wrap{overflow:auto}
    table.fech-table{width:100%;border-collapse:collapse}
    table.fech-table th, table.fech-table td{border-bottom:1px solid #e9efe9;padding:8px 10px;text-align:right;white-space:nowrap}
    table.fech-table th:first-child, table.fech-table td:first-child{text-align:left}
    .delta{font-weight:600}
    .delta.up{color:#1b7f3a}
    .delta.down{color:#b3261e}
    .block-title{font-weight:700;margin:0 0 8px 0}
  `;
  document.head.appendChild(style);
}

function renderOneBlock(title, rowsBuilt) {
  const thead = `
    <thead>
      <tr>
        <th>Indicador</th>
        ${COLS.map((c) => `<th>${c.label}</th>`).join("")}
      </tr>
    </thead>
  `;

  const tbody = `
    <tbody>
      ${rowsBuilt.map((r) => {
        const t = r.metric.type;
        const v = r.values;
        return `
          <tr>
            <td>${r.metric.label}</td>
            <td>${fmtValue(v.cur, t)}</td>
            <td>${fmtDelta(v.vm, t)}</td>
            <td>${fmtDelta(v.vy, t)}</td>
            <td>${fmtValue(v.m1, t)}</td>
            <td>${fmtValue(v.m2, t)}</td>
            <td>${fmtValue(v.m3, t)}</td>
            <td>${fmtValue(v.y1, t)}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  return `
    <div class="fech-card">
      <div class="block-title">${title}</div>
      <div class="fech-table-wrap">
        <table class="fech-table">
          ${thead}
          ${tbody}
        </table>
      </div>
    </div>
  `;
}

// ================= Main render (4ª aba) =================
async function renderFechamento() {
  ensureStylesOnce();

  const container = $("fechamentoView");
  if (!container) return;
  container.innerHTML = "Carregando…";

  const unidades = state.unidade === "TODAS"
    ? Object.keys(TABS_UNIDADES)
    : [state.unidade];

  let html = "";

  // dropdown de mês será montado a partir da primeira unidade carregada
  let globalMonthCols = null;
  let globalMonthByKey = null;
  let globalMonthOptions = null;

  for (const u of unidades) {
    const tab = TABS_UNIDADES[u];

    const { cols, rows } = await fetchGviz(DATA_SHEET_ID, tab);
    const { monthCols, byKeyNum } = buildMonthColumns(cols);

    const rgTec = findSectionRange(rows, "TECNICO");
    const rgPro = findSectionRange(rows, "PROFISSIONALIZANTE");

    // define mês default (mais recente com dados)
    if (!state.monthKey) {
      state.monthKey = pickLatestMonthKey(monthCols, rows, [rgTec, rgPro]);
    }

    // salva opções globais (uma vez)
    if (!globalMonthCols) {
      globalMonthCols = monthCols;
      globalMonthByKey = byKeyNum;
      globalMonthOptions = monthCols.map(c => ({
        keyNum: c.keyNum,
        label: `${monthKeyToLabel(c.mk)} (${c.label})`
      }));
    }

    const mkCur = state.monthKey;

    const rowsTec = rgTec ? buildTableRows(rows, rgTec, byKeyNum, mkCur) : [];
    const rowsPro = rgPro ? buildTableRows(rows, rgPro, byKeyNum, mkCur) : [];

    html += `
      <div class="fech-card">
        <h3 class="fech-h2">${u}</h3>
        <div class="fech-muted">Fonte: ${tab} • Mês selecionado: ${monthKeyToLabel(mkCur)}</div>
      </div>
      ${rgTec ? renderOneBlock("Grau Técnico", rowsTec) : `<div class="fech-card">Não achei o bloco TÉCNICO na aba.</div>`}
      ${rgPro ? renderOneBlock("Profissionalizante", rowsPro) : `<div class="fech-card">Não achei o bloco PROFISSIONALIZANTE na aba.</div>`}
    `;
  }

  // Barra de filtros dentro da 4ª aba
  const monthSelect = `
    <div class="fech-topbar">
      <label><b>Unidade:</b></label>
      <select id="fech_unidade">
        <option value="TODAS"${state.unidade==="TODAS"?" selected":""}>Todas</option>
        ${Object.keys(TABS_UNIDADES).map(u => `<option value="${u}"${state.unidade===u?" selected":""}>${u}</option>`).join("")}
      </select>

      <label><b>Mês:</b></label>
      <select id="fech_mes">
        ${globalMonthOptions ? globalMonthOptions.map(o => {
          const selected = (globalMonthByKey && state.monthKey && monthToKeyNum(state.monthKey) === o.keyNum) ? " selected" : "";
          return `<option value="${o.keyNum}"${selected}>${o.label}</option>`;
        }).join("") : `<option>—</option>`}
      </select>

      <button id="fech_aplicar" style="padding:6px 10px;border-radius:10px;border:1px solid #cfd8cf;cursor:pointer">Aplicar</button>
      <button id="fech_mes_atual" style="padding:6px 10px;border-radius:10px;border:1px solid #cfd8cf;cursor:pointer">Mês mais recente</button>
    </div>
  `;

  container.innerHTML = monthSelect + html;

  // bind eventos
  const selU = $("fech_unidade");
  const selM = $("fech_mes");
  const btnA = $("fech_aplicar");
  const btnCur = $("fech_mes_atual");

  if (btnA) btnA.onclick = async () => {
    state.unidade = selU ? selU.value : state.unidade;
    if (selM) {
      const keyNum = Number(selM.value);
      const y = Math.floor(keyNum / 100);
      const m = keyNum % 100;
      state.monthKey = { y, m };
    }
    await renderFechamento();
  };

  if (btnCur) btnCur.onclick = async () => {
    state.unidade = selU ? selU.value : state.unidade;
    state.monthKey = null; // força recalcular “mais recente”
    await renderFechamento();
  };
}

// ================= START =================
window.addEventListener("load", async () => {
  await renderFechamento();
});
