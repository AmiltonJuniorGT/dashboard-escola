// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAJAZEIRAS: "Números Cajazeiras",
  CAMACARI: "Números Camaçari",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_UNIT_KEY = "CAMACARI";

// ================= MATCH (Coluna A) =================
function norm(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}
function labelHas(label, keys) {
  const L = norm(label);
  return keys.some(k => L.includes(norm(k)));
}

// ✅ Incluímos CADASTROS aqui
const KPI_MATCH = {
  ATIVOS_FINAL: ["ALUNOS ATIVOS T", "ALUNOS ATIVOS P", "ATIVOS", "ATIVOS FINAL"],
  MATRICULAS: ["MATRICULAS REALIZADAS"],
  CADASTROS: ["CADASTROS"],
  FATURAMENTO: ["FATURAMENTO T (R$)", "FATURAMENTO P (R$)", "FATURAMENTO"],
  CUSTO_OPERACIONAL: ["CUSTO OPERACIONAL R$"],
  LUCRO_OPERACIONAL: ["LUCRO OPERACIONAL R$"],
  RESULTADO_LIQUIDO_TOTAL: ["RESULTADO LIQUIDO TOTAL R$"],
  MARGEM: ["MARGEM"],
  INADIMPLENCIA: ["INADIMPLENCIA"],
  EVASAO: ["EVASAO REAL"],
  PARCELAS_RECEBIDAS: ["PARCELAS RECEBIDAS DO MES T", "PARCELAS RECEBIDAS DO MES P", "PARCELAS RECEBIDAS"],
};

// ================= STATE =================
let currentRows = [];
let colLabels = [];
let headerMonths = []; // [{idx, label, key}]
let selectedMonthIdx = null;
let selectedMonthLabel = "";

// ================= DOM =================
const $ = (id) => document.getElementById(id);

const elStatus = $("status");
const elUnit = $("unitSelect");
const elStart = $("dateStart");
const elEnd = $("dateEnd");
const elMode = $("tableMode");
const btnLoad = $("btnLoad");
const btnApply = $("btnApply");
const btnExpand = $("btnExpand");
const tblInfo = $("tblInfo");
const hintBox = $("hintBox");

const metaUnit = $("metaUnit");
const metaPeriod = $("metaPeriod");
const metaBaseMonth = $("metaBaseMonth");

const WRAP = $("tableWrap");
const TABLE = $("kpiTable");

// KPI targets
const KPIS = {
  faturamento: $("kpi_faturamento"),
  custo: $("kpi_custo"),
  resultado: $("kpi_resultado"),
  ativos: $("kpi_ativos"),
  matriculas: $("kpi_matriculas"),
  recebidas: $("kpi_recebidas"),
  inad: $("kpi_inad"),
  evasao: $("kpi_evasao"),
  margem: $("kpi_margem"),
  ativosFinal: $("kpi_ativos_final"),

  // ✅ novos
  cadastros: $("kpi_cadastros"),
  conversao: $("kpi_conversao"),

  faturamento_mom: $("kpi_faturamento_mom"),
  faturamento_yoy: $("kpi_faturamento_yoy"),
  custo_mom: $("kpi_custo_mom"),
  custo_yoy: $("kpi_custo_yoy"),
  resultado_mom: $("kpi_resultado_mom"),
  resultado_yoy: $("kpi_resultado_yoy"),
  ativos_mom: $("kpi_ativos_mom"),
  ativos_yoy: $("kpi_ativos_yoy"),
  matriculas_mom: $("kpi_matriculas_mom"),
  matriculas_yoy: $("kpi_matriculas_yoy"),
  recebidas_mom: $("kpi_recebidas_mom"),
  recebidas_yoy: $("kpi_recebidas_yoy"),
  inad_mom: $("kpi_inad_mom"),
  inad_yoy: $("kpi_inad_yoy"),
  evasao_mom: $("kpi_evasao_mom"),
  evasao_yoy: $("kpi_evasao_yoy"),
  margem_mom: $("kpi_margem_mom"),
  margem_yoy: $("kpi_margem_yoy"),

  ativos_final_mom: $("kpi_ativos_final_mom"),
  ativos_final_yoy: $("kpi_ativos_final_yoy"),
};

// ================= INIT =================
document.addEventListener("DOMContentLoaded", () => setupUI());

function setupUI() {
  const entries = Object.entries(TABS);
  elUnit.innerHTML = entries.map(([k,v]) =>
    `<option value="${k}">${v.replace("Números ","")}</option>`
  ).join("");
  elUnit.value = DEFAULT_UNIT_KEY;

  const now = new Date();
  elStart.value = toISODate(new Date(now.getFullYear(), 0, 1));
  elEnd.value = toISODate(new Date(now.getFullYear(), 11, 31));

  btnLoad.addEventListener("click", () => loadData());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "Carregando…";
  loadData();
}

function toISODate(d){
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}
function setStatus(msg){ elStatus.textContent = msg; }

// ================= GVIZ FETCH =================
function sheetUrl(tabName){
  const base = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?`;
  const tq = encodeURIComponent("select *");
  const sheet = encodeURIComponent(tabName);
  return `${base}tqx=out:json&headers=1&tq=${tq}&sheet=${sheet}`;
}
function parseGviz(text){
  const m = text.match(/google\.visualization\.Query\.setResponse\(([\s\S]*)\);?$/);
  if(!m) throw new Error("GViz não retornou setResponse().");
  return JSON.parse(m[1]);
}
async function loadData(){
  const unitKey = elUnit.value;
  const tabName = TABS[unitKey];
  setStatus(`Carregando: ${tabName}…`);

  try{
    const res = await fetch(sheetUrl(tabName), { cache:"no-store" });
    const txt = await res.text();
    const data = parseGviz(txt);

    if(data.status !== "ok"){
      throw new Error(data.errors?.[0]?.detailed_message || "GViz status != ok");
    }

    colLabels = (data.table.cols || []).map(c => (c?.label ?? "").trim());
    currentRows = gvizTableToRows(data.table);

    // fallback se labels vierem vazios: usa primeira linha como cabeçalho
    const labelsOk = colLabels.slice(1).some(x => x && x.length);
    if(!labelsOk && currentRows.length){
      colLabels = currentRows[0].map(v => String(v ?? "").trim());
      currentRows = currentRows.slice(1);
    }

    detectHeaderMonthsFromCols();
    setStatus(`Dados carregados. (${tabName})`);
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro ao carregar dados.");
    hintBox.textContent = "Erro: " + (err?.message || err);
  }
}
function gvizTableToRows(table){
  const cols = table.cols.length;
  const rows = table.rows.length;
  const out = [];
  for(let r=0;r<rows;r++){
    const row = [];
    for(let c=0;c<cols;c++){
      const cell = table.rows[r].c[c];
      row.push(cell ? (cell.v ?? cell.f ?? null) : null);
    }
    out.push(row);
  }
  return out;
}

// ================= MONTH DETECTION (COL LABELS) =================
const MONTH_MAP = {
  JAN: "01", JANEIRO:"01",
  FEV: "02", FEVEREIRO:"02",
  MAR: "03", MARCO:"03",
  ABR: "04", ABRIL:"04",
  MAI: "05", MAIO:"05",
  JUN: "06", JUNHO:"06",
  JUL: "07", JULHO:"07",
  AGO: "08", AGOSTO:"08",
  SET: "09", SETEMBRO:"09",
  OUT: "10", OUTUBRO:"10",
  NOV: "11", NOVEMBRO:"11",
  DEZ: "12", DEZEMBRO:"12",
};

function monthLabelToKeyFlexible(raw){
  const s0 = norm(raw);
  if(!s0) return null;
  const s = s0.replace(/[.\-]/g, "/").replace(/\s+/g, "/");

  const yMatch = s.match(/(\d{4}|\d{2})/);
  if(!yMatch) return null;
  let yyyy = yMatch[1];
  if(yyyy.length === 2) yyyy = "20" + yyyy;

  let mm = null;
  for(const k of Object.keys(MONTH_MAP)){
    if(s.includes(k)){ mm = MONTH_MAP[k]; break; }
  }
  if(!mm) return null;

  return `${yyyy}-${mm}`;
}

function detectHeaderMonthsFromCols(){
  headerMonths = [];
  for(let c=1;c<colLabels.length;c++){
    const label = colLabels[c];
    const key = monthLabelToKeyFlexible(label);
    if(key){
      headerMonths.push({ idx:c, label, key });
    }
  }

  if(!headerMonths.length){
    hintBox.textContent = "⚠️ Não detectei meses no cabeçalho. Verifique se os cabeçalhos são tipo 'jan/25', 'fev/25' etc.";
  } else {
    hintBox.textContent = `✅ Meses detectados: ${headerMonths.length} (ex: ${headerMonths[0].label} … ${headerMonths[headerMonths.length-1].label})`;
  }
}

function findMonthColByKey(key){
  const m = headerMonths.find(x => x.key === key);
  return m ? m.idx : null;
}

// ================= KPI LOOKUPS =================
function findRowByKpiKey(kpiKey){
  const keys = KPI_MATCH[kpiKey] || [];
  for(let r=0;r<currentRows.length;r++){
    const label = currentRows[r]?.[0];
    if(labelHas(label, keys)) return currentRows[r];
  }
  return null;
}

function getCellNumber(row, colIdx){
  if(!row || colIdx == null) return null;
  const v = row[colIdx];
  if (v == null || v === "") return null;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  const cleaned = s.replace(/\./g,"").replace(",",".").replace(/[^\d\.\-]/g,"");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function fmtByKey(kpiKey, n){
  if(n == null) return "—";

  const isPct =
    kpiKey.includes("MARGEM") ||
    kpiKey.includes("INADIMPLENCIA") ||
    kpiKey.includes("EVASAO");

  const isMoney =
    kpiKey.includes("FATURAMENTO") ||
    kpiKey.includes("CUSTO") ||
    kpiKey.includes("LUCRO") ||
    kpiKey.includes("RESULTADO");

  if(isPct){
    const val = (n <= 1.5) ? (n*100) : n;
    return `${val.toFixed(2).replace(".",",")}%`;
  }
  if(isMoney){
    return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
  }
  const isInt = Math.abs(n - Math.round(n)) < 1e-9;
  return isInt ? String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ".") : n.toLocaleString("pt-BR");
}

function fmtInt(n){
  if(n == null) return "—";
  const isInt = Math.abs(n - Math.round(n)) < 1e-9;
  const val = isInt ? Math.round(n) : n;
  return String(val).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function deltaClass(delta, preferLowerIsBetter=false){
  if(delta == null) return "";
  if(preferLowerIsBetter){
    if(delta < 0) return "good";
    if(delta > 0) return "bad";
    return "";
  }
  if(delta > 0) return "good";
  if(delta < 0) return "bad";
  return "";
}

function fmtDelta(base, compare, kpiKey){
  if(base == null || compare == null) return { text:"—", cls:"" };
  const delta = base - compare;
  const lowerIsBetter = (kpiKey === "INADIMPLENCIA" || kpiKey === "EVASAO");
  const cls = deltaClass(delta, lowerIsBetter);

  if (compare !== 0) {
    const pct = (delta / compare) * 100;
    const s = `${pct >= 0 ? "+" : ""}${pct.toFixed(2).replace(".",",")}%`;
    return { text:s, cls };
  }
  return { text: `${delta >= 0 ? "+" : ""}${delta.toFixed(2).replace(".",",")}`, cls };
}

// ================= APPLY / RENDER =================
function apply(){
  if(!currentRows.length){
    setStatus("Sem dados.");
    return;
  }

  const unitKey = elUnit.value;
  const tabName = TABS[unitKey];

  metaUnit.textContent = tabName.replace("Números ","");
  metaPeriod.textContent = `${brDate(elStart.value)} → ${brDate(elEnd.value)}`;

  const baseDate = new Date(elEnd.value + "T00:00:00");
  const prevDate = new Date(baseDate.getFullYear(), baseDate.getMonth()-1, 1);
  const yoyDate  = new Date(baseDate.getFullYear()-1, baseDate.getMonth(), 1);

  const baseKey = monthKey(baseDate);
  const prevKey = monthKey(prevDate);
  const yoyKey  = monthKey(yoyDate);

  selectedMonthIdx = findMonthColByKey(baseKey);
  if(selectedMonthIdx == null && headerMonths.length){
    selectedMonthIdx = headerMonths[headerMonths.length-1].idx;
  }

  const prevIdx = findMonthColByKey(prevKey);
  const yoyIdx  = findMonthColByKey(yoyKey);

  selectedMonthLabel = (selectedMonthIdx != null ? colLabels[selectedMonthIdx] : "") || "";
  metaBaseMonth.textContent = selectedMonthLabel ? String(selectedMonthLabel) : "—";

  // cards
  renderCard({ valueEl: KPIS.faturamento, momEl: KPIS.faturamento_mom, yoyEl: KPIS.faturamento_yoy, kpiKey: "FATURAMENTO", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.custo, momEl: KPIS.custo_mom, yoyEl: KPIS.custo_yoy, kpiKey: "CUSTO_OPERACIONAL", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });

  renderCard({
    valueEl: KPIS.resultado,
    momEl: KPIS.resultado_mom,
    yoyEl: KPIS.resultado_yoy,
    kpiKey: "LUCRO_OPERACIONAL",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx,
    fallbackKey: "RESULTADO_LIQUIDO_TOTAL"
  });

  renderCard({ valueEl: KPIS.ativos, momEl: KPIS.ativos_mom, yoyEl: KPIS.ativos_yoy, kpiKey: "ATIVOS_FINAL", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });

  // ✅ vendas + cadastros + conversão
  renderSalesBlock(selectedMonthIdx, prevIdx, yoyIdx);

  renderCard({ valueEl: KPIS.recebidas, momEl: KPIS.recebidas_mom, yoyEl: KPIS.recebidas_yoy, kpiKey: "PARCELAS_RECEBIDAS", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.inad, momEl: KPIS.inad_mom, yoyEl: KPIS.inad_yoy, kpiKey: "INADIMPLENCIA", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.evasao, momEl: KPIS.evasao_mom, yoyEl: KPIS.evasao_yoy, kpiKey: "EVASAO", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.margem, momEl: KPIS.margem_mom, yoyEl: KPIS.margem_yoy, kpiKey: "MARGEM", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });

  // Ativos final (mesma base)
  renderCard({ valueEl: KPIS.ativosFinal, momEl: KPIS.ativos_final_mom, yoyEl: KPIS.ativos_final_yoy, kpiKey: "ATIVOS_FINAL", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });

  renderTable();

  tblInfo.textContent = `Base: ${selectedMonthLabel || "—"} • Meses: ${headerMonths.length}`;
  setStatus(`Dados carregados. (${tabName})`);
}

function renderSalesBlock(baseIdx, prevIdx, yoyIdx){
  const rowMat = findRowByKpiKey("MATRICULAS");
  const rowCad = findRowByKpiKey("CADASTROS");

  const matriculas = getCellNumber(rowMat, baseIdx);
  const cadastros  = getCellNumber(rowCad, baseIdx);

  KPIS.matriculas.textContent = fmtInt(matriculas);
  KPIS.cadastros.textContent  = fmtInt(cadastros);

  // conversão = matrículas / cadastros
  let conv = null;
  if(matriculas != null && cadastros != null && cadastros !== 0){
    conv = (matriculas / cadastros) * 100;
  }
  KPIS.conversao.textContent = (conv == null) ? "—" : `${conv.toFixed(2).replace(".",",")}%`;

  // deltas das matrículas (como já era)
  const prev = getCellNumber(rowMat, prevIdx);
  const yoy  = getCellNumber(rowMat, yoyIdx);

  const d1 = fmtDelta(matriculas, prev, "MATRICULAS");
  KPIS.matriculas_mom.textContent = d1.text === "—" ? "vs mês anterior —" : `vs mês anterior ${d1.text}`;
  KPIS.matriculas_mom.className = `kpiDelta ${d1.cls}`;

  const d2 = fmtDelta(matriculas, yoy, "MATRICULAS");
  KPIS.matriculas_yoy.textContent = d2.text === "—" ? "vs ano anterior —" : `vs ano anterior ${d2.text}`;
  KPIS.matriculas_yoy.className = `kpiDelta ${d2.cls}`;
}

function renderCard({ valueEl, momEl, yoyEl, kpiKey, baseIdx, prevIdx, yoyIdx, fallbackKey=null }){
  let row = findRowByKpiKey(kpiKey);
  let usedKey = kpiKey;

  if(!row && fallbackKey){
    row = findRowByKpiKey(fallbackKey);
    usedKey = fallbackKey;
  }

  const base = getCellNumber(row, baseIdx);
  const prev = getCellNumber(row, prevIdx);
  const yoy = getCellNumber(row, yoyIdx);

  valueEl.textContent = fmtByKey(usedKey, base);

  const d1 = fmtDelta(base, prev, usedKey);
  momEl.textContent = d1.text === "—" ? "vs mês anterior —" : `vs mês anterior ${d1.text}`;
  momEl.className = `kpiDelta ${d1.cls}`;

  const d2 = fmtDelta(base, yoy, usedKey);
  yoyEl.textContent = d2.text === "—" ? "vs ano anterior —" : `vs ano anterior ${d2.text}`;
  yoyEl.className = `kpiDelta ${d2.cls}`;
}

// ================= TABLE =================
function renderTable(){
  TABLE.innerHTML = "";

  if(!headerMonths.length){
    TABLE.innerHTML = `<tr><td class="muted">Não identifiquei meses no cabeçalho.</td></tr>`;
    return;
  }

  const mode = elMode.value;
  const cols = buildVisibleColumns(mode);

  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  cols.forEach(c => {
    const th = document.createElement("th");
    th.textContent = c.label;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  TABLE.appendChild(thead);

  const tbody = document.createElement("tbody");

  for(let r=0;r<currentRows.length;r++){
    const row = currentRows[r];
    if(!row) continue;

    const label = String(row[0] ?? "").trim();
    if(!label) continue;

    const tr = document.createElement("tr");
    if(isGroupRow(label, row)) tr.classList.add("trGroup");

    cols.forEach((c, i) => {
      const td = document.createElement("td");
      const v = row[c.idx];
      td.textContent = (i===0) ? label : (v == null ? "" : String(v));
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }

  TABLE.appendChild(tbody);
}

function buildVisibleColumns(mode){
  const cols = [{ idx:0, label:"Indicador" }];

  if(mode === "full"){
    headerMonths.forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }

  const start = new Date(elStart.value + "T00:00:00");
  const end = new Date(elEnd.value + "T00:00:00");
  const aKey = monthKey(start);
  const bKey = monthKey(end);

  const aPos = headerMonths.findIndex(m => m.key === aKey);
  const bPos = headerMonths.findIndex(m => m.key === bKey);

  if(aPos === -1 || bPos === -1){
    headerMonths.slice(Math.max(0, headerMonths.length-8))
      .forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }

  const a = Math.min(aPos, bPos);
  const b = Math.max(aPos, bPos);
  for(let i=a;i<=b;i++){
    cols.push({ idx: headerMonths[i].idx, label: headerMonths[i].label });
  }
  return cols;
}

function isGroupRow(label, row){
  const L = norm(label);
  const looksUpper = L === label.toUpperCase();
  const hasNumbers = row.slice(1).some(v => v != null && String(v).match(/\d/));
  return (looksUpper && !hasNumbers) || L.includes("PLANILHA DE INDICADORES");
}

// ================= HELPERS =================
function brDate(iso){
  if(!iso) return "—";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function monthKey(date){
  const mm = String(date.getMonth()+1).padStart(2,"0");
  return `${date.getFullYear()}-${mm}`;
}
