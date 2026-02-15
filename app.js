// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
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

const KPI_MATCH = {
  ATIVOS_FINAL: ["ALUNOS ATIVOS T", "ALUNOS ATIVOS P", "ATIVOS"],
  VENDAS_MATRICULAS: ["MATRICULAS REALIZADAS"],
  FATURAMENTO: ["FATURAMENTO T (R$)", "FATURAMENTO P (R$)", "FATURAMENTO"],
  CUSTO_OPERACIONAL: ["CUSTO OPERACIONAL R$"],
  LUCRO_OPERACIONAL: ["LUCRO OPERACIONAL R$"],
  MARGEM: ["MARGEM"],
  INADIMPLENCIA: ["INADIMPLENCIA"],
  EVASAO: ["EVASAO REAL"],
  PARCELAS_RECEBIDAS: ["PARCELAS RECEBIDAS DO MES T", "PARCELAS RECEBIDAS DO MES P"],

  RESULTADO_LIQUIDO_TOTAL: ["RESULTADO LIQUIDO TOTAL R$"],
};

// ================= STATE =================
let currentRows = [];
let headerMonths = [];        // [{idx, label}]
let headerRowIndex = -1;
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

// ================= INIT (SEM Google Charts) =================
document.addEventListener("DOMContentLoaded", () => {
  setupUI();
});

function setupUI() {
  // dropdown SEMPRE aparece
  const entries = Object.entries(TABS);
  elUnit.innerHTML = entries.map(([k,v]) =>
    `<option value="${k}">${v.replace("Números ","")}</option>`
  ).join("");
  elUnit.value = DEFAULT_UNIT_KEY;

  // datas default
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  elStart.value = toISODate(start);
  elEnd.value = toISODate(end);

  btnLoad.addEventListener("click", () => loadData());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Selecione a Unidade e clique em Carregar.");
  hintBox.textContent = "Se o dropdown está preenchido, a UI está OK. Agora vamos buscar os dados da planilha.";
}

function toISODate(d){
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}

function setStatus(msg){ elStatus.textContent = msg; }

// ================= GVIZ FETCH (JSONP -> JSON) =================
function sheetUrl(tabName){
  const base = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?`;
  const tq = encodeURIComponent("select *");
  const sheet = encodeURIComponent(tabName);
  // tqx=out:json ajuda a manter formato mais previsível
  return `${base}tqx=out:json&tq=${tq}&sheet=${sheet}`;
}

// Parse GViz response: google.visualization.Query.setResponse({...})
function parseGviz(text){
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if(start === -1 || end === -1) throw new Error("Resposta GViz inválida (sem JSON).");
  const json = text.slice(start, end+1);
  return JSON.parse(json);
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

    currentRows = gvizTableToRows(data.table);
    detectHeaderMonths();

    setStatus(`Dados carregados. (${tabName})`);
    hintBox.textContent = "Dados carregados. Match inteligente ativo (Coluna A).";
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro ao carregar dados.");
    hintBox.textContent =
      "Falha ao ler a planilha via GViz. " +
      "Verifique se a planilha está 'Qualquer pessoa com o link' e se a aba existe exatamente com esse nome. " +
      "Detalhe: " + (err?.message || err);
  }
}

function gvizTableToRows(table){
  const cols = table.cols.length;
  const rows = table.rows.length;
  const out = [];

  for(let r=0;r<rows;r++){
    const row = [];
    for(let c=0;c<cols;c++){
      // v (raw) ou f (formatted)
      const cell = table.rows[r].c[c];
      row.push(cell ? (cell.v ?? cell.f ?? null) : null);
    }
    out.push(row);
  }
  return out;
}

// ================= HEADER/MONTH DETECTION =================
const MONTH_PATTERN = /(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z]*\/\d{2,4}/;

function detectHeaderMonths(){
  headerMonths = [];
  headerRowIndex = -1;

  for(let r=0;r<currentRows.length;r++){
    const row = currentRows[r];
    let hits = 0;
    for(let c=1;c<row.length;c++){
      const v = row[c];
      if (v && MONTH_PATTERN.test(norm(v))) hits++;
    }
    if(hits >= 3){
      headerRowIndex = r;
      for(let c=1;c<row.length;c++){
        const v = row[c];
        if (v && MONTH_PATTERN.test(norm(v))) {
          headerMonths.push({ idx:c, label:String(v).trim() });
        }
      }
      break;
    }
  }

  if(headerRowIndex === -1 || headerMonths.length === 0){
    hintBox.textContent =
      "Atenção: não identifiquei a linha de meses automaticamente. " +
      "Verifique se existe uma linha com 'janeiro/23, fevereiro/25…' nas colunas.";
  }
}

// ================= KPI LOOKUPS =================
function findRowByKpiKey(kpiKey){
  const keys = KPI_MATCH[kpiKey] || [];
  if(!keys.length) return null;

  for(let r=0;r<currentRows.length;r++){
    if (r === headerRowIndex) continue;
    const label = currentRows[r]?.[0];
    if (labelHas(label, keys)) return currentRows[r];
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
  const key = kpiKey;

  const isPct =
    key.includes("MARGEM") ||
    key.includes("INADIMPLENCIA") ||
    key.includes("EVASAO");

  const isMoney =
    key.includes("FATURAMENTO") ||
    key.includes("CUSTO") ||
    key.includes("LUCRO") ||
    key.includes("RESULTADO");

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
    setStatus("Sem dados. Clique em Carregar.");
    return;
  }

  const unitKey = elUnit.value;
  const tabName = TABS[unitKey];

  metaUnit.textContent = tabName.replace("Números ","");
  metaPeriod.textContent = `${brDate(elStart.value)} → ${brDate(elEnd.value)}`;

  const base = new Date(elEnd.value + "T00:00:00");
  const prev = new Date(base.getFullYear(), base.getMonth()-1, 1);
  const yoy = new Date(base.getFullYear()-1, base.getMonth(), 1);

  const baseKey = monthKey(base);
  const prevKey = monthKey(prev);
  const yoyKey = monthKey(yoy);

  selectedMonthIdx = findMonthColByKey(baseKey);
  const prevIdx = findMonthColByKey(prevKey);
  const yoyIdx = findMonthColByKey(yoyKey);

  selectedMonthLabel = selectedMonthIdx != null
    ? currentRows[headerRowIndex]?.[selectedMonthIdx] ?? ""
    : "";

  metaBaseMonth.textContent = selectedMonthLabel ? String(selectedMonthLabel) : "—";

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
  renderCard({ valueEl: KPIS.matriculas, momEl: KPIS.matriculas_mom, yoyEl: KPIS.matriculas_yoy, kpiKey: "VENDAS_MATRICULAS", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.recebidas, momEl: KPIS.recebidas_mom, yoyEl: KPIS.recebidas_yoy, kpiKey: "PARCELAS_RECEBIDAS", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.inad, momEl: KPIS.inad_mom, yoyEl: KPIS.inad_yoy, kpiKey: "INADIMPLENCIA", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.evasao, momEl: KPIS.evasao_mom, yoyEl: KPIS.evasao_yoy, kpiKey: "EVASAO", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.margem, momEl: KPIS.margem_mom, yoyEl: KPIS.margem_yoy, kpiKey: "MARGEM", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });

  renderCard({ valueEl: KPIS.ativosFinal, momEl: KPIS.ativos_final_mom, yoyEl: KPIS.ativos_final_yoy, kpiKey: "ATIVOS_FINAL", baseIdx: selectedMonthIdx, prevIdx, yoyIdx });

  renderTable();

  const baseTxt = selectedMonthLabel ? `Base: ${selectedMonthLabel}` : "Base: —";
  const headerTxt = headerMonths.length ? `Meses detectados: ${headerMonths.length}` : "Meses detectados: 0";
  tblInfo.textContent = `${baseTxt} • ${headerTxt}`;
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

function renderTable(){
  TABLE.innerHTML = "";

  if(headerRowIndex === -1){
    TABLE.innerHTML = `<tr><td class="muted">Não consegui identificar a linha de meses.</td></tr>`;
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
  for(let r=headerRowIndex+1;r<currentRows.length;r++){
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
  const startKey = monthKey(start);
  const endKey = monthKey(end);

  const startIdx = findMonthIndexByKey(startKey);
  const endIdx = findMonthIndexByKey(endKey);

  if(startIdx === -1 || endIdx === -1){
    const slice = headerMonths.slice(Math.max(0, headerMonths.length-8));
    slice.forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }

  const a = Math.min(startIdx, endIdx);
  const b = Math.max(startIdx, endIdx);
  for(let i=a;i<=b;i++){
    const m = headerMonths[i];
    cols.push({ idx:m.idx, label:m.label });
  }
  return cols;
}

function isGroupRow(label, row){
  const L = norm(label);
  const looksUpper = L === label.toUpperCase();
  const hasNumbers = row.slice(1).some(v => v != null && String(v).match(/\d/));
  return (looksUpper && !hasNumbers) || L.includes("PLANILHA DE INDICADORES");
}

// ================= MONTH HELPERS =================
function brDate(iso){
  if(!iso) return "—";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function monthKey(date){
  const mm = String(date.getMonth()+1).padStart(2,"0");
  return `${date.getFullYear()}-${mm}`;
}
function findMonthIndexByKey(key){
  for(let i=0;i<headerMonths.length;i++){
    const k = monthLabelToKey(norm(headerMonths[i].label));
    if(k === key) return i;
  }
  return -1;
}
function findMonthColByKey(key){
  const i = findMonthIndexByKey(key);
  if(i === -1) return null;
  return headerMonths[i].idx;
}
function monthLabelToKey(labelNorm){
  const parts = labelNorm.split("/");
  if(parts.length < 2) return null;

  const mPart = parts[0].trim();
  let yPart = parts[1].trim();

  const map = {
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

  const mm = map[mPart] || map[mPart.slice(0,3)] || null;
  if(!mm) return null;

  if(yPart.length === 2) yPart = "20" + yPart;
  if(yPart.length !== 4) return null;

  return `${yPart}-${mm}`;
}
