const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_UNIT_KEY = "CAMACARI";

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

let currentRows = [];
let colLabels = [];
let colIds = [];
let headerMonths = []; // [{idx,label,key}]
let selectedMonthIdx = null;
let selectedMonthLabel = "";

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

document.addEventListener("DOMContentLoaded", () => setupUI());

function setupUI() {
  const entries = Object.entries(TABS);
  elUnit.innerHTML = entries.map(([k,v]) =>
    `<option value="${k}">${v.replace("Números ","")}</option>`
  ).join("");
  elUnit.value = DEFAULT_UNIT_KEY;

  const now = new Date();
  const start = new Date(now.getFullYear(), 0, 1);
  const end = new Date(now.getFullYear(), 11, 31);
  elStart.value = toISODate(start);
  elEnd.value = toISODate(end);

  btnLoad.addEventListener("click", () => loadData());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "DEBUG: aguardando carregamento…";
  loadData();
}

function toISODate(d){
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}
function setStatus(msg){ elStatus.textContent = msg; }

// ============== GVIZ URL (auto headers) ==============
function sheetUrl(tabName, headersVal){
  const base = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?`;
  const tq = encodeURIComponent("select *");
  const sheet = encodeURIComponent(tabName);
  // ✅ força headers, mas vamos testar 1 e 0 automaticamente
  return `${base}tqx=out:json&headers=${headersVal}&tq=${tq}&sheet=${sheet}`;
}
function parseGviz(text){
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if(start === -1 || end === -1) throw new Error("Resposta GViz inválida (sem JSON).");
  return JSON.parse(text.slice(start, end+1));
}

async function loadData(){
  const unitKey = elUnit.value;
  const tabName = TABS[unitKey];
  setStatus(`Carregando: ${tabName}…`);

  try{
    // tenta headers=1 primeiro; se vier “morto”, tenta 0
    const d1 = await fetchAndParse(tabName, 1);
    const ok1 = extractTable(d1);

    if(!ok1.monthsFound){
      const d0 = await fetchAndParse(tabName, 0);
      const ok0 = extractTable(d0);
      if(ok0.monthsFound){
        hintBox.textContent = ok0.debug + " | (usando headers=0)";
      } else {
        hintBox.textContent = ok0.debug + " | (falhou headers=1 e headers=0)";
      }
    } else {
      hintBox.textContent = ok1.debug + " | (usando headers=1)";
    }

    setStatus(`Dados carregados. (${tabName})`);
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro ao carregar dados.");
    hintBox.textContent = "Erro GViz: " + (err?.message || err);
  }
}

async function fetchAndParse(tabName, headersVal){
  const res = await fetch(sheetUrl(tabName, headersVal), { cache:"no-store" });
  const txt = await res.text();
  const data = parseGviz(txt);
  if(data.status !== "ok"){
    throw new Error(data.errors?.[0]?.detailed_message || "GViz status != ok");
  }
  return data;
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

// ======== MONTH PARSER ========
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
function isMonthLabel(x){ return monthLabelToKeyFlexible(x) != null; }

// ======== extract: tenta cols.label, cols.id, rows[0] ========
function extractTable(data){
  const table = data.table;
  colLabels = (table.cols || []).map(c => (c?.label ?? "").trim());
  colIds    = (table.cols || []).map(c => (c?.id ?? "").trim());
  currentRows = gvizTableToRows(table);

  // DEBUG: amostras
  const sampleLabels = colLabels.slice(0, 8).map(x => x || "∅").join(" | ");
  const sampleIds = colIds.slice(0, 8).map(x => x || "∅").join(" | ");
  const row0 = currentRows[0] ? currentRows[0].slice(0, 8).map(v => String(v ?? "∅")).join(" | ") : "∅";
  const debug =
    `DEBUG cols=${colLabels.length} rows=${currentRows.length} | ` +
    `labels[0..]=${sampleLabels} | ids[0..]=${sampleIds} | row0[0..]=${row0}`;

  // 1) meses em labels
  let months = buildMonthsFromArray(colLabels);
  // 2) se não, meses em ids
  if(months.length < 3) months = buildMonthsFromArray(colIds);

  // 3) se ainda não, tenta primeira linha dos dados como cabeçalho
  if(months.length < 3 && currentRows.length){
    const head = currentRows[0].map(v => String(v ?? "").trim());
    const monthsFromRow = buildMonthsFromArray(head);
    if(monthsFromRow.length >= 3){
      colLabels = head;
      months = monthsFromRow;
      currentRows = currentRows.slice(1); // remove header fake
    }
  }

  headerMonths = months;
  return { monthsFound: headerMonths.length >= 3, debug };
}

function buildMonthsFromArray(arr){
  const out = [];
  for(let i=1;i<arr.length;i++){
    const label = arr[i];
    const key = monthLabelToKeyFlexible(label);
    if(key) out.push({ idx:i, label, key });
  }
  return out;
}

function findMonthColByKey(key){
  const m = headerMonths.find(x => x.key === key);
  return m ? m.idx : null;
}

// ======== KPIs ========
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
  if(v == null || v === "") return null;
  if(typeof v === "number") return v;
  const s = String(v).trim();
  const cleaned = s.replace(/\./g,"").replace(",",".").replace(/[^\d\.\-]/g,"");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function fmtByKey(kpiKey, n){
  if(n == null) return "—";
  const isPct = kpiKey.includes("MARGEM") || kpiKey.includes("INADIMPLENCIA") || kpiKey.includes("EVASAO");
  const isMoney = kpiKey.includes("FATURAMENTO") || kpiKey.includes("CUSTO") || kpiKey.includes("LUCRO") || kpiKey.includes("RESULTADO");
  if(isPct){
    const val = (n <= 1.5) ? (n*100) : n;
    return `${val.toFixed(2).replace(".",",")}%`;
  }
  if(isMoney) return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
  const isInt = Math.abs(n - Math.round(n)) < 1e-9;
  return isInt ? String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ".") : n.toLocaleString("pt-BR");
}

function deltaClass(delta, lowerIsBetter=false){
  if(delta == null) return "";
  if(lowerIsBetter){
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
  if(compare !== 0){
    const pct = (delta/compare)*100;
    return { text: `${pct>=0?"+":""}${pct.toFixed(2).replace(".",",")}%`, cls };
  }
  return { text: `${delta>=0?"+":""}${delta.toFixed(2).replace(".",",")}`, cls };
}

function apply(){
  const unitKey = elUnit.value;
  const tabName = TABS[unitKey];

  metaUnit.textContent = tabName.replace("Números ","");
  metaPeriod.textContent = `${brDate(elStart.value)} → ${brDate(elEnd.value)}`;

  if(headerMonths.length < 3){
    metaBaseMonth.textContent = "—";
    tblInfo.textContent = "Sem meses detectados";
    // zera cards
    Object.values(KPIS).forEach(el => { if(el && el.tagName) el.textContent = "—"; });
    TABLE.innerHTML = `<tr><td class="muted">Sem meses detectados. Veja o DEBUG acima.</td></tr>`;
    return;
  }

  const baseDate = new Date(elEnd.value + "T00:00:00");
  const prevDate = new Date(baseDate.getFullYear(), baseDate.getMonth()-1, 1);
  const yoyDate  = new Date(baseDate.getFullYear()-1, baseDate.getMonth(), 1);

  const baseKey = monthKey(baseDate);
  const prevKey = monthKey(prevDate);
  const yoyKey  = monthKey(yoyDate);

  selectedMonthIdx = findMonthColByKey(baseKey);
  if(selectedMonthIdx == null) selectedMonthIdx = headerMonths[headerMonths.length-1].idx;

  const prevIdx = findMonthColByKey(prevKey);
  const yoyIdx  = findMonthColByKey(yoyKey);

  selectedMonthLabel = (colLabels[selectedMonthIdx] || headerMonths.find(m=>m.idx===selectedMonthIdx)?.label || "");
  metaBaseMonth.textContent = selectedMonthLabel || "—";

  renderCard({ valueEl: KPIS.faturamento, momEl: KPIS.faturamento_mom, yoyEl: KPIS.faturamento_yoy, kpiKey:"FATURAMENTO", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.custo, momEl: KPIS.custo_mom, yoyEl: KPIS.custo_yoy, kpiKey:"CUSTO_OPERACIONAL", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.resultado, momEl: KPIS.resultado_mom, yoyEl: KPIS.resultado_yoy, kpiKey:"LUCRO_OPERACIONAL", baseIdx:selectedMonthIdx, prevIdx, yoyIdx, fallbackKey:"RESULTADO_LIQUIDO_TOTAL" });

  renderCard({ valueEl: KPIS.ativos, momEl: KPIS.ativos_mom, yoyEl: KPIS.ativos_yoy, kpiKey:"ATIVOS_FINAL", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.matriculas, momEl: KPIS.matriculas_mom, yoyEl: KPIS.matriculas_yoy, kpiKey:"VENDAS_MATRICULAS", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.recebidas, momEl: KPIS.recebidas_mom, yoyEl: KPIS.recebidas_yoy, kpiKey:"PARCELAS_RECEBIDAS", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });

  renderCard({ valueEl: KPIS.inad, momEl: KPIS.inad_mom, yoyEl: KPIS.inad_yoy, kpiKey:"INADIMPLENCIA", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.evasao, momEl: KPIS.evasao_mom, yoyEl: KPIS.evasao_yoy, kpiKey:"EVASAO", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });
  renderCard({ valueEl: KPIS.margem, momEl: KPIS.margem_mom, yoyEl: KPIS.margem_yoy, kpiKey:"MARGEM", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });

  renderCard({ valueEl: KPIS.ativosFinal, momEl: KPIS.ativos_final_mom, yoyEl: KPIS.ativos_final_yoy, kpiKey:"ATIVOS_FINAL", baseIdx:selectedMonthIdx, prevIdx, yoyIdx });

  renderTable();

  tblInfo.textContent = `Base: ${selectedMonthLabel} • Meses: ${headerMonths.length}`;
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
  momEl.textContent = d1.text==="—" ? "vs mês anterior —" : `vs mês anterior ${d1.text}`;
  momEl.className = `kpiDelta ${d1.cls}`;

  const d2 = fmtDelta(base, yoy, usedKey);
  yoyEl.textContent = d2.text==="—" ? "vs ano anterior —" : `vs ano anterior ${d2.text}`;
  yoyEl.className = `kpiDelta ${d2.cls}`;
}

function renderTable(){
  TABLE.innerHTML = "";
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
    const label = String(row?.[0] ?? "").trim();
    if(!label) continue;

    const tr = document.createElement("tr");
    if(isGroupRow(label, row)) tr.classList.add("trGroup");

    cols.forEach((c, i) => {
      const td = document.createElement("td");
      const v = row[c.idx];
      td.textContent = i===0 ? label : (v==null ? "" : String(v));
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

function brDate(iso){
  if(!iso) return "—";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function monthKey(date){
  const mm = String(date.getMonth()+1).padStart(2,"0");
  return `${date.getFullYear()}-${mm}`;
}
