// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_UNIT_KEY = "CAMACARI"; // ✅ Camaçari como base

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
let headerMonths = [];        // [{idx, label, key}]
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

// ================= INIT =================
document.addEventListener("DOMContentLoaded", () => setupUI());

function setupUI() {
  const entries = Object.entries(TABS);
  elUnit.innerHTML = entries.map(([k,v]) =>
    `<option value="${k}">${v.replace("Números ","")}</option>`
  ).join("");
  elUnit.value = DEFAULT_UNIT_KEY;

  // padrão: ano atual inteiro
  const now = new Date();
  const start = new Date(now.getFullYear(), 0, 1);
  const end = new Date(now.getFullYear(), 11, 31);
  elStart.value = toISODate(start);
  elEnd.value = toISODate(end);

  btnLoad.addEventListener("click", () => loadData());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Carregando dados…");
  hintBox.textContent = "Usando Camaçari como base. Detectando meses no cabeçalho…";
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
  return `${base}tqx=out:json&tq=${tq}&sheet=${sheet}`;
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
    const res = await fetch(sheetUrl(tabName), { cache:"no-store" });
    const txt = await res.text();
    const data = parseGviz(txt);

    if(data.status !== "ok"){
      throw new Error(data.errors?.[0]?.detailed_message || "GViz status != ok");
    }

    currentRows = gvizTableToRows(data.table);
    detectHeaderMonths(); // ✅ agora assume cabeçalho quando fizer sentido

    setStatus(`Dados carregados. (${tabName})`);
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro ao carregar dados.");
    hintBox.textContent =
      "Falha ao ler a planilha via GViz. Verifique compartilhamento e nomes das abas. " +
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
      const cell = table.rows[r].c[c];
      row.push(cell ? (cell.v ?? cell.f ?? null) : null);
    }
    out.push(row);
  }
  return out;
}

// ================= MONTH DETECTION (cabeçalho) =================
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

  // separadores: "/", ".", "-", espaço
  const s = s0.replace(/[.\-]/g, "/").replace(/\s+/g, "/");

  // ano (2 ou 4)
  const yMatch = s.match(/(\d{4}|\d{2})/);
  if(!yMatch) return null;
  let yyyy = yMatch[1];
  if(yyyy.length === 2) yyyy = "20" + yyyy;

  // mês
  let mm = null;
  for(const k of Object.keys(MONTH_MAP)){
    if(s.includes(k)){
      mm = MONTH_MAP[k];
      break;
    }
  }
  if(!mm) return null;

  return `${yyyy}-${mm}`;
}

function isMonthCell(v){
  return monthLabelToKeyFlexible(v) != null;
}

function detectHeaderMonths(){
  headerMonths = [];
  headerRowIndex = -1;

  if(!currentRows.length) return;

  // 1) tenta assumir que o cabeçalho está na linha 1 (row 0)
  const row0 = currentRows[0];
  let hits0 = 0;
  for(let c=1;c<row0.length;c++){
    if(isMonthCell(row0[c])) hits0++;
  }
  if(hits0 >= 3){
    headerRowIndex = 0;
  } else {
    // 2) fallback: procura em outras linhas (caso tenha título acima)
    for(let r=1;r<Math.min(15, currentRows.length);r++){
      const row = currentRows[r];
      let hits = 0;
      for(let c=1;c<row.length;c++){
        if(isMonthCell(row[c])) hits++;
      }
      if(hits >= 3){
        headerRowIndex = r;
        break;
      }
    }
  }

  if(headerRowIndex >= 0){
    const hr = currentRows[headerRowIndex];
    for(let c=1;c<hr.length;c++){
      if(isMonthCell(hr[c])){
        const label = String(hr[c]).trim();
        const key = monthLabelToKeyFlexible(label);
        headerMonths.push({ idx:c, label, key });
      }
    }
  }

  if(headerRowIndex === -1 || headerMonths.length === 0){
    hintBox.textContent =
      "⚠️ Não identifiquei os meses no cabeçalho. " +
      "Me mande um print só da linha 1 (cabeçalho) da Camaçari se ainda falhar.";
  } else {
    const first = headerMonths[0]?.label || "—";
    const last = headerMonths[headerMonths.length-1]?.label || "—";
    hintBox.textContent =
      `✅ Meses detectados: ${headerMonths.length} (linha ${headerRowIndex+1}). ` +
      `Primeiro: ${first} • Último: ${last}`;
  }
}

function findMonthColByKey(key){
  if(!key) return null;
  const m = headerMonths.find(x => x.key === key);
  return m ? m.idx : null;
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
    return {
