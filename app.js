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
  CADASTROS: ["CADASTROS"],
  CONVERSAO: ["CONVERSAO (%)"],
  FATURAMENTO: ["FATURAMENTO T (R$)", "FATURAMENTO P (R$)", "FATURAMENTO"],
  CUSTO_OPERACIONAL: ["CUSTO OPERACIONAL R$"],
  LUCRO_OPERACIONAL: ["LUCRO OPERACIONAL R$"],
  MARGEM: ["MARGEM"],
  INADIMPLENCIA: ["INADIMPLENCIA"],
  EVASAO: ["EVASAO REAL"],
  PARCELAS_RECEBIDAS: ["PARCELAS RECEBIDAS DO MES T", "PARCELAS RECEBIDAS DO MES P"],
  PARCELAS_A_RECEBER: ["PARCELAS A RECEBER DO MES T", "PARCELAS A RECEBER DO MES P"],
  PERC_RECEBIDAS_ATIVOS: ["PARCELAS RECEBIDAS / ATIVOS (%)"],
  TICKET_MEDIO_PARCELAS: ["TICKET MEDIO DAS PARCELAS (R$)"],

  RESULTADO_LIQUIDO_TOTAL: ["RESULTADO LIQUIDO TOTAL R$"],
  RESULTADO_PERCENTUAL: ["RESULTADO %"],
  SALDO_FINAL: ["SALDO FINAL"],

  GERACAO_CAIXA: ["GERACAO DE CAIXA"],
  SALDO_ITAU: ["SALDO DO BANCO ITAU"],
  SALDO_CAIXA: ["SALDO BANCO - CAIXA"],
  SALDO_SICOOB: ["SALDO BANCO - SICOOB"],
  DINHEIRO: ["DINHEIRO"],
  SALDO_TOTAL: ["SALDO TOTAL", " SALDO TOTAL"],

  CUSTO_TOTAL: ["CUSTO TOTAL", "CUSTO TOTAL R$"],
  INVESTIMENTO: ["INVESTIMENTO", "INVESTIMENTOS ATACAMA (R$)"],
  EMPRESTIMO: ["EMPRESTIMO"],
  DISTRIBUICAO_SOCIOS: ["DISTRIBUICAO SOCIOS"],
  CINCO_PORCENTO_GESTOR: ["5% DO GESTOR"],

  PESQ_INSTITUCIONAL: ["PESQUISA INSTITUCIONAL"],
  PESQ_INSTRUTOR: ["PESQUISA DO INSTRUTOR"],
};

// ================= STATE =================
let gvizReady = false;
let currentRows = [];         // array of arrays (full sheet)
let headerMonths = [];        // [{idx, label}]
let headerRowIndex = -1;      // index of header row
let selectedMonthIdx = null;  // column idx in sheet
let selectedMonthLabel = "";  // label like "fevereiro/25"

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
google.charts.load("current", { packages: ["corechart", "table"] });
google.charts.setOnLoadCallback(() => {
  gvizReady = true;
  elStatus.textContent = "Pronto.";
  setupUI();
});

function setupUI() {
  // populate unit select
  const entries = Object.entries(TABS);
  elUnit.innerHTML = entries.map(([k,v]) => `<option value="${k}">${v.replace("Números ","")}</option>`).join("");
  elUnit.value = DEFAULT_UNIT_KEY;

  // default dates: current month range
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  elStart.value = toISODate(start);
  elEnd.value = toISODate(end);

  btnLoad.addEventListener("click", () => loadData());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  loadData();
}

function toISODate(d){
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}

// ================= GVIZ FETCH =================
function sheetUrl(tabName){
  const base = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?`;
  const tq = encodeURIComponent("select *");
  const sheet = encodeURIComponent(tabName);
  return `${base}tq=${tq}&sheet=${sheet}`;
}

function loadData(){
  if(!gvizReady) return;
  const unitKey = elUnit.value;
  const tabName = TABS[unitKey];

  setStatus(`Carregando: ${tabName}…`);

  const query = new google.visualization.Query(sheetUrl(tabName));
  query.send((res) => {
    if (res.isError()) {
      setStatus(`Erro ao carregar: ${res.getMessage()}`);
      hintBox.textContent = `Erro: ${res.getDetailedMessage?.() || res.getMessage()}`;
      return;
    }
    const dt = res.getDataTable();
    currentRows = dataTableToRows(dt);
    detectHeaderMonths();
    setStatus(`Dados carregados. (${tabName})`);
    apply();
  });
}

function setStatus(msg){ elStatus.textContent = msg; }

function dataTableToRows(dt){
  const rows = [];
  const cols = dt.getNumberOfColumns();
  const rCount = dt.getNumberOfRows();
  for(let r=0;r<rCount;r++){
    const row = [];
    for(let c=0;c<cols;c++){
      row.push(dt.getValue(r,c));
    }
    rows.push(row);
  }
  return rows;
}

// ================= HEADER/MONTH DETECTION =================
const MONTH_PATTERN = /(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z]*\/\d{2,4}/;

function detectHeaderMonths(){
  headerMonths = [];
  headerRowIndex = -1;

  for(let r=0;r<currentRows.length;r++){
    const row = currentRows[r];
    // try find month pattern on any cell except first
    let hits = 0;
    for(let c=1;c<row.length;c++){
      const v = row[c];
      if (v && MONTH_PATTERN.test(norm(v))) hits++;
    }
    if(hits >= 3){ // heuristic: header row has several months
      headerRowIndex = r;
      // build headerMonths from that row
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
  } else {
    hintBox.textContent = "Dados carregados. Match inteligente ativo (Coluna A).";
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
  // try parse BR formats
  const s = String(v).trim();
  const cleaned = s
    .replace(/\./g,"")      // thousands dot
    .replace(",",".")       // decimal comma
    .replace(/[^\d\.\-]/g,""); // remove R$, %, spaces
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function fmtByKey(kpiKey, n){
  if(n == null) return "—";
  const key = kpiKey;

  const isPct =
    key.includes("MARGEM") ||
    key.includes("INADIMPLENCIA") ||
    key.includes("EVASAO") ||
    key.includes("PERC_") ||
    /%/.test(key);

  const isMoney =
    key.includes("FATURAMENTO") ||
    key.includes("CUSTO") ||
    key.includes("LUCRO") ||
    key.includes("RESULTADO") ||
    key.includes("SALDO") ||
    key.includes("INVEST") ||
    key.includes("GERACAO") ||
    key.includes("DISTRIBUICAO") ||
    key.includes("EMPRESTIMO");

  if(isPct){
    // if value is like 0.0916 -> show 9,16% OR already 9,16 -> show 9,16%
    const val = (n <= 1.5) ? (n*100) : n;
    return `${val.toFixed(2).replace(".",",")}%`;
  }

  if(isMoney){
    return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
  }

  // default integer-ish
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

  // percentage change (when compare != 0)
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

  // Base month = month of end date
  const base = new Date(elEnd.value + "T00:00:00");
  const prev = new Date(base.getFullYear(), base.getMonth()-1, 1);
  const yoy = new Date(base.getFullYear()-1, base.getMonth(), 1);

  // find best month column for base/prev/yoy using headerMonths labels
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

  // Cards
  renderCard({
    valueEl: KPIS.faturamento,
    momEl: KPIS.faturamento_mom,
    yoyEl: KPIS.faturamento_yoy,
    kpiKey: "FATURAMENTO",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  renderCard({
    valueEl: KPIS.custo,
    momEl: KPIS.custo_mom,
    yoyEl: KPIS.custo_yoy,
    kpiKey: "CUSTO_OPERACIONAL",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  // Resultado: por padrão usamos LUCRO_OPERACIONAL (técnico/prof) e fallback para RESULTADO_LIQUIDO_TOTAL (grau educacional)
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

  renderCard({
    valueEl: KPIS.ativos,
    momEl: KPIS.ativos_mom,
    yoyEl: KPIS.ativos_yoy,
    kpiKey: "ATIVOS_FINAL",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  renderCard({
    valueEl: KPIS.matriculas,
    momEl: KPIS.matriculas_mom,
    yoyEl: KPIS.matriculas_yoy,
    kpiKey: "VENDAS_MATRICULAS",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  renderCard({
    valueEl: KPIS.recebidas,
    momEl: KPIS.recebidas_mom,
    yoyEl: KPIS.recebidas_yoy,
    kpiKey: "PARCELAS_RECEBIDAS",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  renderCard({
    valueEl: KPIS.inad,
    momEl: KPIS.inad_mom,
    yoyEl: KPIS.inad_yoy,
    kpiKey: "INADIMPLENCIA",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  renderCard({
    valueEl: KPIS.evasao,
    momEl: KPIS.evasao_mom,
    yoyEl: KPIS.evasao_yoy,
    kpiKey: "EVASAO",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  renderCard({
    valueEl: KPIS.margem,
    momEl: KPIS.margem_mom,
    yoyEl: KPIS.margem_yoy,
    kpiKey: "MARGEM",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  // Wide: ativos final (mesmo KPI ATIVOS_FINAL)
  renderCard({
    valueEl: KPIS.ativosFinal,
    momEl: KPIS.ativos_final_mom,
    yoyEl: KPIS.ativos_final_yoy,
    kpiKey: "ATIVOS_FINAL",
    baseIdx: selectedMonthIdx,
    prevIdx,
    yoyIdx
  });

  // Table
  renderTable();

  // info
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

  // THEAD
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  cols.forEach((c, i) => {
    const th = document.createElement("th");
    th.textContent = c.label;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  TABLE.appendChild(thead);

  // TBODY
  const tbody = document.createElement("tbody");

  // We will iterate rows after headerRowIndex
  for(let r=headerRowIndex+1;r<currentRows.length;r++){
    const row = currentRows[r];
    if(!row) continue;

    const label = String(row[0] ?? "").trim();
    if(!label) continue;

    // group row heuristic: label all caps and no numeric in other cols
    const isGroup = isGroupRow(label, row);

    const tr = document.createElement("tr");
    if(isGroup) tr.classList.add("trGroup");

    cols.forEach((c, i) => {
      const td = document.createElement("td");
      const v = row[c.idx];

      if(i === 0){
        td.textContent = label;
      } else {
        // keep original for table, but align numbers
        td.textContent = v == null ? "" : String(v);
      }
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }

  TABLE.appendChild(tbody);
}

function buildVisibleColumns(mode){
  // always show label col 0
  const cols = [{ idx:0, label:"Indicador" }];

  if(mode === "full"){
    // show all detected month columns
    headerMonths.forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }

  // interval: filter based on dateStart/dateEnd keys
  const start = new Date(elStart.value + "T00:00:00");
  const end = new Date(elEnd.value + "T00:00:00");
  const startKey = monthKey(start);
  const endKey = monthKey(end);

  const startIdx = findMonthIndexByKey(startKey);
  const endIdx = findMonthIndexByKey(endKey);

  // fallback: show last 8 months if range not found
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
  // group rows usually have no numbers
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
  // key is YYYY-MM. Compare against header label like "fevereiro/25" or "jan/2023"
  for(let i=0;i<headerMonths.length;i++){
    const lab = norm(headerMonths[i].label);
    const k = monthLabelToKey(lab);
    if(k === key) return i;
  }
  return -1;
}
function findMonthColByKey(key){
  const i = findMonthIndexByKey(key);
  if(i === -1) return null;
  return headerMonths[i].idx;
}

// Convert "FEVEREIRO/25" -> "2025-02"
function monthLabelToKey(labelNorm){
  // accepts "FEVEREIRO/25", "JAN/2023", "JANEIRO/23"
  const parts = labelNorm.split("/");
  if(parts.length < 2) return null;

  const mPart = parts[0].trim();
  const yPart = parts[1].trim();

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

  let mm = map[mPart] || map[mPart.slice(0,3)] || null;
  if(!mm) return null;

  let yyyy = yPart;
  if(yyyy.length === 2){
    // assume 20x
