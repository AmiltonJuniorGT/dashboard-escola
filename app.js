// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_SELECTED = ["CAMACARI"];

// ================= KPI MATCH (Coluna A) =================
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
  FAT_T: ["FATURAMENTO T (R$)"],
  FAT_P: ["FATURAMENTO P (R$)"],
  FAT:   ["FATURAMENTO"],

  CUSTO: ["CUSTO OPERACIONAL R$"],
  LUCRO: ["LUCRO OPERACIONAL R$"],
  RESULTADO_LIQ: ["RESULTADO LIQUIDO TOTAL R$"],

  MAT: ["MATRICULAS REALIZADAS"],
  CAD: ["CADASTROS"],
  RECEBIDAS_T: ["PARCELAS RECEBIDAS DO MES T"],
  RECEBIDAS_P: ["PARCELAS RECEBIDAS DO MES P"],
  RECEBIDAS: ["PARCELAS RECEBIDAS DO MES", "PARCELAS RECEBIDAS"],

  ATIVOS_T: ["ALUNOS ATIVOS T"],
  ATIVOS_P: ["ALUNOS ATIVOS P"],
  ATIVOS: ["ATIVOS"],

  INAD: ["INADIMPLENCIA"],
  EVASAO: ["EVASAO REAL"],
  MARGEM: ["MARGEM"],
};

// ================= STATE =================
const cache = new Map(); // unitKey -> { rows, colLabels, headerMonths }
let selectedUnits = [...DEFAULT_SELECTED];

// ================= DOM =================
const $ = (id) => document.getElementById(id);

const elStatus = $("status");
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

const unitBtn = $("unitBtn");
const unitMenu = $("unitMenu");
const unitSummary = $("unitSummary");

const WRAP = $("tableWrap");
const TABLE = $("kpiTable");

// Cards containers
const CARD = {
  faturamento: $("card_faturamento"),
  custo: $("card_custo"),
  resultado: $("card_resultado"),
  ativos: $("card_ativos"),
  matriculas: $("card_matriculas"),
  cadastros: $("card_cadastros"),
  conversao: $("card_conversao"),
  recebidas: $("card_recebidas"),
  inad: $("card_inad"),
  evasao: $("card_evasao"),
  margem: $("card_margem"),
};
const TOTAL = {
  faturamento: $("total_faturamento"),
  custo: $("total_custo"),
  resultado: $("total_resultado"),
  ativos: $("total_ativos"),
  matriculas: $("total_matriculas"),
  cadastros: $("total_cadastros"),
  conversao: $("total_conversao"),
  recebidas: $("total_recebidas"),
  inad: $("total_inad"),
  evasao: $("total_evasao"),
  margem: $("total_margem"),
};

// ================= INIT =================
document.addEventListener("DOMContentLoaded", () => setupUI());

function setStatus(msg){ elStatus.textContent = msg; }

function setupUI() {
  const now = new Date();
  elStart.value = toISODate(new Date(now.getFullYear(),0,1));
  elEnd.value = toISODate(new Date(now.getFullYear(),11,31));

  renderUnitMenu();
  refreshUnitSummary();

  // ✅ toggle robusto (não depende de CSS/propagação)
  unitBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    unitMenu.classList.toggle("open");
    unitBtn.setAttribute("aria-expanded", unitMenu.classList.contains("open") ? "true" : "false");
  });

  // fecha ao clicar fora
  document.addEventListener("click", (e) => {
    const inside = e.target.closest("#multiSelect");
    if(!inside){
      unitMenu.classList.remove("open");
      unitBtn.setAttribute("aria-expanded", "false");
    }
  });

  btnLoad.addEventListener("click", () => loadSelectedUnits());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "Carregando dados…";
  loadSelectedUnits();
}

// ================= UI HELPERS =================
function prettyUnit(unitKey){
  return TABS[unitKey].replace("Números ","");
}
function refreshUnitSummary(){
  unitSummary.textContent = selectedUnits.length
    ? selectedUnits.map(prettyUnit).join(" • ")
    : "Selecione ao menos 1 unidade";
}
function renderUnitMenu(){
  const keys = Object.keys(TABS);
  unitMenu.innerHTML = keys.map(k => {
    const checked = selectedUnits.includes(k) ? "checked" : "";
    return `
      <label class="chkRow">
        <input type="checkbox" data-unit="${k}" ${checked}/>
        <span>${prettyUnit(k)}</span>
      </label>
    `;
  }).join("") + `
    <div class="multiFoot">
      <button class="linkBtn" id="selAll" type="button">Todos</button>
      <button class="linkBtn" id="selBase" type="button">Só Camaçari</button>
    </div>
  `;

  unitMenu.querySelectorAll('input[type="checkbox"]').forEach(chk => {
    chk.addEventListener("change", () => {
      const u = chk.dataset.unit;
      if(chk.checked){
        if(!selectedUnits.includes(u)) selectedUnits.push(u);
      } else {
        selectedUnits = selectedUnits.filter(x => x !== u);
      }
      if(!selectedUnits.length){
        selectedUnits = ["CAMACARI"];
        // re-render pra refletir
        renderUnitMenu();
      }
      refreshUnitSummary();
    });
  });

  unitMenu.querySelector("#selAll").addEventListener("click", (e) => {
    e.stopPropagation();
    selectedUnits = Object.keys(TABS);
    renderUnitMenu();
    refreshUnitSummary();
  });
  unitMenu.querySelector("#selBase").addEventListener("click", (e) => {
    e.stopPropagation();
    selectedUnits = ["CAMACARI"];
    renderUnitMenu();
    refreshUnitSummary();
  });
}

// ================= DATE HELPERS =================
function toISODate(d){
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${mm}-${dd}`;
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
function addMonths(d, delta){
  return new Date(d.getFullYear(), d.getMonth()+delta, 1);
}

// ================= FORMAT =================
function cleanNumber(v){
  if(v == null || v === "") return null;
  if(typeof v === "number") return v;
  const s = String(v).trim();
  const cleaned = s.replace(/\./g,"").replace(",",".").replace(/[^\d\.\-]/g,"");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}
function fmtMoney(n){
  if(n == null) return "—";
  return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
}
function fmtInt(n){
  if(n == null) return "—";
  const v = Math.round(n);
  return String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}
function fmtPct(n){
  if(n == null) return "—";
  return `${n.toFixed(2).replace(".",",")}%`;
}
function pctDelta(base, comp){
  if(base == null || comp == null || comp === 0) return null;
  return ((base - comp)/comp)*100;
}
function deltaCls(pct, lowerIsBetter=false){
  if(pct == null) return "";
  if(lowerIsBetter){
    if(pct < 0) return "good";
    if(pct > 0) return "bad";
    return "";
  }
  if(pct > 0) return "good";
  if(pct < 0) return "bad";
  return "";
}

// ================= GVIZ =================
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

// ================= MONTH DETECTION =================
const MONTH_MAP = {
  JAN:"01",JANEIRO:"01",FEV:"02",FEVEREIRO:"02",MAR:"03",MARCO:"03",
  ABR:"04",ABRIL:"04",MAI:"05",MAIO:"05",JUN:"06",JUNHO:"06",
  JUL:"07",JULHO:"07",AGO:"08",AGOSTO:"08",SET:"09",SETEMBRO:"09",
  OUT:"10",OUTUBRO:"10",NOV:"11",NOVEMBRO:"11",DEZ:"12",DEZEMBRO:"12"
};
function monthLabelToKey(raw){
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
function detectHeaderMonths(colLabels){
  const headerMonths = [];
  for(let c=1;c<colLabels.length;c++){
    const label = (colLabels[c] ?? "").trim();
    const key = monthLabelToKey(label);
    if(key) headerMonths.push({ idx:c, label, key });
  }
  return headerMonths;
}
function findMonthColByKey(headerMonths, key){
  const m = headerMonths.find(x => x.key === key);
  return m ? m.idx : null;
}

// ================= LOAD UNITS =================
async function loadSelectedUnits(){
  setStatus("Carregando unidades…");
  hintBox.textContent = "Carregando dados das unidades selecionadas…";

  const tasks = selectedUnits.map(u => loadUnit(u));
  try{
    await Promise.all(tasks);
    setStatus("Dados carregados.");
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro ao carregar dados.");
    hintBox.textContent = "Erro: " + (err?.message || err);
  }
}

async function loadUnit(unitKey){
  if(cache.has(unitKey)) return;
  const tabName = TABS[unitKey];
  const res = await fetch(sheetUrl(tabName), { cache:"no-store" });
  const txt = await res.text();
  const data = parseGviz(txt);
  if(data.status !== "ok"){
    throw new Error(data.errors?.[0]?.detailed_message || `GViz status != ok (${tabName})`);
  }

  let colLabels = (data.table.cols || []).map(c => (c?.label ?? "").trim());
  let rows = gvizTableToRows(data.table);

  const labelsOk = colLabels.slice(1).some(x => x && x.length);
  if(!labelsOk && rows.length){
    colLabels = rows[0].map(v => String(v ?? "").trim());
    rows = rows.slice(1);
  }

  const headerMonths = detectHeaderMonths(colLabels);
  cache.set(unitKey, { rows, colLabels, headerMonths });
}

// ================= KPI ACCESS =================
function findRow(rows, kpiKey){
  const keys = KPI_MATCH[kpiKey] || [];
  for(let r=0;r<rows.length;r++){
    const label = rows[r]?.[0];
    if(labelHas(label, keys)) return rows[r];
  }
  return null;
}
function getVal(rows, kpiKey, colIdx){
  if(colIdx == null) return null;
  const row = findRow(rows, kpiKey);
  return cleanNumber(row?.[colIdx]);
}

// ================= PERIOD CALC =================
function monthRangeKeys(startISO, endISO){
  const s = new Date(startISO+"T00:00:00");
  const e = new Date(endISO+"T00:00:00");
  const a = new Date(s.getFullYear(), s.getMonth(), 1);
  const b = new Date(e.getFullYear(), e.getMonth(), 1);

  const keys = [];
  let cur = new Date(a);
  while(cur <= b){
    keys.push(monthKey(cur));
    cur = addMonths(cur, 1);
  }
  return keys;
}
function buildPeriodIndices(headerMonths, keys){
  return keys.map(k => findMonthColByKey(headerMonths, k)).filter(i => i != null);
}

// Aggregators
function sumPeriod(rows, headerMonths, kpiKey, periodKeys){
  const idxs = buildPeriodIndices(headerMonths, periodKeys);
  if(!idxs.length) return null;
  let total = 0, any = false;
  for(const idx of idxs){
    const v = getVal(rows, kpiKey, idx);
    if(v != null){ total += v; any = true; }
  }
  return any ? total : null;
}
function avgPctPeriod(rows, headerMonths, kpiKey, periodKeys){
  const idxs = buildPeriodIndices(headerMonths, periodKeys);
  if(!idxs.length) return null;
  let sum = 0, n = 0;
  for(const idx of idxs){
    let v = getVal(rows, kpiKey, idx);
    if(v == null) continue;
    if(v <= 1.5) v = v*100;
    sum += v; n++;
  }
  return n ? (sum/n) : null;
}
function ativosPeriod(rows, headerMonths, periodKeys){
  const idxs = buildPeriodIndices(headerMonths, periodKeys);
  if(!idxs.length) return null;
  let sum = 0, n = 0;
  for(const idx of idxs){
    const t = getVal(rows, "ATIVOS_T", idx);
    const p = getVal(rows, "ATIVOS_P", idx);
    const a = getVal(rows, "ATIVOS", idx);
    const v = (t!=null || p!=null) ? ((t||0) + (p||0)) : a;
    if(v != null){ sum += v; n++; }
  }
  return n ? (sum/n) : null;
}
function faturamentoPeriod(rows, headerMonths, periodKeys){
  const t = sumPeriod(rows, headerMonths, "FAT_T", periodKeys);
  const p = sumPeriod(rows, headerMonths, "FAT_P", periodKeys);
  if(t!=null || p!=null) return (t||0) + (p||0);
  return sumPeriod(rows, headerMonths, "FAT", periodKeys);
}
function recebidasPeriod(rows, headerMonths, periodKeys){
  const t = sumPeriod(rows, headerMonths, "RECEBIDAS_T", periodKeys);
  const p = sumPeriod(rows, headerMonths, "RECEBIDAS_P", periodKeys);
  if(t!=null || p!=null) return (t||0) + (p||0);
  return sumPeriod(rows, headerMonths, "RECEBIDAS", periodKeys);
}
function resultadoPeriod(rows, headerMonths, periodKeys){
  const lucro = sumPeriod(rows, headerMonths, "LUCRO", periodKeys);
  if(lucro != null) return lucro;
  return sumPeriod(rows, headerMonths, "RESULTADO_LIQ", periodKeys);
}
function conversaoPeriod(rows, headerMonths, periodKeys){
  const mat = sumPeriod(rows, headerMonths, "MAT", periodKeys);
  const cad = sumPeriod(rows, headerMonths, "CAD", periodKeys);
  if(mat == null || cad == null || cad === 0) return null;
  return (mat/cad)*100;
}

// ================= APPLY =================
function apply(){
  if(!selectedUnits.length){
    hintBox.textContent = "Selecione ao menos 1 unidade.";
    return;
  }
  for(const u of selectedUnits){
    if(!cache.has(u)){
      hintBox.textContent = "Algumas unidades ainda não carregaram. Clique em Carregar.";
      return;
    }
  }

  const startISO = elStart.value;
  const endISO = elEnd.value;

  metaUnit.textContent = selectedUnits.map(prettyUnit).join(" • ");
  metaPeriod.textContent = `${brDate(startISO)} → ${brDate(endISO)}`;

  const baseKeys = monthRangeKeys(startISO, endISO);
  const monthsCount = baseKeys.length;

  // período anterior equivalente
  const startDate = new Date(startISO+"T00:00:00");
  const endDate = new Date(endISO+"T00:00:00");
  const baseStartMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const prevEndMonth = addMonths(baseStartMonth, -1);
  const prevStartMonth = addMonths(prevEndMonth, -(monthsCount-1));
  const prevKeys = monthRangeKeys(toISODate(prevStartMonth), toISODate(prevEndMonth));

  // mesmo período ano anterior
  const yoyStart = addMonths(baseStartMonth, -12);
  const yoyEnd = addMonths(new Date(endDate.getFullYear(), endDate.getMonth(), 1), -12);
  const yoyKeys = monthRangeKeys(toISODate(yoyStart), toISODate(yoyEnd));

  metaBaseMonth.textContent = baseKeys.at(-1) || "—";
  hintBox.textContent = `Período com ${monthsCount} mês(es). Comparando com período anterior equivalente e mesmo período do ano anterior.`;

  renderCardComparative("faturamento", selectedUnits, (u)=>calcUnit(u, "faturamento", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("custo", selectedUnits, (u)=>calcUnit(u, "custo", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("resultado", selectedUnits, (u)=>calcUnit(u, "resultado", baseKeys, prevKeys, yoyKeys));

  renderCardComparative("ativos", selectedUnits, (u)=>calcUnit(u, "ativos", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("matriculas", selectedUnits, (u)=>calcUnit(u, "matriculas", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("cadastros", selectedUnits, (u)=>calcUnit(u, "cadastros", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("conversao", selectedUnits, (u)=>calcUnit(u, "conversao", baseKeys, prevKeys, yoyKeys));

  renderCardComparative("recebidas", selectedUnits, (u)=>calcUnit(u, "recebidas", baseKeys, prevKeys, yoyKeys));

  renderCardComparative("inad", selectedUnits, (u)=>calcUnit(u, "inad", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("evasao", selectedUnits, (u)=>calcUnit(u, "evasao", baseKeys, prevKeys, yoyKeys));
  renderCardComparative("margem", selectedUnits, (u)=>calcUnit(u, "margem", baseKeys, prevKeys, yoyKeys));

  renderTableForPrimary(selectedUnits[0], startISO, endISO);

  setStatus("Aplicado.");
}

// ================= KPI CALC PER UNIT =================
function prettyUnit(unitKey){
  return TABS[unitKey].replace("Números ","");
}

function calcUnit(unitKey, metric, baseKeys, prevKeys, yoyKeys){
  const { rows, headerMonths } = cache.get(unitKey);
  const lowerIsBetter = (metric === "inad" || metric === "evasao");

  let base=null, prev=null, yoy=null, fmt="";

  switch(metric){
    case "faturamento":
      base = faturamentoPeriod(rows, headerMonths, baseKeys);
      prev = faturamentoPeriod(rows, headerMonths, prevKeys);
      yoy  = faturamentoPeriod(rows, headerMonths, yoyKeys);
      fmt = "money"; break;

    case "custo":
      base = sumPeriod(rows, headerMonths, "CUSTO", baseKeys);
      prev = sumPeriod(rows, headerMonths, "CUSTO", prevKeys);
      yoy  = sumPeriod(rows, headerMonths, "CUSTO", yoyKeys);
      fmt = "money"; break;

    case "resultado":
      base = resultadoPeriod(rows, headerMonths, baseKeys);
      prev = resultadoPeriod(rows, headerMonths, prevKeys);
      yoy  = resultadoPeriod(rows, headerMonths, yoyKeys);
      fmt = "money"; break;

    case "ativos":
      base = ativosPeriod(rows, headerMonths, baseKeys);
      prev = ativosPeriod(rows, headerMonths, prevKeys);
      yoy  = ativosPeriod(rows, headerMonths, yoyKeys);
      fmt = "int"; break;

    case "matriculas":
      base = sumPeriod(rows, headerMonths, "MAT", baseKeys);
      prev = sumPeriod(rows, headerMonths, "MAT", prevKeys);
      yoy  = sumPeriod(rows, headerMonths, "MAT", yoyKeys);
      fmt = "int"; break;

    case "cadastros":
      base = sumPeriod(rows, headerMonths, "CAD", baseKeys);
      prev = sumPeriod(rows, headerMonths, "CAD", prevKeys);
      yoy  = sumPeriod(rows, headerMonths, "CAD", yoyKeys);
      fmt = "int"; break;

    case "conversao":
      base = conversaoPeriod(rows, headerMonths, baseKeys);
      prev = conversaoPeriod(rows, headerMonths, prevKeys);
      yoy  = conversaoPeriod(rows, headerMonths, yoyKeys);
      fmt = "pct"; break;

    case "recebidas":
      base = recebidasPeriod(rows, headerMonths, baseKeys);
      prev = recebidasPeriod(rows, headerMonths, prevKeys);
      yoy  = recebidasPeriod(rows, headerMonths, yoyKeys);
      fmt = "money"; break;

    case "inad":
      base = avgPctPeriod(rows, headerMonths, "INAD", baseKeys);
      prev = avgPctPeriod(rows, headerMonths, "INAD", prevKeys);
      yoy  = avgPctPeriod(rows, headerMonths, "INAD", yoyKeys);
      fmt = "pct"; break;

    case "evasao":
      base = avgPctPeriod(rows, headerMonths, "EVASAO", baseKeys);
      prev = avgPctPeriod(rows, headerMonths, "EVASAO", prevKeys);
      yoy  = avgPctPeriod(rows, headerMonths, "EVASAO", yoyKeys);
      fmt = "pct"; break;

    case "margem":
      base = avgPctPeriod(rows, headerMonths, "MARGEM", baseKeys);
      prev = avgPctPeriod(rows, headerMonths, "MARGEM", prevKeys);
      yoy  = avgPctPeriod(rows, headerMonths, "MARGEM", yoyKeys);
      fmt = "pct"; break;
  }

  const prevPct = pctDelta(base, prev);
  const yoyPct  = pctDelta(base, yoy);

  return { unitKey, base, prev, yoy, prevPct, yoyPct, fmt, lowerIsBetter };
}

function formatBy(fmt, v){
  if(v == null) return "—";
  if(fmt === "money") return fmtMoney(v);
  if(fmt === "int") return fmtInt(v);
  return fmtPct(v);
}

// ================= CARD RENDER =================
function renderCardComparative(metric, units, calcFn){
  const lines = units.map(u => calcFn(u));
  const total = consolidateTotal(lines);
  TOTAL[metric].textContent = `Total: ${total}`;
  CARD[metric].innerHTML = lines.map(x => renderUnitLine(x)).join("");
}

function consolidateTotal(lines){
  const bases = lines.map(x=>x.base).filter(v=>v!=null);
  if(!bases.length) return "—";
  const fmt = lines[0]?.fmt || "money";
  if(fmt === "money") return fmtMoney(bases.reduce((a,b)=>a+b,0));
  if(fmt === "int") return fmtInt(bases.reduce((a,b)=>a+b,0));
  return fmtPct(bases.reduce((a,b)=>a+b,0) / bases.length);
}

function renderUnitLine(x){
  const prevClass = `pct ${deltaCls(x.prevPct, x.lowerIsBetter)}`;
  const yoyClass  = `pct ${deltaCls(x.yoyPct, x.lowerIsBetter)}`;

  const prevPctTxt = (x.prevPct==null) ? "—" : `${x.prevPct>=0?"+":""}${x.prevPct.toFixed(2).replace(".",",")}%`;
  const yoyPctTxt  = (x.yoyPct==null) ? "—" : `${x.yoyPct>=0?"+":""}${x.yoyPct.toFixed(2).replace(".",",")}%`;

  return `
    <div class="unitLine unit-${x.unitKey}">
      <div class="unitTop">
        <div class="unitName">${prettyUnit(x.unitKey)}</div>
        <div class="unitValue">${formatBy(x.fmt, x.base)}</div>
      </div>
      <div class="unitBottom">
        <span class="tag">Per. ant.: <b class="${prevClass}">${prevPctTxt}</b></span>
        <span class="tag">Ano ant.: <b>${formatBy(x.fmt, x.yoy)}</b> <b class="${yoyClass}">${yoyPctTxt}</b></span>
      </div>
    </div>
  `;
}

// ================= TABLE (PRIMARY UNIT) =================
function renderTableForPrimary(unitKey, startISO, endISO){
  const pack = cache.get(unitKey);
  if(!pack){ TABLE.innerHTML = `<tr><td class="muted">Sem dados.</td></tr>`; return; }

  const { rows, headerMonths } = pack;
  TABLE.innerHTML = "";

  if(!headerMonths.length){
    TABLE.innerHTML = `<tr><td class="muted">Não identifiquei meses no cabeçalho da unidade principal.</td></tr>`;
    return;
  }

  const mode = elMode.value;
  const cols = buildVisibleColumns(mode, headerMonths, startISO, endISO);

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
  for(const row of rows){
    const label = String(row?.[0] ?? "").trim();
    if(!label) continue;

    const tr = document.createElement("tr");
    if(isGroupRow(label, row)) tr.classList.add("trGroup");

    cols.forEach((c,i)=>{
      const td = document.createElement("td");
      const v = row[c.idx];
      td.textContent = (i===0) ? label : (v==null ? "" : String(v));
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
  TABLE.appendChild(tbody);

  tblInfo.textContent = `Tabela: ${prettyUnit(unitKey)} • ${brDate(startISO)} → ${brDate(endISO)}`;
}

function buildVisibleColumns(mode, headerMonths, startISO, endISO){
  const cols = [{ idx:0, label:"Indicador" }];

  if(mode === "full"){
    headerMonths.forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }

  const keys = monthRangeKeys(startISO, endISO);
  const idxs = buildPeriodIndices(headerMonths, keys);

  if(!idxs.length){
    headerMonths.slice(Math.max(0, headerMonths.length-8)).forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }

  const min = Math.min(...idxs), max = Math.max(...idxs);
  const subset = headerMonths.filter(m => m.idx>=min && m.idx<=max);
  subset.forEach(m => cols.push({ idx:m.idx, label:m.label }));
  return cols;
}

function isGroupRow(label, row){
  const L = norm(label);
  const looksUpper = L === label.toUpperCase();
  const hasNumbers = row.slice(1).some(v => v != null && String(v).match(/\d/));
  return (looksUpper && !hasNumbers) || L.includes("PLANILHA DE INDICADORES");
}
