// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_UNIT_KEY = "CAMACARI";

// ================= HELPERS =================
function norm(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}
function labelHas(label, keys) {
  const L = norm(label);
  return (keys || []).some(k => L.includes(norm(k)));
}
function brDate(iso){
  if(!iso) return "—";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function toISODate(d){
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}
function monthKey(date){
  const mm = String(date.getMonth()+1).padStart(2,"0");
  return `${date.getFullYear()}-${mm}`;
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
function fmtPct(p){
  if(p == null) return "—";
  return `${p.toFixed(2).replace(".",",")}%`;
}
function cleanNumber(v){
  if(v == null || v === "") return null;
  if(typeof v === "number") return v;
  const s = String(v).trim();
  const cleaned = s.replace(/\./g,"").replace(",",".").replace(/[^\d\.\-]/g,"");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}
function pctDelta(base, compare){
  if(base == null || compare == null || compare === 0) return null;
  return ((base - compare)/compare)*100;
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

// ================= KPI MATCHES (coluna A) =================
const KPI_MATCH = {
  FAT_T: ["FATURAMENTO T (R$)"],
  FAT_P: ["FATURAMENTO P (R$)"],
  CUSTO: ["CUSTO OPERACIONAL R$"],
  LUCRO: ["LUCRO OPERACIONAL R$"],
  RESULTADO_LIQ: ["RESULTADO LIQUIDO TOTAL R$"],
  ATIVOS: ["ALUNOS ATIVOS T", "ALUNOS ATIVOS P", "ATIVOS"],
  MATRICULAS: ["MATRICULAS REALIZADAS"],
  CADASTROS: ["CADASTROS"],
  INAD: ["INADIMPLENCIA"],
  EVASAO: ["EVASAO REAL"],
  MARGEM: ["MARGEM"],
};

// ================= EXECUTIVE CARDS CONFIG =================
const EXEC_CARDS = [
  { id:"fat", title:"Faturamento", type:"money_sum", keys:["FAT_T","FAT_P"] },
  { id:"custo", title:"Custos", type:"money", key:"CUSTO" },
  { id:"resultado", title:"Resultado", type:"money_fallback", keys:["LUCRO","RESULTADO_LIQ"] },
  { id:"ativos", title:"Ativos", type:"int", key:"ATIVOS" },
  { id:"mat", title:"Matrículas", type:"int", key:"MATRICULAS" },
  { id:"cad", title:"Cadastros", type:"int", key:"CADASTROS" },
  { id:"conv", title:"Conversão", type:"conversion", num:"MATRICULAS", den:"CADASTROS" },
  { id:"inad", title:"Inadimplência", type:"pct", key:"INAD", lowerIsBetter:true },
];

// ================= STATE =================
let currentRows = [];
let colLabels = [];
let headerMonths = []; // [{idx,label,key}]
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
const GRID = $("kpiGrid");
const WRAP = $("tableWrap");
const TABLE = $("kpiTable");

// Drawer
const drawer = $("drawer");
const drawerBackdrop = $("drawerBackdrop");
const drawerClose = $("drawerClose");
const drawerTitle = $("drawerTitle");
const drawerSub = $("drawerSub");
const d_base = $("d_base");
const d_prev = $("d_prev");
const d_yoy = $("d_yoy");
const d_mom = $("d_mom");
const d_yoy_pct = $("d_yoy_pct");
const spark = $("spark");
const drawerTable = $("drawerTable");
const drawerBreakdownBox = $("drawerBreakdownBox");
const drawerBreakdown = $("drawerBreakdown");

// ================= INIT =================
document.addEventListener("DOMContentLoaded", () => setupUI());

function setStatus(msg){ elStatus.textContent = msg; }

function setupUI() {
  // unidade select
  elUnit.innerHTML = Object.entries(TABS).map(([k,v]) =>
    `<option value="${k}">${v.replace("Números ","")}</option>`
  ).join("");
  elUnit.value = DEFAULT_UNIT_KEY;

  // datas padrão: ano atual
  const now = new Date();
  elStart.value = toISODate(new Date(now.getFullYear(),0,1));
  elEnd.value = toISODate(new Date(now.getFullYear(),11,31));

  btnLoad.addEventListener("click", () => loadData());
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  drawerClose.addEventListener("click", closeDrawer);
  drawerBackdrop.addEventListener("click", closeDrawer);
  document.addEventListener("keydown", (e) => { if(e.key === "Escape") closeDrawer(); });

  // cria cards executivos
  renderExecutiveCards();

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "Carregando dados…";
  loadData();
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

async function loadData(){
  const tabName = TABS[elUnit.value];
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

    // fallback: se labels vierem vazios, usa primeira linha como cabeçalho
    const labelsOk = colLabels.slice(1).some(x => x && x.length > 0);
    if(!labelsOk && currentRows.length){
      colLabels = currentRows[0].map(v => String(v ?? "").trim());
      currentRows = currentRows.slice(1);
    }

    detectHeaderMonths();
    setStatus(`Dados carregados. (${tabName})`);
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro ao carregar dados.");
    hintBox.textContent = "Erro: " + (err?.message || err);
  }
}

// ================= MONTH DETECTION =================
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
function detectHeaderMonths(){
  headerMonths = [];
  for(let c=1;c<colLabels.length;c++){
    const label = colLabels[c];
    const key = monthLabelToKey(label);
    if(key) headerMonths.push({ idx:c, label, key });
  }
  if(!headerMonths.length){
    hintBox.textContent = "⚠️ Não detectei meses no cabeçalho. Confirme o formato (ex: jan/25).";
  } else {
    hintBox.textContent = `✅ Meses: ${headerMonths.length} (ex: ${headerMonths[0].label} … ${headerMonths.at(-1).label})`;
  }
}
function findMonthColByKey(key){
  const m = headerMonths.find(x => x.key === key);
  return m ? m.idx : null;
}

// ================= DATA ACCESS =================
function findRow(kpiKey){
  const keys = KPI_MATCH[kpiKey] || [];
  for(let r=0;r<currentRows.length;r++){
    const label = currentRows[r]?.[0];
    if(labelHas(label, keys)) return currentRows[r];
  }
  return null;
}
function getVal(kpiKey, colIdx){
  const row = findRow(kpiKey);
  return cleanNumber(row?.[colIdx]);
}

// ================= EXEC CARD RENDER =================
function renderExecutiveCards(){
  GRID.innerHTML = "";
  EXEC_CARDS.forEach(cfg => {
    const card = document.createElement("div");
    card.className = "kpiCard";
    card.dataset.cardId = cfg.id;

    card.innerHTML = `
      <div class="kpiTitle">${cfg.title}</div>
      <div class="kpiValue" id="val_${cfg.id}">—</div>
      <div class="kpiSub">
        <div class="kpiDelta" id="mom_${cfg.id}">vs mês —</div>
        <div class="kpiDelta" id="yoy_${cfg.id}">vs ano —</div>
      </div>
    `;

    card.addEventListener("click", () => openAnalysis(cfg));
    GRID.appendChild(card);
  });
}

function apply(){
  const tabName = TABS[elUnit.value];
  metaUnit.textContent = tabName.replace("Números ","");
  metaPeriod.textContent = `${brDate(elStart.value)} → ${brDate(elEnd.value)}`;

  if(!headerMonths.length){
    metaBaseMonth.textContent = "—";
    return;
  }

  // mês base = fim do período (se não existir, pega último disponível)
  const baseDate = new Date(elEnd.value + "T00:00:00");
  const baseKey = monthKey(baseDate);
  selectedMonthIdx = findMonthColByKey(baseKey);
  if(selectedMonthIdx == null) selectedMonthIdx = headerMonths.at(-1).idx;

  selectedMonthLabel = colLabels[selectedMonthIdx] || headerMonths.find(m=>m.idx===selectedMonthIdx)?.label || "—";
  metaBaseMonth.textContent = selectedMonthLabel;

  // indices prev / yoy
  const prevDate = new Date(baseDate.getFullYear(), baseDate.getMonth()-1, 1);
  const yoyDate  = new Date(baseDate.getFullYear()-1, baseDate.getMonth(), 1);
  const prevIdx = findMonthColByKey(monthKey(prevDate));
  const yoyIdx  = findMonthColByKey(monthKey(yoyDate));

  // preenche cards
  EXEC_CARDS.forEach(cfg => {
    const { base, prev, yoy, lowerIsBetter } = computeCard(cfg, selectedMonthIdx, prevIdx, yoyIdx);
    const valEl = $(`val_${cfg.id}`);
    const momEl = $(`mom_${cfg.id}`);
    const yoyEl = $(`yoy_${cfg.id}`);

    valEl.textContent = formatCardValue(cfg, base);

    const dm = pctDelta(base, prev);
    momEl.textContent = dm == null ? "vs mês —" : `vs mês ${dm>=0?"+":""}${dm.toFixed(2).replace(".",",")}%`;
    momEl.className = `kpiDelta ${deltaClass(dm, !!lowerIsBetter)}`;

    const dy = pctDelta(base, yoy);
    yoyEl.textContent = dy == null ? "vs ano —" : `vs ano ${dy>=0?"+":""}${dy.toFixed(2).replace(".",",")}%`;
    yoyEl.className = `kpiDelta ${deltaClass(dy, !!lowerIsBetter)}`;
  });

  renderTable();
  tblInfo.textContent = `Base: ${selectedMonthLabel} • Meses: ${headerMonths.length}`;
}

function computeCard(cfg, baseIdx, prevIdx, yoyIdx){
  let base=null, prev=null, yoy=null;

  if(cfg.type === "money_sum"){
    base = sumNullable(cfg.keys.map(k=>getVal(k, baseIdx)));
    prev = sumNullable(cfg.keys.map(k=>getVal(k, prevIdx)));
    yoy  = sumNullable(cfg.keys.map(k=>getVal(k, yoyIdx)));
  } else if(cfg.type === "money_fallback"){
    const a = getVal(cfg.keys[0], baseIdx);
    const b = getVal(cfg.keys[1], baseIdx);
    base = (a!=null ? a : b);

    const ap = getVal(cfg.keys[0], prevIdx);
    const bp = getVal(cfg.keys[1], prevIdx);
    prev = (ap!=null ? ap : bp);

    const ay = getVal(cfg.keys[0], yoyIdx);
    const by = getVal(cfg.keys[1], yoyIdx);
    yoy = (ay!=null ? ay : by);
  } else if(cfg.type === "conversion"){
    const nB = getVal(cfg.num, baseIdx);
    const dB = getVal(cfg.den, baseIdx);
    base = (nB!=null && dB!=null && dB!==0) ? (nB/dB)*100 : null;

    const nP = getVal(cfg.num, prevIdx);
    const dP = getVal(cfg.den, prevIdx);
    prev = (nP!=null && dP!=null && dP!==0) ? (nP/dP)*100 : null;

    const nY = getVal(cfg.num, yoyIdx);
    const dY = getVal(cfg.den, yoyIdx);
    yoy = (nY!=null && dY!=null && dY!==0) ? (nY/dY)*100 : null;
  } else {
    base = getVal(cfg.key, baseIdx);
    prev = getVal(cfg.key, prevIdx);
    yoy  = getVal(cfg.key, yoyIdx);
  }

  return { base, prev, yoy, lowerIsBetter: cfg.lowerIsBetter };
}

function formatCardValue(cfg, v){
  if(v == null) return "—";
  if(cfg.type === "int") return fmtInt(v);
  if(cfg.type === "pct" || cfg.type === "conversion") return fmtPct(v);
  return fmtMoney(v);
}

function sumNullable(arr){
  const nums = arr.filter(x => x != null);
  if(!nums.length) return null;
  return nums.reduce((a,b)=>a+b,0);
}

// ================= TABLE (mesma planilha, para auditoria) =================
function renderTable(){
  TABLE.innerHTML = "";
  if(!headerMonths.length){
    TABLE.innerHTML = `<tr><td class="muted">Sem meses detectados.</td></tr>`;
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
  for(const row of currentRows){
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
    headerMonths.slice(Math.max(0, headerMonths.length-8)).forEach(m => cols.push({ idx:m.idx, label:m.label }));
    return cols;
  }
  const a = Math.min(aPos,bPos);
  const b = Math.max(aPos,bPos);
  for(let i=a;i<=b;i++) cols.push({ idx: headerMonths[i].idx, label: headerMonths[i].label });
  return cols;
}

function isGroupRow(label, row){
  const L = norm(label);
  const looksUpper = L === label.toUpperCase();
  const hasNumbers = row.slice(1).some(v => v != null && String(v).match(/\d/));
  return (looksUpper && !hasNumbers) || L.includes("PLANILHA DE INDICADORES");
}

// ================= DRAWER ANALYSIS =================
function openAnalysis(cfg){
  if(!headerMonths.length) return;

  const tabName = TABS[elUnit.value].replace("Números ","");
  drawerTitle.textContent = cfg.title;
  drawerSub.textContent = `${tabName} • Base: ${selectedMonthLabel}`;

  const baseDate = new Date(elEnd.value + "T00:00:00");
  const prevDate = new Date(baseDate.getFullYear(), baseDate.getMonth()-1, 1);
  const yoyDate  = new Date(baseDate.getFullYear()-1, baseDate.getMonth(), 1);
  const prevIdx = findMonthColByKey(monthKey(prevDate));
  const yoyIdx  = findMonthColByKey(monthKey(yoyDate));

  const { base, prev, yoy, lowerIsBetter } = computeCard(cfg, selectedMonthIdx, prevIdx, yoyIdx);

  d_base.textContent = formatCardValue(cfg, base);
  d_prev.textContent = formatCardValue(cfg, prev);
  d_yoy.textContent  = formatCardValue(cfg, yoy);

  const dm = pctDelta(base, prev);
  d_mom.textContent = dm==null ? "—" : `${dm>=0?"+":""}${dm.toFixed(2).replace(".",",")}%`;
  d_mom.className = `dValue ${deltaClass(dm, !!lowerIsBetter)}`;

  const dy = pctDelta(base, yoy);
  d_yoy_pct.textContent = dy==null ? "—" : `${dy>=0?"+":""}${dy.toFixed(2).replace(".",",")}%`;
  d_yoy_pct.className = `dValue ${deltaClass(dy, !!lowerIsBetter)}`;

  // últimos 12 meses (ou menos, se não tiver)
  const series = buildLastMonthsSeries(cfg, 12);
  renderDrawerTable(series, cfg);
  renderSpark(series);

  // breakdown T/P (quando for faturamento)
  if(cfg.id === "fat"){
    renderBreakdownTP(selectedMonthIdx, prevIdx);
  } else {
    drawerBreakdownBox.style.display = "none";
    drawerBreakdown.innerHTML = "";
  }

  drawerBackdrop.classList.add("open");
  drawer.classList.add("open");
}

function closeDrawer(){
  drawerBackdrop.classList.remove("open");
  drawer.classList.remove("open");
}

function buildLastMonthsSeries(cfg, n){
  // pega últimos n meses detectados
  const slice = headerMonths.slice(Math.max(0, headerMonths.length - n));
  return slice.map(m => {
    const { base } = computeCard(cfg, m.idx, null, null);
    return { label: m.label, value: base };
  });
}

function renderDrawerTable(series, cfg){
  drawerTable.innerHTML = "";
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr><th>Mês</th><th>Valor</th></tr>`;
  drawerTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  for(const it of series){
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${it.label}</td><td>${formatCardValue(cfg, it.value)}</td>`;
    tbody.appendChild(tr);
  }
  drawerTable.appendChild(tbody);
}

function renderSpark(series){
  const ctx = spark.getContext("2d");
  ctx.clearRect(0,0,spark.width,spark.height);

  const values = series.map(s => (s.value==null ? null : Number(s.value))).filter(v => v!=null);
  if(values.length < 2){
    ctx.font = "12px system-ui";
    ctx.fillStyle = "#6b7280";
    ctx.fillText("Sem dados suficientes para gráfico.", 10, 22);
    return;
  }

  const min = Math.min(...values);
  const max = Math.max(...values);
  const pad = 14;
  const W = spark.width - pad*2;
  const H = spark.height - pad*2;

  function x(i){ return pad + (W * (i/(series.length-1))); }
  function y(v){
    if(max === min) return pad + H/2;
    return pad + (H * (1 - (v - min)/(max - min)));
  }

  // linha base
  ctx.strokeStyle = "#e5e7eb";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(pad, pad+H);
  ctx.lineTo(pad+W, pad+H);
  ctx.stroke();

  // linha série
  ctx.strokeStyle = "#111827";
  ctx.lineWidth = 2;
  ctx.beginPath();
  let started = false;
  series.forEach((s,i)=>{
    if(s.value==null) return;
    const xi = x(i);
    const yi = y(s.value);
    if(!started){ ctx.moveTo(xi,yi); started=true; }
    else ctx.lineTo(xi,yi);
  });
  ctx.stroke();

  // último ponto
  const lastIdx = [...series].reverse().findIndex(s=>s.value!=null);
  if(lastIdx !== -1){
    const i = series.length - 1 - lastIdx;
    ctx.fillStyle = "#111827";
    ctx.beginPath();
    ctx.arc(x(i), y(series[i].value), 3.5, 0, Math.PI*2);
    ctx.fill();
  }

  // min/max labels
  ctx.font = "11px system-ui";
  ctx.fillStyle = "#6b7280";
  ctx.fillText("min", pad, pad+10);
  ctx.fillText("max", pad, pad+H-6);
}

// Breakdown T/P para faturamento
function renderBreakdownTP(nowIdx, prevIdx){
  const tNow = getVal("FAT_T", nowIdx);
  const pNow = getVal("FAT_P", nowIdx);
  const tPrev = getVal("FAT_T", prevIdx);
  const pPrev = getVal("FAT_P", prevIdx);

  const totNow = sumNullable([tNow,pNow]);
  const totPrev = sumNullable([tPrev,pPrev]);

  drawerBreakdown.innerHTML = `
    <div class="breakItem"><div class="t">T (base)</div><div class="v">${fmtMoney(tNow)}</div></div>
    <div class="breakItem"><div class="t">P (base)</div><div class="v">${fmtMoney(pNow)}</div></div>
    <div class="breakItem"><div class="t">Total (base)</div><div class="v">${fmtMoney(totNow)}</div></div>

    <div class="breakItem"><div class="t">T (mês ant.)</div><div class="v">${fmtMoney(tPrev)}</div></div>
    <div class="breakItem"><div class="t">P (mês ant.)</div><div class="v">${fmtMoney(pPrev)}</div></div>
    <div class="breakItem"><div class="t">Total (mês ant.)</div><div class="v">${fmtMoney(totPrev)}</div></div>
  `;

  drawerBreakdownBox.style.display = "block";
}
