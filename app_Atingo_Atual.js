// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_SELECTED = ["CAMACARI"];

// ================= NORMALIZE =================
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

// ================= KPI MATCH (Coluna A) =================
// (mantém flexível, porque os nomes podem ter variações)
const KPI_MATCH = {
  // Financeiro
  FAT_T: ["FATURAMENTO T (R$)"],
  FAT_P: ["FATURAMENTO P (R$)"],
  FAT:   ["FATURAMENTO"],

  CUSTO: ["CUSTO OPERACIONAL R$"],
  LUCRO: ["LUCRO OPERACIONAL R$"],
  RESULTADO_LIQ: ["RESULTADO LIQUIDO TOTAL R$", "RESULTADO LIQUIDO TOTAL"],

  // Volume
  MAT: ["MATRICULAS REALIZADAS", "MATRÍCULAS REALIZADAS"],
  CAD: ["CADASTROS", "CADASTRO"],
  RECEBIDAS_T: ["PARCELAS RECEBIDAS DO MES T", "PARCELAS RECEBIDAS DO MÊS T"],
  RECEBIDAS_P: ["PARCELAS RECEBIDAS DO MES P", "PARCELAS RECEBIDAS DO MÊS P"],
  RECEBIDAS: ["PARCELAS RECEBIDAS DO MES", "PARCELAS RECEBIDAS", "PARCELAS RECEBIDAS DO MÊS"],

  ATIVOS_T: ["ALUNOS ATIVOS T"],
  ATIVOS_P: ["ALUNOS ATIVOS P"],
  ATIVOS: ["ATIVOS", "ALUNOS ATIVOS"],

  // Saúde
  INAD: ["INADIMPLENCIA", "INADIMPLÊNCIA"],
  EVASAO: ["EVASAO REAL", "EVASÃO REAL"],
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
const hintBox = $("hintBox");

const metaUnit = $("metaUnit");
const metaPeriod = $("metaPeriod");
const metaBaseMonth = $("metaBaseMonth");

const unitBtn = $("unitBtn");
const unitBtnLabel = $("unitBtnLabel");
const unitMenu = $("unitMenu");
const unitSummary = $("unitSummary");

const WRAP = $("tableWrap");
const TABLE = $("kpiTable");
const tblInfo = $("tblInfo");

// Cards containers (todas as linhas)
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

function setStatus(msg){ if(elStatus) elStatus.textContent = msg; }

// ================= UI SETUP =================
function setupUI() {
  const now = new Date();
  elStart.value = toISODate(new Date(now.getFullYear(),0,1));
  elEnd.value = toISODate(new Date(now.getFullYear(),11,31));

  renderUnitMenu();
  refreshUnitSummary();

  // dropdown robusto
  unitBtn?.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    unitMenu.classList.toggle("open");
    unitBtn.setAttribute("aria-expanded", unitMenu.classList.contains("open") ? "true" : "false");
  });

  document.addEventListener("click", (e) => {
    const inside = e.target.closest("#multiSelect");
    if(!inside){
      unitMenu.classList.remove("open");
      unitBtn?.setAttribute("aria-expanded", "false");
    }
  });

  btnLoad?.addEventListener("click", () => loadSelectedUnits(true));
  btnApply?.addEventListener("click", () => apply());
  btnExpand?.addEventListener("click", () => WRAP?.classList.toggle("expanded"));

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "Carregando dados…";
  loadSelectedUnits(false);
}

// ================= UI HELPERS =================
function prettyUnit(unitKey){ return TABS[unitKey].replace("Números ",""); }

function refreshUnitSummary(){
  const txt = selectedUnits.length
    ? selectedUnits.map(prettyUnit).join(" • ")
    : "Selecione ao menos 1 unidade";
  if(unitSummary) unitSummary.textContent = txt;
  if(unitBtnLabel) unitBtnLabel.textContent = selectedUnits.length ? `${selectedUnits.length} unidade(s)` : "Selecionar unidades";
}

function renderUnitMenu(){
  if(!unitMenu) return;
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
        renderUnitMenu();
      }
      refreshUnitSummary();
    });
  });

  unitMenu.querySelector("#selAll")?.addEventListener("click", (e) => {
    e.stopPropagation();
    selectedUnits = Object.keys(TABS);
    renderUnitMenu();
    refreshUnitSummary();
  });

  unitMenu.querySelector("#selBase")?.addEventListener("click", (e) => {
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
function monthsN(keys){ return (keys && keys.length) ? keys.length : 1; }
function prevMonthKey(startISO){
  const d = new Date(startISO+"T00:00:00");
  const startMonth = new Date(d.getFullYear(), d.getMonth(), 1);
  return monthKey(addMonths(startMonth, -1));
}
function shiftKeys(keys, deltaMonths){
  if(!keys.length) return [];
  const first = new Date(keys[0] + "-01T00:00:00");
  const last = new Date(keys[keys.length-1] + "-01T00:00:00");
  const a = addMonths(new Date(first.getFullYear(), first.getMonth(), 1), deltaMonths);
  const b = addMonths(new Date(last.getFullYear(), last.getMonth(), 1), deltaMonths);
  return monthRangeKeys(toISODate(a), toISODate(b));
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
function buildPeriodIndices(headerMonths, keys){
  return keys.map(k => findMonthColByKey(headerMonths, k)).filter(i => i != null);
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

// ================= PERIOD AGGREGATORS =================
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
    // se vier como 0.09 em vez de 9%
    if(v <= 1.5) v = v*100;
    sum += v; n++;
  }
  return n ? (sum/n) : null;
}

// Ativos: preferir T+P, senão "ATIVOS"
function ativosPeriodAvg(rows, headerMonths, periodKeys){
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
  return n ? (sum/n) : null; // média mensal
}

// Financeiro: Faturamento = T+P (se existir) senão FAT
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

// month values para “mês anterior” no Financeiro
function faturamentoMonth(rows, headerMonths, keyYYYYMM){
  const col = findMonthColByKey(headerMonths, keyYYYYMM);
  if(col == null) return null;
  const t = getVal(rows, "FAT_T", col);
  const p = getVal(rows, "FAT_P", col);
  if(t!=null || p!=null) return (t||0) + (p||0);
  return getVal(rows, "FAT", col);
}
function custoMonth(rows, headerMonths, keyYYYYMM){
  const col = findMonthColByKey(headerMonths, keyYYYYMM);
  if(col == null) return null;
  return getVal(rows, "CUSTO", col);
}
function resultadoMonth(rows, headerMonths, keyYYYYMM){
  const col = findMonthColByKey(headerMonths, keyYYYYMM);
  if(col == null) return null;
  const lucro = getVal(rows, "LUCRO", col);
  if(lucro != null) return lucro;
  return getVal(rows, "RESULTADO_LIQ", col);
}

// Conversão do período = matrículas / cadastros
function conversaoPeriod(rows, headerMonths, periodKeys){
  const mat = sumPeriod(rows, headerMonths, "MAT", periodKeys);
  const cad = sumPeriod(rows, headerMonths, "CAD", periodKeys);
  if(mat == null || cad == null || cad === 0) return null;
  return (mat/cad)*100;
}

// ================= LOAD UNITS =================
async function loadSelectedUnits(forceReload=false){
  setStatus("Carregando unidades…");
  hintBox.textContent = "Carregando dados das unidades selecionadas…";

  try{
    if(forceReload){
      for(const u of selectedUnits) cache.delete(u);
    }
    await Promise.all(selectedUnits.map(u => loadUnit(u)));
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

  // fallback: se labels vazios e primeira linha é cabeçalho
  const labelsOk = colLabels.slice(1).some(x => x && x.length);
  if(!labelsOk && rows.length){
    colLabels = rows[0].map(v => String(v ?? "").trim());
    rows = rows.slice(1);
  }

  const headerMonths = detectHeaderMonths(colLabels);
  cache.set(unitKey, { rows, colLabels, headerMonths });
}

// ================= RENDER (MANTÉM PADRÃO ANTIGO NAS LINHAS ABAIXO) =================

// Finance (modelo do seu print)
function unitLineFinanceHTML(unitKey, title, mediaMensal, mesAnteriorValor, pctVsMesAnt, mediaAnoAnt, pctVsAnoAnt){
  const pct1cls = deltaCls(pctVsMesAnt);
  const pct2cls = deltaCls(pctVsAnoAnt);

  return `
    <div class="unitLine unit-${unitKey}">
      <div class="unitTop">
        <div class="unitName">${title}</div>
        <div class="unitValue">${mediaMensal}</div>
      </div>

      <div class="unitBottom">
        <div class="rowLine">
          <span>Per. ant.: <b>${mesAnteriorValor}</b></span>
          <span class="pct ${pct1cls}">${fmtPct(pctVsMesAnt)}</span>
        </div>
        <div class="rowLine">
          <span>Ano ant.: <b>${mediaAnoAnt}</b></span>
          <span class="pct ${pct2cls}">${fmtPct(pctVsAnoAnt)}</span>
        </div>
      </div>
    </div>
  `;
}

// Linhas abaixo (mantém “como antes”: valor + % vs mês anterior + % vs ano anterior)
function unitLineSimpleHTML(unitKey, title, valueStr, pctMoM, pctYoY, lowerIsBetter=false){
  const pct1cls = deltaCls(pctMoM, lowerIsBetter);
  const pct2cls = deltaCls(pctYoY, lowerIsBetter);

  return `
    <div class="unitLine unit-${unitKey}">
      <div class="unitTop">
        <div class="unitName">${title}</div>
        <div class="unitValue">${valueStr}</div>
      </div>

      <div class="unitBottom">
        <div class="rowLine">
          <span>vs mês anterior</span>
          <span class="pct ${pct1cls}">${fmtPct(pctMoM)}</span>
        </div>
        <div class="rowLine">
          <span>vs ano anterior</span>
          <span class="pct ${pct2cls}">${fmtPct(pctYoY)}</span>
        </div>
      </div>
    </div>
  `;
}

function setTotalPill(el, value){
  if(!el) return;
  el.textContent = `Total: ${value}`;
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

  const baseKeys = monthRangeKeys(startISO, endISO);
  const nMonths = monthsN(baseKeys);

  // período anterior equivalente (mesma quantidade de meses, terminando no mês anterior ao início)
  const startDate = new Date(startISO+"T00:00:00");
  const baseStartMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const prevEndMonth = addMonths(baseStartMonth, -1);
  const prevStartMonth = addMonths(prevEndMonth, -(nMonths-1));
  const prevKeys = monthRangeKeys(toISODate(prevStartMonth), toISODate(prevEndMonth));

  // mesmo período ano anterior (-12 meses)
  const yoyKeys = shiftKeys(baseKeys, -12);

  // mês anterior (somente para Finance no seu modelo)
  const pmKey = prevMonthKey(startISO);

  // metas topo
  metaUnit.textContent = selectedUnits.map(prettyUnit).join(", ");
  metaPeriod.textContent = `${brDate(startISO)} → ${brDate(endISO)}`;
  metaBaseMonth.textContent = baseKeys.length ? baseKeys[baseKeys.length-1] : "—";

  // aviso cabeçalhos
  const baseUnit = selectedUnits[0];
  const baseData = cache.get(baseUnit);
  if(!baseData?.headerMonths?.length){
    hintBox.textContent = "⚠️ Não detectei meses nos cabeçalhos. Confirme se está como 'jan./25', 'fev./25' etc.";
  } else {
    hintBox.textContent = "";
  }

  // ================== FINANCEIRO (PRINT) ==================
  let htmlFat = "", htmlCusto = "", htmlRes = "";

  let sumFatAll = 0, sumCustoAll = 0, sumResAll = 0;
  let anyFat=false, anyCusto=false, anyRes=false;

  for(const unitKey of selectedUnits){
    const { rows, headerMonths } = cache.get(unitKey);

    // Faturamento
    const fatTotal = faturamentoPeriod(rows, headerMonths, baseKeys);
    const fatMedia = (fatTotal == null) ? null : fatTotal / nMonths;
    const fatMesAnt = faturamentoMonth(rows, headerMonths, pmKey);
    const fatVsMesAnt = pctDelta(fatMedia, fatMesAnt);

    const fatYoyTotal = faturamentoPeriod(rows, headerMonths, yoyKeys);
    const fatYoyMedia = (fatYoyTotal == null) ? null : fatYoyTotal / nMonths;
    const fatVsYoy = pctDelta(fatMedia, fatYoyMedia);

    htmlFat += unitLineFinanceHTML(
      unitKey,
      prettyUnit(unitKey),
      fatMedia == null ? "—" : fmtMoney(fatMedia),
      fatMesAnt == null ? "—" : fmtMoney(fatMesAnt),
      fatVsMesAnt,
      fatYoyMedia == null ? "—" : fmtMoney(fatYoyMedia),
      fatVsYoy
    );

    if(fatTotal != null){ sumFatAll += fatTotal; anyFat = true; }

    // Custos
    const custoTotal = sumPeriod(rows, headerMonths, "CUSTO", baseKeys);
    const custoMedia = (custoTotal == null) ? null : custoTotal / nMonths;
    const custoMesAnt = custoMonth(rows, headerMonths, pmKey);
    const custoVsMesAnt = pctDelta(custoMedia, custoMesAnt);

    const custoYoyTotal = sumPeriod(rows, headerMonths, "CUSTO", yoyKeys);
    const custoYoyMedia = (custoYoyTotal == null) ? null : custoYoyTotal / nMonths;
    const custoVsYoy = pctDelta(custoMedia, custoYoyMedia);

    htmlCusto += unitLineFinanceHTML(
      unitKey,
      prettyUnit(unitKey),
      custoMedia == null ? "—" : fmtMoney(custoMedia),
      custoMesAnt == null ? "—" : fmtMoney(custoMesAnt),
      custoVsMesAnt,
      custoYoyMedia == null ? "—" : fmtMoney(custoYoyMedia),
      custoVsYoy
    );

    if(custoTotal != null){ sumCustoAll += custoTotal; anyCusto = true; }

    // Resultado
    const resTotal = resultadoPeriod(rows, headerMonths, baseKeys);
    const resMedia = (resTotal == null) ? null : resTotal / nMonths;
    const resMesAnt = resultadoMonth(rows, headerMonths, pmKey);
    const resVsMesAnt = pctDelta(resMedia, resMesAnt);

    const resYoyTotal = resultadoPeriod(rows, headerMonths, yoyKeys);
    const resYoyMedia = (resYoyTotal == null) ? null : resYoyTotal / nMonths;
    const resVsYoy = pctDelta(resMedia, resYoyMedia);

    htmlRes += unitLineFinanceHTML(
      unitKey,
      prettyUnit(unitKey),
      resMedia == null ? "—" : fmtMoney(resMedia),
      resMesAnt == null ? "—" : fmtMoney(resMesAnt),
      resVsMesAnt,
      resYoyMedia == null ? "—" : fmtMoney(resYoyMedia),
      resVsYoy
    );

    if(resTotal != null){ sumResAll += resTotal; anyRes = true; }
  }

  CARD.faturamento.innerHTML = htmlFat || "<div class='muted'>—</div>";
  CARD.custo.innerHTML = htmlCusto || "<div class='muted'>—</div>";
  CARD.resultado.innerHTML = htmlRes || "<div class='muted'>—</div>";

  setTotalPill(TOTAL.faturamento, anyFat ? fmtMoney(sumFatAll) : "—");
  setTotalPill(TOTAL.custo, anyCusto ? fmtMoney(sumCustoAll) : "—");
  setTotalPill(TOTAL.resultado, anyRes ? fmtMoney(sumResAll) : "—");

  // ================== LINHAS ABAIXO (SEM MUDAR PADRÃO) ==================
  renderSimpleCards(baseKeys, prevKeys, yoyKeys);

  // ================== TABELA ==================
  renderTable(baseUnit, baseKeys);

  setStatus("Aplicado.");
}

// Mantém a estrutura antiga: valor + % vs mês anterior + % vs ano anterior
function renderSimpleCards(baseKeys, prevKeys, yoyKeys){
  let htmlAtivos="", htmlMat="", htmlCad="", htmlConv="", htmlRec="";
  let htmlInad="", htmlEvasao="", htmlMargem="";

  let totAtivos=0, anyAt=false;
  let totMat=0, anyMat=false;
  let totCad=0, anyCad=false;
  let totRec=0, anyRec=false;

  // para percentuais totais (conversão/margem) eu faço média simples por unidade (mantém coerência de “como estava”)
  let sumConv=0, nConv=0;
  let sumInad=0, nInad=0;
  let sumEvas=0, nEvas=0;
  let sumMarg=0, nMarg=0;

  for(const unitKey of selectedUnits){
    const { rows, headerMonths } = cache.get(unitKey);

    // Ativos (média mensal no período)
    const aNow = ativosPeriodAvg(rows, headerMonths, baseKeys);
    const aPrev = ativosPeriodAvg(rows, headerMonths, prevKeys);
    const aYoy = ativosPeriodAvg(rows, headerMonths, yoyKeys);

    htmlAtivos += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      aNow == null ? "—" : fmtInt(aNow),
      pctDelta(aNow, aPrev),
      pctDelta(aNow, aYoy)
    );

    if(aNow != null){ totAtivos += aNow; anyAt=true; }

    // Matrículas (total no período)
    const mNow = sumPeriod(rows, headerMonths, "MAT", baseKeys);
    const mPrev = sumPeriod(rows, headerMonths, "MAT", prevKeys);
    const mYoy = sumPeriod(rows, headerMonths, "MAT", yoyKeys);

    htmlMat += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      mNow == null ? "—" : fmtInt(mNow),
      pctDelta(mNow, mPrev),
      pctDelta(mNow, mYoy)
    );
    if(mNow != null){ totMat += mNow; anyMat=true; }

    // Cadastros (total)
    const cNow = sumPeriod(rows, headerMonths, "CAD", baseKeys);
    const cPrev = sumPeriod(rows, headerMonths, "CAD", prevKeys);
    const cYoy = sumPeriod(rows, headerMonths, "CAD", yoyKeys);

    htmlCad += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      cNow == null ? "—" : fmtInt(cNow),
      pctDelta(cNow, cPrev),
      pctDelta(cNow, cYoy)
    );
    if(cNow != null){ totCad += cNow; anyCad=true; }

    // Conversão (mat/cad)
    const convNow = conversaoPeriod(rows, headerMonths, baseKeys);
    const convPrev = conversaoPeriod(rows, headerMonths, prevKeys);
    const convYoy = conversaoPeriod(rows, headerMonths, yoyKeys);

    htmlConv += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      convNow == null ? "—" : fmtPct(convNow),
      pctDelta(convNow, convPrev),
      pctDelta(convNow, convYoy)
    );
    if(convNow != null){ sumConv += convNow; nConv++; }

    // Recebidas (total)
    const rNow = recebidasPeriod(rows, headerMonths, baseKeys);
    const rPrev = recebidasPeriod(rows, headerMonths, prevKeys);
    const rYoy = recebidasPeriod(rows, headerMonths, yoyKeys);

    htmlRec += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      rNow == null ? "—" : fmtMoney(rNow),
      pctDelta(rNow, rPrev),
      pctDelta(rNow, rYoy)
    );
    if(rNow != null){ totRec += rNow; anyRec=true; }

    // Inadimplência (média % no período) => lower is better
    const inadNow = avgPctPeriod(rows, headerMonths, "INAD", baseKeys);
    const inadPrev = avgPctPeriod(rows, headerMonths, "INAD", prevKeys);
    const inadYoy = avgPctPeriod(rows, headerMonths, "INAD", yoyKeys);

    htmlInad += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      inadNow == null ? "—" : fmtPct(inadNow),
      pctDelta(inadNow, inadPrev),
      pctDelta(inadNow, inadYoy),
      true
    );
    if(inadNow != null){ sumInad += inadNow; nInad++; }

    // Evasão (média % no período) => lower is better
    const evNow = avgPctPeriod(rows, headerMonths, "EVASAO", baseKeys);
    const evPrev = avgPctPeriod(rows, headerMonths, "EVASAO", prevKeys);
    const evYoy = avgPctPeriod(rows, headerMonths, "EVASAO", yoyKeys);

    htmlEvasao += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      evNow == null ? "—" : fmtPct(evNow),
      pctDelta(evNow, evPrev),
      pctDelta(evNow, evYoy),
      true
    );
    if(evNow != null){ sumEvas += evNow; nEvas++; }

    // Margem (média % no período)
    const mgNow = avgPctPeriod(rows, headerMonths, "MARGEM", baseKeys);
    const mgPrev = avgPctPeriod(rows, headerMonths, "MARGEM", prevKeys);
    const mgYoy = avgPctPeriod(rows, headerMonths, "MARGEM", yoyKeys);

    htmlMargem += unitLineSimpleHTML(
      unitKey,
      prettyUnit(unitKey),
      mgNow == null ? "—" : fmtPct(mgNow),
      pctDelta(mgNow, mgPrev),
      pctDelta(mgNow, mgYoy)
    );
    if(mgNow != null){ sumMarg += mgNow; nMarg++; }
  }

  // Render nas caixas
  CARD.ativos.innerHTML = htmlAtivos || "<div class='muted'>—</div>";
  CARD.matriculas.innerHTML = htmlMat || "<div class='muted'>—</div>";
  CARD.cadastros.innerHTML = htmlCad || "<div class='muted'>—</div>";
  CARD.conversao.innerHTML = htmlConv || "<div class='muted'>—</div>";
  CARD.recebidas.innerHTML = htmlRec || "<div class='muted'>—</div>";

  CARD.inad.innerHTML = htmlInad || "<div class='muted'>—</div>";
  CARD.evasao.innerHTML = htmlEvasao || "<div class='muted'>—</div>";
  CARD.margem.innerHTML = htmlMargem || "<div class='muted'>—</div>";

  // Totais (no topo do card)
  setTotalPill(TOTAL.ativos, anyAt ? fmtInt(totAtivos) : "—");
  setTotalPill(TOTAL.matriculas, anyMat ? fmtInt(totMat) : "—");
  setTotalPill(TOTAL.cadastros, anyCad ? fmtInt(totCad) : "—");
  setTotalPill(TOTAL.recebidas, anyRec ? fmtMoney(totRec) : "—");

  // Percentuais totais: média das unidades (mantém a ideia antiga)
  setTotalPill(TOTAL.conversao, nConv ? fmtPct(sumConv/nConv) : "—");
  setTotalPill(TOTAL.inad, nInad ? fmtPct(sumInad/nInad) : "—");
  setTotalPill(TOTAL.evasao, nEvas ? fmtPct(sumEvas/nEvas) : "—");
  setTotalPill(TOTAL.margem, nMarg ? fmtPct(sumMarg/nMarg) : "—");
}

// ================= TABLE =================
function renderTable(unitKey, baseKeys){
  if(!TABLE) return;
  const { rows, colLabels, headerMonths } = cache.get(unitKey);
  const mode = elMode?.value || "interval";

  const indices = mode === "full"
    ? headerMonths.map(h => h.idx)
    : buildPeriodIndices(headerMonths, baseKeys);

  const headCols = indices.map(i => colLabels[i] || "");
  if(tblInfo) tblInfo.textContent = `Tabela: ${prettyUnit(unitKey)} • Colunas: ${mode === "full" ? "todas" : "período"}`;

  const thead = `
    <thead>
      <tr>
        <th>Indicador (Coluna A)</th>
        ${headCols.map(h => `<th>${h}</th>`).join("")}
      </tr>
    </thead>
  `;

  const tbodyRows = rows
    .filter(r => String(r?.[0] ?? "").trim().length)
    .map(r => {
      const label = String(r[0] ?? "");
      return `
        <tr>
          <td>${label}</td>
          ${indices.map(i => `<td>${r[i] ?? ""}</td>`).join("")}
        </tr>
      `;
    }).join("");

  TABLE.innerHTML = thead + `<tbody>${tbodyRows}</tbody>`;
}
