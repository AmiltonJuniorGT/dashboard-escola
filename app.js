// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";
const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};
const DEFAULT_SELECTED = ["CAMACARI"];

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

const tblInfo = $("tblInfo");
const WRAP = $("tableWrap");
const TABLE = $("kpiTable");

const CARD = {
  faturamento: $("card_faturamento"),
  custo: $("card_custo"),
  resultado: $("card_resultado"),
};
const TOTAL = {
  faturamento: $("total_faturamento"),
  custo: $("total_custo"),
  resultado: $("total_resultado"),
};

// ================= STATE =================
let selectedUnits = [...DEFAULT_SELECTED];
const cache = new Map(); // unitKey -> { rows, colLabels, headerMonths }

// ================= NORMALIZE / MATCH =================
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
};

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
function addMonths(d, delta){ return new Date(d.getFullYear(), d.getMonth()+delta, 1); }
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
function monthsN(keys){ return keys?.length ? keys.length : 1; }
function prevMonthKey(startISO){
  const d = new Date(startISO+"T00:00:00");
  const startMonth = new Date(d.getFullYear(), d.getMonth(), 1);
  return monthKey(addMonths(startMonth, -1));
}
function shiftPeriodKeys(keys, deltaMonths){
  if(!keys.length) return [];
  const first = keys[0] + "-01";
  const last = keys[keys.length-1] + "-01";
  const a = addMonths(new Date(first+"T00:00:00"), deltaMonths);
  const b = addMonths(new Date(last+"T00:00:00"), deltaMonths);
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
function fmtPct(n){
  if(n == null) return "—";
  return `${n.toFixed(2).replace(".",",")}%`;
}
function pctDelta(base, comp){
  if(base == null || comp == null || comp === 0) return null;
  return ((base - comp)/comp)*100;
}
function deltaCls(pct){
  if(pct == null) return "";
  if(pct > 0) return "good";
  if(pct < 0) return "bad";
  return "";
}
function setStatus(msg){ elStatus.textContent = msg; }

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
  const out = [];
  for(const r of table.rows){
    const row = [];
    for(const c of r.c){
      row.push(c ? (c.v ?? c.f ?? null) : null);
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
    const key = monthLabelToKey(colLabels[c]);
    if(key) headerMonths.push({ idx:c, label:colLabels[c], key });
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
function findRow(rows, keys){
  for(const row of rows){
    const label = row?.[0];
    if(labelHas(label, keys)) return row;
  }
  return null;
}
function getVal(rows, kpiKey, colIdx){
  if(colIdx == null) return null;
  const keys = KPI_MATCH[kpiKey] || [];
  const row = findRow(rows, keys);
  return cleanNumber(row?.[colIdx]);
}

// ================= FINANCE AGG =================
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
function faturamentoPeriod(rows, headerMonths, periodKeys){
  const t = sumPeriod(rows, headerMonths, "FAT_T", periodKeys);
  const p = sumPeriod(rows, headerMonths, "FAT_P", periodKeys);
  if(t!=null || p!=null) return (t||0) + (p||0);
  return sumPeriod(rows, headerMonths, "FAT", periodKeys);
}
function faturamentoMonth(rows, headerMonths, keyYYYYMM){
  const col = findMonthColByKey(headerMonths, keyYYYYMM);
  if(col == null) return null;
  const t = getVal(rows, "FAT_T", col);
  const p = getVal(rows, "FAT_P", col);
  if(t!=null || p!=null) return (t||0) + (p||0);
  return getVal(rows, "FAT", col);
}
function custoPeriod(rows, headerMonths, periodKeys){
  return sumPeriod(rows, headerMonths, "CUSTO", periodKeys);
}
function custoMonth(rows, headerMonths, keyYYYYMM){
  const col = findMonthColByKey(headerMonths, keyYYYYMM);
  if(col == null) return null;
  return getVal(rows, "CUSTO", col);
}
function resultadoPeriod(rows, headerMonths, periodKeys){
  const lucro = sumPeriod(rows, headerMonths, "LUCRO", periodKeys);
  if(lucro != null) return lucro;
  return sumPeriod(rows, headerMonths, "RESULTADO_LIQ", periodKeys);
}
function resultadoMonth(rows, headerMonths, keyYYYYMM){
  const col = findMonthColByKey(headerMonths, keyYYYYMM);
  if(col == null) return null;
  const lucro = getVal(rows, "LUCRO", col);
  if(lucro != null) return lucro;
  return getVal(rows, "RESULTADO_LIQ", col);
}

// ================= UI =================
function prettyUnit(unitKey){ return TABS[unitKey].replace("Números ",""); }

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
        renderUnitMenu();
      }
      refreshUnitSummary();
    });
  });

  unitMenu.querySelector("#selAll").addEventListener("click", (e) => {
    e.stopPropagation();
    selectedUnits = Object.keys(TABS);
    renderUnitMenu(); refreshUnitSummary();
  });
  unitMenu.querySelector("#selBase").addEventListener("click", (e) => {
    e.stopPropagation();
    selectedUnits = ["CAMACARI"];
    renderUnitMenu(); refreshUnitSummary();
  });
}

function refreshUnitSummary(){
  unitSummary.textContent = selectedUnits.length
    ? selectedUnits.map(prettyUnit).join(" • ")
    : "Selecione ao menos 1 unidade";
  unitBtnLabel.textContent = selectedUnits.length ? `${selectedUnits.length} unidade(s)` : "Selecionar unidades";
}

document.addEventListener("DOMContentLoaded", () => {
  const now = new Date();
  elStart.value = toISODate(new Date(now.getFullYear(),0,1));
  elEnd.value = toISODate(new Date(now.getFullYear(),11,31));

  renderUnitMenu();
  refreshUnitSummary();

  unitBtn.addEventListener("click", (e) => {
    e.preventDefault(); e.stopPropagation();
    unitMenu.classList.toggle("open");
  });
  document.addEventListener("click", (e) => {
    if(!e.target.closest("#multiSelect")){
      unitMenu.classList.remove("open");
    }
  });

  btnLoad.addEventListener("click", () => loadSelectedUnits(true));
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "Carregando dados…";
  loadSelectedUnits(false);
});

// ================= LOAD =================
async function loadUnit(unitKey, force=false){
  if(!force && cache.has(unitKey)) return;

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

async function loadSelectedUnits(force=false){
  try{
    setStatus("Carregando unidades…");
    hintBox.textContent = "Carregando dados das unidades selecionadas…";
    await Promise.all(selectedUnits.map(u => loadUnit(u, force)));
    setStatus("Dados carregados.");
    apply();
  }catch(err){
    console.error(err);
    setStatus("Erro.");
    hintBox.textContent = "Erro ao carregar: " + (err?.message || err);
  }
}

// ================= RENDER HELPERS =================
function unitLineHTML(unitKey, title, mainValue, monthPrevValue, monthPrevPct, yoyValue, yoyPct){
  const cls = `unitLine unit-${unitKey}`;
  const pct1cls = deltaCls(monthPrevPct);
  const pct2cls = deltaCls(yoyPct);

  return `
    <div class="${cls}">
      <div class="unitTop">
        <div class="unitName">${title}</div>
        <div class="unitValue">${mainValue}</div>
      </div>

      <div class="unitBottom">
        <div class="rowLine">
          <span>Mês ant.: <b>${monthPrevValue}</b></span>
          <span class="pct ${pct1cls}">${fmtPct(monthPrevPct)}</span>
        </div>

        <div class="rowLine">
          <span>Ano ant.: <b>${yoyValue}</b></span>
          <span class="pct ${pct2cls}">${fmtPct(yoyPct)}</span>
        </div>
      </div>
    </div>
  `;
}

function setTotalPill(el, value){
  el.textContent = `Total: ${value}`;
}

// ================= APPLY (PRINT LOGIC) =================
function apply(){
  if(!selectedUnits.length) return;

  const startISO = elStart.value;
  const endISO = elEnd.value;

  const baseKeys = monthRangeKeys(startISO, endISO);
  const n = monthsN(baseKeys);
  const pmKey = prevMonthKey(startISO);
  const yoyKeys = shiftPeriodKeys(baseKeys, -12);

  metaUnit.textContent = selectedUnits.map(prettyUnit).join(", ");
  metaPeriod.textContent = `${brDate(startISO)} → ${brDate(endISO)}`;
  metaBaseMonth.textContent = baseKeys.length ? baseKeys[baseKeys.length-1] : "—";

  // Validação meses detectados
  const baseUnit = selectedUnits[0];
  const baseData = cache.get(baseUnit);
  if(!baseData?.headerMonths?.length){
    hintBox.textContent = "⚠️ Não detectei meses nos cabeçalhos. Confirme se está como 'jan./25', 'fev./25' etc.";
  } else {
    hintBox.textContent = "";
  }

  // ---------- FATURAMENTO (Topo = total do período / linhas = média mensal + comps)
  let fatTotalAll = 0, fatAny = false;
  let fatHTML = "";

  // ---------- CUSTO
  let custoTotalAll = 0, custoAny = false;
  let custoHTML = "";

  // ---------- RESULTADO
  let resTotalAll = 0, resAny = false;
  let resHTML = "";

  for(const unitKey of selectedUnits){
    const { rows, headerMonths } = cache.get(unitKey);

    // FAT: total período
    const fatTotal = faturamentoPeriod(rows, headerMonths, baseKeys);
    if(fatTotal != null){ fatTotalAll += fatTotal; fatAny = true; }
    const fatMedia = (fatTotal == null) ? null : (fatTotal / n);

    const fatMesAnt = faturamentoMonth(rows, headerMonths, pmKey);
    const fatVsMesAnt = pctDelta(fatMedia, fatMesAnt);

    const fatYoyTotal = faturamentoPeriod(rows, headerMonths, yoyKeys);
    const fatYoyMedia = (fatYoyTotal == null) ? null : (fatYoyTotal / n);
    const fatVsYoy = pctDelta(fatMedia, fatYoyMedia);

    fatHTML += unitLineHTML(
      unitKey,
      prettyUnit(unitKey),
      fatMedia == null ? "—" : fmtMoney(fatMedia),
      fatMesAnt == null ? "—" : fmtMoney(fatMesAnt),
      fatVsMesAnt,
      fatYoyMedia == null ? "—" : fmtMoney(fatYoyMedia),
      fatVsYoy
    );

    // CUSTO
    const cTotal = custoPeriod(rows, headerMonths, baseKeys);
    if(cTotal != null){ custoTotalAll += cTotal; custoAny = true; }
    const cMedia = (cTotal == null) ? null : (cTotal / n);

    const cMesAnt = custoMonth(rows, headerMonths, pmKey);
    const cVsMesAnt = pctDelta(cMedia, cMesAnt);

    const cYoyTotal = custoPeriod(rows, headerMonths, yoyKeys);
    const cYoyMedia = (cYoyTotal == null) ? null : (cYoyTotal / n);
    const cVsYoy = pctDelta(cMedia, cYoyMedia);

    custoHTML += unitLineHTML(
      unitKey,
      prettyUnit(unitKey),
      cMedia == null ? "—" : fmtMoney(cMedia),
      cMesAnt == null ? "—" : fmtMoney(cMesAnt),
      cVsMesAnt,
      cYoyMedia == null ? "—" : fmtMoney(cYoyMedia),
      cVsYoy
    );

    // RESULTADO
    const rTotal = resultadoPeriod(rows, headerMonths, baseKeys);
    if(rTotal != null){ resTotalAll += rTotal; resAny = true; }
    const rMedia = (rTotal == null) ? null : (rTotal / n);

    const rMesAnt = resultadoMonth(rows, headerMonths, pmKey);
    const rVsMesAnt = pctDelta(rMedia, rMesAnt);

    const rYoyTotal = resultadoPeriod(rows, headerMonths, yoyKeys);
    const rYoyMedia = (rYoyTotal == null) ? null : (rYoyTotal / n);
    const rVsYoy = pctDelta(rMedia, rYoyMedia);

    resHTML += unitLineHTML(
      unitKey,
      prettyUnit(unitKey),
      rMedia == null ? "—" : fmtMoney(rMedia),
      rMesAnt == null ? "—" : fmtMoney(rMesAnt),
      rVsMesAnt,
      rYoyMedia == null ? "—" : fmtMoney(rYoyMedia),
      rVsYoy
    );
  }

  CARD.faturamento.innerHTML = fatHTML || "<div class='muted'>—</div>";
  CARD.custo.innerHTML = custoHTML || "<div class='muted'>—</div>";
  CARD.resultado.innerHTML = resHTML || "<div class='muted'>—</div>";

  setTotalPill(TOTAL.faturamento, fatAny ? fmtMoney(fatTotalAll) : "—");
  setTotalPill(TOTAL.custo, custoAny ? fmtMoney(custoTotalAll) : "—");
  setTotalPill(TOTAL.resultado, resAny ? fmtMoney(resTotalAll) : "—");

  // tabela da unidade base (primeira selecionada)
  renderTable(baseUnit, baseKeys);
}

// ================= TABLE =================
function renderTable(unitKey, baseKeys){
  const { rows, colLabels, headerMonths } = cache.get(unitKey);
  const mode = elMode.value;

  const indices = mode === "full"
    ? headerMonths.map(h => h.idx)
    : buildPeriodIndices(headerMonths, baseKeys);

  const headCols = indices.map(i => colLabels[i] || "");
  tblInfo.textContent = `Tabela: ${prettyUnit(unitKey)} • Colunas: ${mode === "full" ? "todas" : "período"}`;

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
