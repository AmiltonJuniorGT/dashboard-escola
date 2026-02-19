// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS = {
  CAMACARI: "Números Camaçari",
  CAJAZEIRAS: "Números Cajazeiras",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

const DEFAULT_SELECTED = ["CAMACARI"];

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

  // dropdown robusto
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

  btnLoad.addEventListener("click", () => loadSelectedUnits(true));
  btnApply.addEventListener("click", () => apply());
  btnExpand.addEventListener("click", () => WRAP.classList.toggle("expanded"));

  setStatus("Pronto. Carregando…");
  hintBox.textContent = "Carregando dados…";
  loadSelectedUnits(false);
}

// ================= UI HELPERS =================
function prettyUnit(unitKey){ return TABS[unitKey].replace("Números ",""); }

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
      <button class="linkBtn" id="selBase" type="button">Só Camaçari</butt
