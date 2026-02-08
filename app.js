// =====================================================
// DASHBOARD – APP.JS (VERSÃO ESTÁVEL)
// =====================================================

// ================= CONFIG =============================
const SHEET_ID = "1sSk34SsgQ_2aAgr2Uqra4nVs9C0TVh-h"; 
const UNIDADES_SHEET_ID = "164MCqlJzWgUDeFzm8TxLhuxvX_-VeCPGQwrBvJIJWyQ";

const TABS = {
  DADOS: "DADOS_API",
  CURSOS: "CURSOS_ABC",
  TURNOS: "TURMAS_DECISAO",
  CONFIG: "CONFIG_API",
  CAJAZEIRAS: "Números Cajazeiras",
  CAMACARI: "Números Camaçari",
  SAO_CRISTOVAO: "Números São Cristóvão",
};

// ================= ESTADO GLOBAL ======================
const state = {
  periodoInicio: null,
  periodoFim: null,
  marca: "C",          // C | T | P
  unidade: "TODAS",    // TODAS | Cajazeiras | Camaçari | São Cristóvão
};

// ================= UTIL ===============================
const $ = id => document.getElementById(id);

function pad(n) { return String(n).padStart(2, "0"); }
function iso(d) { return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
function firstDay(isoDate) {
  const d = new Date(isoDate);
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-01`;
}

function fmt(v, money = false) {
  if (v === "" || v === null || v === undefined) return "—";
  const n = Number(v);
  if (!isNaN(n)) {
    return money ? n.toLocaleString("pt-BR",{style:"currency",currency:"BRL"}) : n.toLocaleString("pt-BR");
  }
  return v;
}

// ================= SHEETS =============================
function sheetIdByTab(tab) {
  return (
    tab === TABS.CAJAZEIRAS ||
    tab === TABS.CAMACARI ||
    tab === TABS.SAO_CRISTOVAO
  ) ? UNIDADES_SHEET_ID : SHEET_ID;
}

async function fetchSheet(tab) {
  const url =
    `https://docs.google.com/spreadsheets/d/${sheetIdByTab(tab)}/gviz/tq?` +
    `sheet=${encodeURIComponent(tab)}&tq=${encodeURIComponent("select A,B")}`;

  const res = await fetch(url);
  if (!res.ok) throw new Error(`Erro ao carregar ${tab}`);

  const text = await res.text();
  const json = JSON.parse(text.substring(text.indexOf("{"), text.lastIndexOf("}") + 1));
  return json.table.rows.map(r => ({
    key: r.c[0]?.v,
    value: r.c[1]?.v
  }));
}

function rowsToObj(rows) {
  const o = {};
  rows.forEach(r => { if (r.key) o[r.key] = r.value; });
  return o;
}

// ================= CONFIG PADRÃO ======================
async function loadDefaultConfig() {
  try {
    const cfg = rowsToObj(await fetchSheet(TABS.CONFIG));
    state.periodoFim = cfg.periodo_fim_padrao;
    state.periodoInicio = cfg.periodo_inicio_padrao;
    state.marca = cfg.marca_padrao || "C";
    state.unidade = cfg.unidade_padrao || "TODAS";
  } catch {
    const hoje = iso(new Date());
    state.periodoFim = hoje;
    state.periodoInicio = firstDay(hoje);
  }
}

// ================= UI ================================
function syncFilters() {
  $("f_ini").value = state.periodoInicio;
  $("f_fim").value = state.periodoFim;
  $("f_marca").value = state.marca;
  $("f_unidade").value = state.unidade;

  $("status").textContent =
    `Período: ${state.periodoInicio} → ${state.periodoFim} | Marca: ${state.marca} | Unidade: ${state.unidade}`;
}

// ================= RENDER =============================
async function renderFechamento() {
  const unidades = {
    Cajazeiras: TABS.CAJAZEIRAS,
    Camaçari: TABS.CAMACARI,
    "São Cristóvão": TABS.SAO_CRISTOVAO
  };

  const ativos = state.unidade === "TODAS"
    ? Object.keys(unidades)
    : [state.unidade];

  let html = "";

  for (const u of ativos) {
    const data = rowsToObj(await fetchSheet(unidades[u]));
    html += `
      <div class="card">
        <h3>${u}</h3>
        const prefix = state.marca === "T" ? "gt_" : (state.marca === "P" ? "gp_" : "ge_");

        const fatKey = prefix + "faturamento";
        const matKey = prefix + "matriculas";
        const atiKey = prefix + "ativos";
        
        <p>Faturamento: <strong>${fmt(data[fatKey], true)}</strong></p>
        <p>Matrículas: <strong>${fmt(data[matKey])}</strong></p>
        <p>Ativos: <strong>${fmt(data[atiKey])}</strong></p>

      </div>
    `;
  }

  $("fech_cards").innerHTML = html;
}

// ================= EVENTOS ===========================
function wireUI() {
  $("f_apply").onclick = async () => {
    state.periodoInicio = $("f_ini").value;
    state.periodoFim = $("f_fim").value;
    state.marca = $("f_marca").value;
    state.unidade = $("f_unidade").value;
    await renderFechamento();
    syncFilters();
  };

  $("f_reset").onclick = async () => {
    await loadDefaultConfig();
    syncFilters();
    await renderFechamento();
  };
}

// ================= START =============================
window.onload = async () => {
  await loadDefaultConfig();
  syncFilters();
  wireUI();
  await renderFechamento();
};
