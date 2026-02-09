// ================= CONFIG =================
const SHEET_ID = "1sSk34SsgQ_2aAgr2Uqra4nVs9C0TVh-h";
const UNIDADES_SHEET_ID = "164MCqlJzWgUDeFzm8TxLhuxvX_-VeCPGQwrBvJIJWyQ";

const TABS = {
  CONFIG: "Config",
  CAJ: "Números Cajazeiras",
  CAM: "Números Camaçari",
  SAO: "Números São Cristóvão"
};

// ================= STATE ==================
const state = {
  periodoInicio: null,
  periodoFim: null,
  marca: "C",
  unidade: "TODAS"
};

// ================= UTIL ===================
const $ = id => document.getElementById(id);

function iso(d){
  return d.toISOString().slice(0,10);
}
function firstDay(y,m){
  return `${y}-${String(m).padStart(2,"0")}-01`;
}
function lastDay(y,m){
  return `${y}-${String(m).padStart(2,"0")}-31`;
}

// ================= SHEETS =================
async function fetchSheet(sheetId, tab){
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?sheet=${encodeURIComponent(tab)}`;
  const r = await fetch(url);
  const t = await r.text();
  const j = JSON.parse(t.substring(t.indexOf("{"), t.lastIndexOf("}")+1));
  return j.table.rows.map(r => ({
    k: r.c[0]?.v,
    v: r.c[1]?.v
  }));
}

function rowsToObj(rows){
  const o = {};
  rows.forEach(r => o[r.k] = r.v);
  return o;
}

// ================= CONFIG LOAD =============
async function loadConfig(){
  try {
    const cfg = rowsToObj(await fetchSheet(SHEET_ID, TABS.CONFIG));
    const mm = String(cfg["Mês (mm)"]).padStart(2,"0");
    const aa = "20" + cfg["Ano (aa)"];

    state.periodoInicio = firstDay(aa, mm);
    state.periodoFim = lastDay(aa, mm);

  } catch {
    const d = new Date();
    state.periodoInicio = firstDay(d.getFullYear(), d.getMonth()+1);
    state.periodoFim = iso(d);
  }
}

// ================= UI =====================
function syncUI(){
  $("f_ini").value = state.periodoInicio;
  $("f_fim").value = state.periodoFim;
  $("f_marca").value = state.marca;
  $("f_unidade").value = state.unidade;

  $("status").innerText =
    `Período: ${state.periodoInicio} → ${state.periodoFim} | Marca: ${state.marca} | Unidade: ${state.unidade}`;
}

// ================= FECHAMENTO =============
async function renderFechamento(){
  const map = {
    Cajazeiras: TABS.CAJ,
    Camaçari: TABS.CAM,
    "São Cristóvão": TABS.SAO
  };

  const prefix = state.marca === "T" ? "gt_" :
                 state.marca === "P" ? "gp_" : "ge_";

  const unidades = state.unidade === "TODAS"
    ? Object.keys(map)
    : [state.unidade];

  let html = "";

  for (const u of unidades){
    const data = rowsToObj(await fetchSheet(UNIDADES_SHEET_ID, map[u]));
    html += `
      <div class="card">
        <h3>${u}</h3>
        Faturamento: ${data[prefix+"faturamento"] ?? "—"}<br>
        Matrículas: ${data[prefix+"matriculas"] ?? "—"}<br>
        Ativos: ${data[prefix+"ativos"] ?? "—"}
      </div>
    `;
  }

  $("fech_cards").innerHTML = html;
}

// ================= EVENTS =================
function bind(){
  $("f_apply").onclick = async () => {
    state.periodoInicio = $("f_ini").value;
    state.periodoFim = $("f_fim").value;
    state.marca = $("f_marca").value;
    state.unidade = $("f_unidade").value;
    syncUI();
    await renderFechamento();
  };

  $("f_reset").onclick = async () => {
    await loadConfig();
    syncUI();
    await renderFechamento();
  };
}

// ================= START ==================
window.onload = async () => {
  await loadConfig();
  syncUI();
  bind();
  await renderFechamento();
};
