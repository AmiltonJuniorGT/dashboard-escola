// ================= CONFIG =================
const SHEET_ID = "1sSk34SsgQ_2aAgr2Uqra4nVs9C0TVh-h";
const UNIDADES_SHEET_ID = "164MCqlJzWgUDeFzm8TxLhuxvX_-VeCPGQwrBvJIJWyQ";

const TABS = {
  CONFIG: "Config",
  CAJ: "Números Cajazeiras",
  CAM: "Números Camaçari",
  SAO: "Números São Cristóvão"
};

const state = {
  periodoInicio: "",
  periodoFim: "",
  marca: "T",     // T = Grau Técnico | P = Profissionalizante | C = Consolidado
  unidade: "TODAS"
};

const $ = id => document.getElementById(id);

// ================= UTIL =================
function iso(d){ return d.toISOString().slice(0,10); }
function firstDay(y,m){ return `${y}-${String(m).padStart(2,"0")}-01`; }
function lastDay(y,m){ return `${y}-${String(m).padStart(2,"0")}-31`; }
function norm(v){ return String(v ?? "").trim(); }

// ================= SHEETS =================
async function fetchSheet(sheetId, tab){
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?sheet=${encodeURIComponent(tab)}`;
  const res = await fetch(url);
  const txt = await res.text();
  const json = JSON.parse(txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1));
  return json.table.rows.map(r => ({
    k: norm(r.c[0]?.v),
    v: r.c[1]?.v
  }));
}

function rowsToObj(rows){
  const o = {};
  rows.forEach(r => { if (r.k) o[r.k] = r.v; });
  return o;
}

// ================= CONFIG =================
async function loadConfig(){
  try{
    const cfg = rowsToObj(await fetchSheet(SHEET_ID, TABS.CONFIG));
    const mm = String(cfg["Mês (mm)"] ?? "").padStart(2,"0");
    const aa = "20" + String(cfg["Ano (aa)"] ?? "");
    state.periodoInicio = firstDay(aa, mm);
    state.periodoFim = lastDay(aa, mm);
  } catch {
    const d = new Date();
    state.periodoInicio = firstDay(d.getFullYear(), d.getMonth()+1);
    state.periodoFim = iso(d);
  }
}

// ================= UI =================
function syncUI(){
  $("f_ini").value = state.periodoInicio;
  $("f_fim").value = state.periodoFim;
  $("f_marca").value = state.marca;
  $("f_unidade").value = state.unidade;
  $("status").innerText =
    `Período: ${state.periodoInicio} → ${state.periodoFim} | Marca: ${state.marca} | Unidade: ${state.unidade}`;
}

// ================= FECHAMENTO =================
function prefixMarca(){
  if(state.marca === "T") return "gt_";
  if(state.marca === "P") return "gp_";
  return ""; // consolidado depois
}

async function renderFechamento(){
  const MAP = {
    Cajazeiras: TABS.CAJ,
    Camaçari: TABS.CAM,
    "São Cristóvão": TABS.SAO
  };

  const unidades = state.unidade === "TODAS"
    ? Object.keys(MAP)
    : [state.unidade];

  let html = "";
  const pref = prefixMarca();

  for(const u of unidades){
    const tab = MAP[u];
    const data = rowsToObj(await fetchSheet(UNIDADES_SHEET_ID, tab));

    const fat = data[pref + "faturamento"];
    const mat = data[pref + "matriculas"];
    const ati = data[pref + "ativos"];

    html += `
      <div class="card">
        <h3>${u}</h3>
        Faturamento: ${fat ?? "—"}<br>
        Matrículas: ${mat ?? "—"}<br>
        Ativos: ${ati ?? "—"}
      </div>`;
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

// ================= START =================
window.onload = async () => {
  await loadConfig();
  syncUI();
  bind();
  await renderFechamento();
};
