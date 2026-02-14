const DATA_SHEET_ID = "12PUef42a3r3_fxJS8SulkbeXoGP3COtN";
const TABS_UNIDADES = {
  Cajazeiras: "Números Cajazeiras",
  //Camaçari: "Números Camaçari",
 //"São Cristóvão": "Números Sao Cristovao"
};

const state = {
  unidade: "TODAS",
  monthKey: null
};

const $ = (id) => document.getElementById(id);

function setStatus(msg, isError = false) {
  const el = $("status");
  el.textContent = msg;
  el.style.color = isError ? "red" : "#1f5130";
}

function initUnidades() {
  const sel = $("unidade");
  sel.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "TODAS";
  optAll.textContent = "Todas";
  sel.appendChild(optAll);

  Object.keys(TABS_UNIDADES).forEach(u => {
    const opt = document.createElement("option");
    opt.value = u;
    opt.textContent = u;
    sel.appendChild(opt);
  });
}

function getCurrentMonthKey() {
  const d = new Date();
  return { y: d.getFullYear(), m: d.getMonth() + 1 };
}

function monthLabelToKey(label) {
  const meses = {
    janeiro:1, fevereiro:2, marco:3, março:3, abril:4, maio:5, junho:6,
    julho:7, agosto:8, setembro:9, outubro:10, novembro:11, dezembro:12
  };
  const clean = label.toLowerCase().replace(".", "");
  const parts = clean.split(/\s+/);
  if (parts.length < 2) return null;
  const mesTxt = parts[0];
  const anoTxt = parts[1];
  const m = meses[mesTxt];
  const y = 2000 + parseInt(anoTxt);
  if (!m || isNaN(y)) return null;
  return { y, m };
}

async function fetchSheet(tabName) {
  const url = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?sheet=${encodeURIComponent(tabName)}&tq=select *`;
  const res = await fetch(url);
  const txt = await res.text();
  const json = JSON.parse(txt.substring(47).slice(0, -2));
  return json.table.rows.map(r => r.c.map(c => c ? c.v : ""));
}

function buildFechamento(data, monthKey) {
  const container = $("fechamento");
  container.innerHTML = "";

  const header = data[0];
  const monthColIndex = header.findIndex(h => {
    if (!h) return false;
    const key = monthLabelToKey(String(h));
    return key && key.y === monthKey.y && key.m === monthKey.m;
  });

  if (monthColIndex === -1) {
    container.innerHTML = "<div class='erro'>Não encontrei colunas de mês no cabeçalho.</div>";
    return;
  }

  const table = document.createElement("table");
  table.className = "tabela";

  data.slice(1).forEach(row => {
    if (!row[0]) return;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${row[0]}</td><td>${row[monthColIndex] ?? ""}</td>`;
    table.appendChild(tr);
  });

  container.appendChild(table);
}

async function loadFechamento() {
  try {
    setStatus("Carregando dados...");
    const monthKey = state.monthKey || getCurrentMonthKey();
    state.monthKey = monthKey;

    const tabs = state.unidade === "TODAS"
      ? Object.values(TABS_UNIDADES)
      : [TABS_UNIDADES[state.unidade]];

    for (const tab of tabs) {
      const data = await fetchSheet(tab);
      buildFechamento(data, monthKey);
    }

    setStatus("Dados carregados com sucesso.");
  } catch (e) {
    console.error(e);
    setStatus("Erro ao atualizar ❌", true);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  initUnidades();
  state.monthKey = getCurrentMonthKey();
  loadFechamento();

  $("btnAplicar").addEventListener("click", () => {
    state.unidade = $("unidade").value;
    loadFechamento();
  });
});
