// ====== CONFIG (troque só isso) ======
const SHEET_ID = "1sSk34SsgQ_2aAgr2Uqra4nVs9C0TVh-h"; // ex: 1ho2rVMQW-E4dnygqrhUPUHSnItYqhnq0

const TABS = {
  DADOS: "DADOS_API",
  CURSOS: "CURSOS_ABC",
  TURNOS: "TURMAS_DECISAO",

const UNIDADES_SHEET_ID = "1JdheQGLm6AhyOF_a5HL6wOelfpI-GvxV"; // NOVA planilha

const TABS = {
  // novas abas
  CAJAZEIRAS: "Números Cajazeiras",
  CAMACARI: "Números Camaçari",
  SAO_CRISTOVAO: "Números São Cristóvão",
};


// ====== helpers ======
const el = (id) => document.getElementById(id);

function brl(n){
  const v = Number(n);
  if (Number.isNaN(v)) return "—";
  return v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
}
function pct(n){
  const v = Number(n);
  if (Number.isNaN(v)) return "—";
  return (v*100).toFixed(0) + "%";
}

async function carregarUnidades(){
  const unidades = [
    { nome: "Cajazeiras", aba: TABS.CAJAZEIRAS },
    { nome: "Camaçari", aba: TABS.CAMACARI },
    { nome: "São Cristóvão", aba: TABS.SAO_CRISTOVAO },
  ];

  const resultados = [];

  for (const u of unidades){
    const json = await fetchGVizFrom(
      UNIDADES_SHEET_ID,
      u.aba,
      "A:B"
    );
    const { rows } = tableToRows(json);
    const kv = rowsToKV([], rows);
    resultados.push({ nome: u.nome, ...kv });
  }

  return resultados;
}

// ====== fetch gviz ======
async function fetchGViz(sheetName, range="A:Z"){
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}&range=${encodeURIComponent(range)}`;
  const res = await fetch(url, { cache: "no-store" });
  const txt = await res.text();
  const jsonText = txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1);
  return JSON.parse(jsonText);
}

function tableToRows(json){
  const cols = (json.table?.cols || []).map(c => c.label || c.id);
  const rows = (json.table?.rows || []).map(r => (r.c || []).map(c => c?.v ?? null));
  return { cols, rows };
}

function rowsToKV(cols, rows){
  // espera 2 colunas: chave | valor
  const kv = {};
  const kIdx = 0, vIdx = 1;
  rows.forEach(r=>{
    const k = r[kIdx];
    const v = r[vIdx];
    if (k != null) kv[String(k).trim()] = v;
  });
  return kv;
}

async function fetchGVizFrom(sheetId, sheetName, range="A:Z"){
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}&range=${encodeURIComponent(range)}`;
  const res = await fetch(url, { cache: "no-store" });
  const txt = await res.text();
  const jsonText = txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1);
  return JSON.parse(jsonText);
}

function normalizeKey(k){
  return String(k || "")
    .trim()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // remove acentos
}

function prettyLabel(k){
  return String(k)
    .replace(/_/g, " ")
    .replace(/\b\w/g, m => m.toUpperCase());
}

function isPctKey(k){
  const s = normalizeKey(k);
  return s.includes("margem") || s.includes("inadimpl") || s.includes("evas") || s.includes("perda") || s.includes("ocup") || s.includes("percent") || s.endsWith("_pct");
}

function isMoneyKey(k){
  const s = normalizeKey(k);
  return s.includes("fatur") || s.includes("receita") || s.includes("custo") || s.includes("lucro") || s.includes("ticket") || s.includes("valor") || s.includes("mensal");
}

function fmtValue(key, v){
  if (v == null || v === "") return "—";
  const n = Number(v);
  if (Number.isNaN(n)) return String(v);

  if (isPctKey(key)) return (n * 100).toFixed(1).replace(".", ",") + "%";
  if (isMoneyKey(key)) return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });

  // inteiro vs decimal
  if (Math.abs(n) >= 1000 || Number.isInteger(n)) return n.toLocaleString("pt-BR");
  return n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function betterIsHigher(key){
  const s = normalizeKey(key);
  // indicadores onde "menor é melhor"
  if (s.includes("custo") || s.includes("inadimpl") || s.includes("evas") || s.includes("perda") || s.includes("cancel")) return false;
  return true;
}

async function carregarUnidades(){
  const map = [
    { id:"cajazeiras", nome:"Cajazeiras", aba:TABS_UNIDADES.cajazeiras },
    { id:"camacari", nome:"Camaçari", aba:TABS_UNIDADES.camacari },
    { id:"sao_cristovao", nome:"São Cristóvão", aba:TABS_UNIDADES.sao_cristovao },
  ];

  const unidades = [];
  for (const u of map){
    const json = await fetchGVizFrom(UNIDADES_SHEET_ID, u.aba, "A:B");
    const rows = (json.table?.rows || []).map(r => (r.c || []).map(c => c?.v ?? null));

    const kv = {};
    for (const r of rows){
      const k = r[0];
      const v = r[1];
      if (k != null && String(k).trim() !== "") kv[String(k).trim()] = v;
    }

    unidades.push({ ...u, kv });
  }

  return unidades;
}


// ====== render ======
function renderResumo(kv, turnos){
  const ticket = Number(kv.ticket_medio ?? 370);
  const perda = Number(kv.perda ?? 0.30);
  const custo = Number(kv.custo_aluno ?? 220);

  const ticketLiq = ticket * (1 - perda);

  el("kpi_ticket").textContent = brl(ticket);
  el("kpi_perda").textContent = pct(perda);
  el("kpi_ticket_liq").textContent = brl(ticketLiq);
  el("kpi_custo").textContent = brl(custo);
  el("kpi_lucro_turma").textContent = brl(kv.lucro_turma ?? (ticketLiq - custo) * Number(kv.alunos_turma ?? 35));
  el("kpi_lucro_total").textContent = brl(kv.lucro_total ?? null);

  // Turnos decisão (render simples)
  const lines = turnos.map(t=>{
    const turno = t.turno ?? "—";
    const lucro = brl(t.lucro_turma);
    const dec = (t.decisao || "").toUpperCase();
    const icon = dec.includes("NÃO") ? "❌" : "✅";
    return `<div style="margin:6px 0;"><strong>${turno}</strong>: ${icon} ${dec || "—"} <span style="color:#60786a">(${lucro})</span></div>`;
  }).join("");

  el("turnos_box").innerHTML = lines || "—";
}

function renderCursosABC(cursos){
  // tabela simples (p/ reunião hoje)
  const total = cursos.reduce((s,c)=>s + (Number(c.matriculas)||0), 0);
  const sorted = [...cursos].sort((a,b)=>(Number(b.matriculas)||0)-(Number(a.matriculas)||0));

  let acum = 0;
  const rows = sorted.map(c=>{
    const m = Number(c.matriculas)||0;
    const pctv = total ? (m/total) : 0;
    acum += pctv;
    const classe = acum <= 0.80 ? "A" : (acum <= 0.95 ? "B" : "C");
    return `
      <tr>
        <td>${c.curso}</td>
        <td style="text-align:right">${m.toLocaleString("pt-BR")}</td>
        <td style="text-align:right">${(pctv*100).toFixed(1)}%</td>
        <td style="text-align:right">${(acum*100).toFixed(1)}%</td>
        <td style="text-align:center"><strong>${classe}</strong></td>
      </tr>`;
  }).join("");

  el("cursos_table").innerHTML = `
    <div style="margin-bottom:8px;color:#4a5a52">Total matrículas: <strong>${total.toLocaleString("pt-BR")}</strong></div>
    <table style="width:100%;border-collapse:collapse">
      <thead>
        <tr style="text-align:left;border-bottom:1px solid #e4efe8">
          <th>Curso</th>
          <th style="text-align:right">Matrículas</th>
          <th style="text-align:right">% do Total</th>
          <th style="text-align:right">% Acum.</th>
          <th style="text-align:center">ABC</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
}

function renderCapacidade(kv){
  const salas = Number(kv.salas ?? 19);
  const turnos = Number(kv.turnos ?? 5);
  const alunosTurma = Number(kv.alunos_turma ?? 35);

  const slots = salas * turnos;
  const capTotal = slots * alunosTurma;

  el("cap_salas").textContent = salas.toLocaleString("pt-BR");
  el("cap_turnos").textContent = turnos.toLocaleString("pt-BR");
  el("cap_slots").textContent = slots.toLocaleString("pt-BR");
  el("cap_alunos_turma").textContent = alunosTurma.toLocaleString("pt-BR");
  el("cap_total_alunos").textContent = capTotal.toLocaleString("pt-BR");
}

function renderUnidadesCards(unidades){
  const box = document.getElementById("unidades_cards");
  if (!box) return;

  // escolhe algumas chaves “prováveis” para aparecer no card (se existirem)
  const cardKeys = ["faturamento", "receita", "lucro", "margem", "ticket_medio", "matriculas", "ativos", "inadimplencia", "evasao"];

  box.innerHTML = unidades.map(u=>{
    const keys = Object.keys(u.kv);
    const pick = cardKeys
      .map(k => keys.find(x => normalizeKey(x) === normalizeKey(k)))
      .filter(Boolean)
      .slice(0, 6);

    const lines = pick.map(k => `<div style="display:flex;justify-content:space-between;gap:10px;margin:6px 0">
        <span style="color:#4a5a52">${prettyLabel(k)}</span>
        <strong>${fmtValue(k, u.kv[k])}</strong>
      </div>`).join("");

    return `
      <div class="kpi" style="border:1px solid #e4efe8;border-radius:14px;padding:12px;background:#fff">
        <div style="font-weight:800;color:#1e3a2f;margin-bottom:6px">${u.nome}</div>
        ${lines || `<div style="color:#4a5a52">Sem dados (aba vazia?)</div>`}
      </div>
    `;
  }).join("");
}

//UNIDADES
function renderUnidadesTabela(unidades){
  const host = document.getElementById("unidades_table");
  if (!host) return;

  // união de todas as chaves
  const allKeysSet = new Set();
  unidades.forEach(u => Object.keys(u.kv).forEach(k => allKeysSet.add(k)));

  // ordenação com prioridade (se existir)
  const preferred = [
    "faturamento","receita","receita_total","custo_total","lucro","margem",
    "ticket_medio","ativos","matriculas","evasao","inadimplencia","perda"
  ];

  const allKeys = Array.from(allKeysSet);
  allKeys.sort((a,b)=>{
    const na = normalizeKey(a), nb = normalizeKey(b);
    const ia = preferred.indexOf(na);
    const ib = preferred.indexOf(nb);
    if (ia !== -1 || ib !== -1) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
    return na.localeCompare(nb, "pt-BR");
  });

  // calcula melhor/pior por linha quando forem números
  function getNumeric(k, v){
    const n = Number(v);
    return Number.isNaN(n) ? null : n;
  }

  const head = `
    <thead>
      <tr style="text-align:left;border-bottom:1px solid #e4efe8">
        <th>Indicador</th>
        ${unidades.map(u=>`<th style="text-align:right">${u.nome}</th>`).join("")}
      </tr>
    </thead>
  `;

  const bodyRows = allKeys.map(k=>{
    const values = unidades.map(u => u.kv[k]);
    const nums = values.map(v => getNumeric(k, v));

    const validNums = nums
      .map((n,i)=>({n,i}))
      .filter(x=>x.n != null);

    let bestIdx = null;
    if (validNums.length >= 2){
      const higherIsBetter = betterIsHigher(k);
      bestIdx = validNums.reduce((best,cur)=>{
        if (best == null) return cur;
        return higherIsBetter ? (cur.n > best.n ? cur : best) : (cur.n < best.n ? cur : best);
      }, null)?.i ?? null;
    }

    const tds = unidades.map((u, i)=>{
      const v = u.kv[k];
      const formatted = fmtValue(k, v);
      const style = (bestIdx === i) ? "color:#1f7a3b;font-weight:800" : "color:#1e3a2f";
      return `<td style="text-align:right;${style}">${formatted}</td>`;
    }).join("");

    return `<tr style="border-bottom:1px solid #f0f5f2">
      <td style="color:#4a5a52">${prettyLabel(k)}</td>
      ${tds}
    </tr>`;
  }).join("");

  host.innerHTML = `
    <div style="overflow:auto;border:1px solid #e4efe8;border-radius:12px">
      <table style="width:100%;border-collapse:collapse;min-width:720px">
        ${head}
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
  `;
}

function renderUnidades(unidades){
  renderUnidadesCards(unidades);
  renderUnidadesTabela(unidades);
}


// ====== main ======
async function main(){
  const status = el("status");

  //if (!SHEET_ID || SHEET_ID === "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX"){
  //  status.textContent = "Configure o SHEET_ID no app.js (uma única vez).";
  //  status.style.color = "darkred";
  //  return;
  // }

  try{
    status.textContent = "Atualizando…";

    // DADOS_API (KV)
    const jDados = await fetchGViz(TABS.DADOS, "A:B");
    const { cols: c1, rows: r1 } = tableToRows(jDados);
    const kv = rowsToKV(c1, r1);

    // CURSOS_ABC
    const jCursos = await fetchGViz(TABS.CURSOS, "A:B");
    const { rows: r2 } = tableToRows(jCursos);
    const cursos = r2
      .filter(r => r[0] != null)
      .map(r => ({ curso: String(r[0]), matriculas: Number(r[1] || 0) }));

    // TURNOS_DECISAO
    const jTurnos = await fetchGViz(TABS.TURNOS, "A:D");
    const { cols: c3, rows: r3 } = tableToRows(jTurnos);
    const idx = (name) => c3.findIndex(x => String(x).toLowerCase() === name);
    const iTurno = idx("turno");
    const iRec = idx("receita_liquida_aluno");
    const iLuc = idx("lucro_turma");
    const iDec = idx("decisao");

    const turnos = r3
      .filter(r => r[iTurno] != null)
      .map(r => ({
        turno: r[iTurno],
        receita_liquida_aluno: r[iRec],
        lucro_turma: r[iLuc],
        decisao: r[iDec],
      }));

    renderResumo(kv, turnos);
    renderCursosABC(cursos);
    renderCapacidade(kv);

    status.textContent = "Atualizado com sucesso ✅";
    status.style.color = "#2d6a4f";
  } catch(e){
    console.error(e);
    status.textContent = "Erro ao ler o Google Sheets (verifique publicação/compartilhamento e nomes das abas).";
    status.style.color = "darkred";
  }
}

// UNIDADES (nova tela)
    const unidades = await carregarUnidades();
    renderUnidades(unidades);

main();
