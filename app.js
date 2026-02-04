// ====== CONFIG (troque só isso) ======
const SHEET_ID = "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX"; // ex: 1ho2rVMQW-E4dnygqrhUPUHSnItYqhnq0

const TABS = {
  DADOS: "DADOS_API",
  CURSOS: "CURSOS_ABC",
  TURNOS: "TURMAS_DECISAO",
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

// ====== main ======
async function main(){
  const status = el("status");

  if (!SHEET_ID || SHEET_ID === "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX"){
    status.textContent = "Configure o SHEET_ID no app.js (uma única vez).";
    status.style.color = "darkred";
    return;
  }

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

main();
