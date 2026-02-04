const CONFIG = {
  SHEET_ID: "1ho2rVMQW-E4dnygqrhUPUHSnItYqhnq0",
  ABA_KV: "DADOS_API",
  ABA_CURSOS: "CURSOS_ABC",
  ABA_TURNOS: "TURMAS_DECISAO",
};

const $ = (id) => document.getElementById(id);

function brl(n){
  const v = Number(n);
  if (Number.isNaN(v)) return "—";
  return v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
}
function pct01(n){
  const v = Number(n);
  if (Number.isNaN(v)) return "—";
  return (v*100).toFixed(1).replace(".",",") + "%";
}
function num(n){
  const v = Number(n);
  return Number.isNaN(v) ? null : v;
}

async function gvizFetch(sheetName, range){
  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}&range=${encodeURIComponent(range)}`;
  const res = await fetch(url, { cache: "no-store" });
  const txt = await res.text();
  const jsonText = txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1);
  return JSON.parse(jsonText);
}

async function fetchKV(){
  const j = await gvizFetch(CONFIG.ABA_KV, "A:B");
  const rows = j?.table?.rows || [];
  const kv = {};
  rows.forEach(r=>{
    const k = r.c?.[0]?.v;
    const v = r.c?.[1]?.v;
    if (k != null) kv[String(k).trim()] = v;
  });
  return kv;
}

async function fetchCursos(){
  // curso | matriculas
  const j = await gvizFetch(CONFIG.ABA_CURSOS, "A:B");
  const rows = j?.table?.rows || [];
  const data = [];
  rows.forEach(r=>{
    const curso = r.c?.[0]?.v;
    const matriculas = r.c?.[1]?.v;
    if(curso != null) data.push({curso:String(curso), matriculas:Number(matriculas||0)});
  });
  return data;
}

async function fetchTurnos(){
  // turno | receita_liquida_aluno | lucro_turma | decisao
  const j = await gvizFetch(CONFIG.ABA_TURNOS, "A:D");
  const rows = j?.table?.rows || [];
  const data = [];
  rows.forEach(r=>{
    const turno = r.c?.[0]?.v;
    if(turno == null) return;
    data.push({
      turno: String(turno),
      receita_liq_aluno: num(r.c?.[1]?.v),
      lucro_turma: num(r.c?.[2]?.v),
      decisao: (r.c?.[3]?.v != null) ? String(r.c?.[3]?.v) : "—"
    });
  });
  return data;
}

function setBadge(el, state){
  el.classList.remove("ok","warn","bad");
  el.classList.add(state);
}

function renderResumo(kv, cursos){
  const ticket = num(kv.ticket_medio ?? 370) ?? 370;
  const perda = num(kv.perda ?? 0.30) ?? 0.30;
  const custo = num(kv.custo_aluno ?? 220) ?? 220;
  const alunosTurma = num(kv.alunos_turma ?? 35) ?? 35;
  const salas = num(kv.salas ?? 19) ?? 19;
  const turnos = num(kv.turnos ?? 5) ?? 5;

  const ticketLiq = ticket * (1 - perda);
  const lucroAluno = ticketLiq - custo;
  const lucroTurma = lucroAluno * alunosTurma;

  const totalMatriculas = cursos.reduce((s,x)=>s + (x.matriculas||0), 0);
  const turmasPossiveis = salas * turnos;

  $("kpi_matriculas").textContent = totalMatriculas.toLocaleString("pt-BR");
  $("kpi_ticket").textContent = brl(ticket);
  $("kpi_ticket_liq").textContent = brl(ticketLiq);
  $("kpi_lucro_turma").textContent = brl(lucroTurma);
  $("kpi_capacidade").textContent = `${salas} × ${turnos}`;
  $("kpi_turmas_possiveis").textContent = turmasPossiveis.toLocaleString("pt-BR");

  // premissas
  const prem = [
    `Ticket nominal: ${brl(ticket)} (não reduz)`,
    `Perda estrutural: ${pct01(perda)}`,
    `Ticket líquido: ${brl(ticketLiq)}`,
    `Custo por aluno: ${brl(custo)}`,
    `Alunos por turma: ${alunosTurma}`,
    `Salas: ${salas}`,
    `Turnos: ${turnos}`,
    `Turmas possíveis: ${turmasPossiveis}`
  ];
  const ul = $("premissas");
  ul.innerHTML = prem.map(x=>`<li>${x}</li>`).join("");

  // badges/regras (piloto)
  const turmasCLimite = num(kv.turmas_c_limite ?? 11) ?? 11;
  const lucroMinTurma = num(kv.lucro_min_turma ?? 1000) ?? 1000;

  // Só indicativo (sem turmas C reais ainda no piloto)
  $("alertas_texto").textContent =
    `Regras ativas: Ticket ≥ ${brl(370)} • Turmas C ≤ ${turmasCLimite} • Abrir turma só acima de ${brl(lucroMinTurma)} (lucro/turma).`;

  $("b_ticket").textContent = "Ticket OK";
  setBadge($("b_ticket"), ticket >= 370 ? "ok" : "bad");

  $("b_c").textContent = `C ≤ ${turmasCLimite}`;
  setBadge($("b_c"), "ok");

  $("b_lucro").textContent = `Lucro/turma ${brl(lucroTurma)}`;
  setBadge($("b_lucro"), lucroTurma >= lucroMinTurma ? "ok" : "warn");
}

function computeABC(cursos){
  const arr = [...cursos].filter(x=>x.matriculas>0);
  arr.sort((a,b)=>b.matriculas - a.matriculas);
  const total = arr.reduce((s,x)=>s+x.matriculas,0);
  let acc = 0;
  return arr.map(x=>{
    const p = total ? x.matriculas/total : 0;
    acc += p;
    const cls = acc <= 0.80 ? "A" : (acc <= 0.95 ? "B" : "C");
    return {...x, pct:p, acc, cls};
  });
}

function renderCursos(cursos){
  const ordem = $("ordem_cursos").value;

  let base = computeABC(cursos);
  if(ordem === "asc") base = [...base].sort((a,b)=>a.matriculas-b.matriculas);
  if(ordem === "nome") base = [...base].sort((a,b)=>a.curso.localeCompare(b.curso,"pt-BR"));

  const total = cursos.reduce((s,x)=>s + (x.matriculas||0), 0);
  $("total_matriculas").textContent = total.toLocaleString("pt-BR");

  const tb = $("tbl_cursos").querySelector("tbody");
  tb.innerHTML = base.map(x=>`
    <tr>
      <td>${x.curso}</td>
      <td class="num">${x.matriculas.toLocaleString("pt-BR")}</td>
      <td class="num">${pct01(x.pct)}</td>
      <td class="num">${pct01(x.acc)}</td>
      <td><span class="pill ${x.cls==="A"?"ok":x.cls==="B"?"warn":"bad"}">${x.cls}</span></td>
    </tr>
  `).join("");

  const top3 = computeABC(cursos).slice(0,3);
  const topShare = top3.reduce((s,x)=>s+x.pct,0);
  $("insight_abc").textContent =
    `Insight rápido: Top 3 cursos = ${pct01(topShare)} das matrículas (piloto).`;
}

function renderTurnos(turnos){
  const tb = $("tbl_turnos").querySelector("tbody");
  tb.innerHTML = turnos.map(x=>{
    const d = (x.decisao||"").toUpperCase();
    const cls = d.includes("NÃO") ? "bad" : d.includes("SÓ") ? "warn" : "ok";
    return `
      <tr>
        <td>${x.turno}</td>
        <td class="num">${x.receita_liq_aluno==null?"—":brl(x.receita_liq_aluno)}</td>
        <td class="num">${x.lucro_turma==null?"—":brl(x.lucro_turma)}</td>
        <td><span class="pill ${cls}">${x.decisao}</span></td>
      </tr>
    `;
  }).join("");
}

function setupTabs(){
  const tabs = document.querySelectorAll(".tab");
  const map = {
    executivo: $("tab-executivo"),
    cursos: $("tab-cursos"),
    turnos: $("tab-turnos"),
  };

  function show(key){
    Object.values(map).forEach(s=>s.classList.add("hidden"));
    map[key].classList.remove("hidden");
    tabs.forEach(b=>b.classList.toggle("active", b.dataset.tab===key));
  }

  tabs.forEach(b=>{
    b.addEventListener("click", ()=>show(b.dataset.tab));
  });

  show("executivo");
}

async function main(){
  setupTabs();

  const status = $("status");
  try{
    status.textContent = "Buscando dados do Google Sheets…";

    const [kv, cursos, turnos] = await Promise.all([
      fetchKV(),
      fetchCursos(),
      fetchTurnos()
    ]);

    renderResumo(kv, cursos);
    renderCursos(cursos);
    renderTurnos(turnos);

    $("ordem_cursos").addEventListener("change", ()=>renderCursos(cursos));

    status.textContent = "Atualizado com sucesso ✅";
  } catch(e){
    console.error(e);
    status.textContent = "Erro ao ler o Google Sheets. Verifique: permissões da planilha, nomes das abas e colunas.";
    status.style.color = "darkred";
  }
}

main();
