const SHEET_ID = "1bxfnVdmo0tcv1IKtSDkGYKW5V2plpKyN0oSkGXM_hI0";
const SHEET_NAME = "DADOS_API";

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
const el = id => document.getElementById(id);

async function fetchDados(){
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}&range=A:B`;
  const res = await fetch(url, { cache: "no-store" });
  const txt = await res.text();
  const json = JSON.parse(txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1));
  const rows = json.table.rows || [];
  const kv = {};
  rows.forEach(r=>{
    const k = r.c?.[0]?.v;
    const v = r.c?.[1]?.v;
    if(k!=null) kv[k]=v;
  });
  return kv;
}

async function main(){
  const status = el("status");
  try{
    status.textContent = "Atualizando…";
    const kv = await fetchDados();

    el("ticket_medio").textContent = brl(kv.ticket_medio);
    el("perda").textContent = pct(kv.perda);
    el("lucro_turma").textContent = brl(kv.lucro_turma);
    el("lucro_total").textContent = brl(kv.lucro_total);
    el("ocupacao").textContent = pct(kv.ocupacao);

    el("turmas_a").textContent = kv.turmas_a ?? "—";
    el("turmas_b").textContent = kv.turmas_b ?? "—";
    el("turmas_c").textContent = kv.turmas_c ?? "—";
    el("min_alunos").textContent = kv.min_alunos ?? "—";

    el("cen_70").textContent = brl(kv.cen_70);
    el("cen_80").textContent = brl(kv.cen_80);
    el("cen_90").textContent = brl(kv.cen_90);

    status.textContent = "Atualizado com sucesso ✅";
  } catch (e){
    console.error(e);
    status.textContent = "Erro ao ler o Google Sheets";
    status.style.color = "darkred";
  }
}

main();
