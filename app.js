/**
 * Dashboard Escola – Mobile/PWA
 * Fonte: Google Sheets (aba DADOS_API) publicada ou compartilhada.
 *
 * Passo 1 (recomendado): no Google Sheets crie uma aba chamada DADOS_API com 2 colunas:
 *   A: chave   |  B: valor
 * Exemplo:
 *   ticket_medio | 370
 *   perda        | 0.30
 *   custo_aluno  | 220
 *   alunos_turma | 35
 *   turmas_a     | 40
 *   turmas_b     | 25
 *   turmas_c     | 11
 *   ocupacao     | 0.80
 *   min_alunos   | 28
 *   lucro_turma  | 1365
 *   lucro_total  | 103740
 *   cen_70       | 90090
 *   cen_80       | 103740
 *   cen_90       | 117390
 *
 * Passo 2: publique a planilha (Arquivo > Compartilhar > Publicar na web) ou use o modo "Qualquer pessoa com link - leitor".
 *
 * Passo 3: Cole aqui o SHEET_ID e (opcional) o NOME_DA_ABA (padrão: DADOS_API).
 */

/**
const CONFIG = {
  SHEET_ID: "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX",
  //"1bxfnVdmo0tcv1IKtSDkGYKW5V2plpKyN0oSkGXM_hI0", //1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX  //
 * Dashboard Escola – Mobile/PWA
 * Fonte: Google Sheets (aba DADOS_API) publicada ou compartilhada.
 *
 * Passo 1 (recomendado): no Google Sheets crie uma aba chamada DADOS_API com 2 colunas:
 *   A: chave   |  B: valor
 * Exemplo:
 *   ticket_medio | 370
 *   perda        | 0.30
 *   custo_aluno  | 220
 *   alunos_turma | 35
 *   turmas_a     | 40
 *   turmas_b     | 25
 *   turmas_c     | 11
 *   ocupacao     | 0.80
 *   min_alunos   | 28
 *   lucro_turma  | 1365
 *   lucro_total  | 103740
 *   cen_70       | 90090
 *   cen_80       | 103740
 *   cen_90       | 117390
 *
 * Passo 2: publique a planilha (Arquivo > Compartilhar > Publicar na web) ou use o modo "Qualquer pessoa com link - leitor".
 *
 * Passo 3: Cole aqui o SHEET_ID e (opcional) o NOME_DA_ABA (padrão: DADOS_API).
 /*
//"1bxfnVdmo0tcv1IKtSDkGYKW5V2plpKyN0oSkGXM_hI0", //1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX

const CONFIG = {
  SHEET_ID: "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX",
  ABA: "DADOS_API"
};

const SHEET_ID = CONFIG.SHEET_ID;
const SHEET_NAME = CONFIG.ABA;

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
function byId(id){ return document.getElementById(id); }

async function fetchDados(){
  // Método robusto: usar a Google Visualization API (retorna JSON sem precisar API key)
  // URL:
  // https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:json&sheet=DADOS_API&range=A:B
  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(CONFIG.ABA)}&range=A:B`;
  const res = await fetch(url, {cache:"no-store"});
  const txt = await res.text();

  // A resposta vem como "google.visualization.Query.setResponse({...});"
  const jsonText = txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1);
  const data = JSON.parse(jsonText);

  const rows = (data.table && data.table.rows) ? data.table.rows : [];
  const kv = {};
  for (const r of rows){
    const k = r.c?.[0]?.v;
    const v = r.c?.[1]?.v;
    if (k != null) kv[String(k).trim()] = v;
  }
  return kv;
}

function setBadge(el, ok, textOk, textBad){
  if(ok){
    el.classList.remove("bad");
    el.classList.add("ok");
    el.textContent = textOk;
  } else {
    el.classList.remove("ok");
    el.classList.add("bad");
    el.textContent = textBad;
  }
}

function render(kv){
  // Defaults
  const ticket = Number(kv.ticket_medio ?? 370);
  const perda = Number(kv.perda ?? 0.30);
  const turmasC = Number(kv.turmas_c ?? 0);
  const ocup = Number(kv.ocupacao ?? 0);
  const minAlunos = Number(kv.min_alunos ?? 999);

  byId("ticket_medio").textContent = brl(ticket);
  byId("perda").textContent = pct(perda);
  byId("lucro_turma").textContent = brl(kv.lucro_turma);
  byId("lucro_total").textContent = brl(kv.lucro_total);
  byId("ocupacao").textContent = pct(ocup);

  byId("turmas_a").textContent = kv.turmas_a ?? "—";
  byId("turmas_b").textContent = kv.turmas_b ?? "—";
  byId("turmas_c").textContent = kv.turmas_c ?? "—";
  byId("min_alunos").textContent = (kv.min_alunos ?? "—");

  byId("cen_70").textContent = brl(kv.cen_70);
  byId("cen_80").textContent = brl(kv.cen_80);
  byId("cen_90").textContent = brl(kv.cen_90);

  // Alerts (sim para tudo)
  const okTicket = ticket >= 370;
  const okC = turmasC <= 11;
  const ok30 = minAlunos >= 30;
  const okOcup = ocup >= 0.70;

  setBadge(byId("alert_ticket"), okTicket, "Ticket OK", "Ticket < 370");
  setBadge(byId("alert_c"), okC, "Turmas C OK", "C > 11");
  setBadge(byId("alert_30"), ok30, "Turmas ≥ 30", "< 30 alunos");
  setBadge(byId("alert_ocup"), okOcup, "Ocupação OK", "< 70%");

  const msgs = [];
  if(!okTicket) msgs.push("⚠️ Ticket médio abaixo de R$ 370.");
  if(!okC) msgs.push("⚠️ Turmas Classe C acima de 11.");
  if(!ok30) msgs.push("⚠️ Existe turma com menos de 30 alunos.");
  if(!okOcup) msgs.push("⚠️ Ocupação abaixo de 70% (risco de lucro).");
  byId("alertas_texto").textContent = msgs.length ? msgs.join(" ") : "✅ Tudo dentro das regras.";
}

async function main(){
  const status = byId("status");
  if (CONFIG.SHEET_ID === "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX"){ //1bxfnVdmo0tcv1IKtSDkGYKW5V2plpKyN0oSkGXM_hI0"
    status.textContent = "Falta configurar o SHEET_ID no app.js. Veja o README.";
    status.style.color = "#8a0000";
    return;
  }
  try{
    status.textContent = "Buscando dados do Google Sheets…";
    const kv = await fetchDados();
    render(kv);
    status.textContent = "Atualizado com sucesso ✅";
  } catch(e){
    console.error(e);
    status.textContent = "Não consegui ler o Google Sheets. Verifique publicação/compartilhamento e o SHEET_ID.";
    status.style.color = "#8a0000";
  }
}

main();
  ABA: "DADOS_API"
};

const SHEET_ID = CONFIG.SHEET_ID;
const SHEET_NAME = CONFIG.ABA;

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
function byId(id){ return document.getElementById(id); }

async function fetchDados(){
  // Método robusto: usar a Google Visualization API (retorna JSON sem precisar API key)
  // URL:
  // https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:json&sheet=DADOS_API&range=A:B
  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(CONFIG.ABA)}&range=A:B`;
  const res = await fetch(url, {cache:"no-store"});
  const txt = await res.text();

  // A resposta vem como "google.visualization.Query.setResponse({...});"
  const jsonText = txt.substring(txt.indexOf("{"), txt.lastIndexOf("}")+1);
  const data = JSON.parse(jsonText);

  const rows = (data.table && data.table.rows) ? data.table.rows : [];
  const kv = {};
  for (const r of rows){
    const k = r.c?.[0]?.v;
    const v = r.c?.[1]?.v;
    if (k != null) kv[String(k).trim()] = v;
  }
  return kv;
}

function setBadge(el, ok, textOk, textBad){
  if(ok){
    el.classList.remove("bad");
    el.classList.add("ok");
    el.textContent = textOk;
  } else {
    el.classList.remove("ok");
    el.classList.add("bad");
    el.textContent = textBad;
  }
}

function render(kv){
  // Defaults
  const ticket = Number(kv.ticket_medio ?? 370);
  const perda = Number(kv.perda ?? 0.30);
  const turmasC = Number(kv.turmas_c ?? 0);
  const ocup = Number(kv.ocupacao ?? 0);
  const minAlunos = Number(kv.min_alunos ?? 999);

  byId("ticket_medio").textContent = brl(ticket);
  byId("perda").textContent = pct(perda);
  byId("lucro_turma").textContent = brl(kv.lucro_turma);
  byId("lucro_total").textContent = brl(kv.lucro_total);
  byId("ocupacao").textContent = pct(ocup);

  byId("turmas_a").textContent = kv.turmas_a ?? "—";
  byId("turmas_b").textContent = kv.turmas_b ?? "—";
  byId("turmas_c").textContent = kv.turmas_c ?? "—";
  byId("min_alunos").textContent = (kv.min_alunos ?? "—");

  byId("cen_70").textContent = brl(kv.cen_70);
  byId("cen_80").textContent = brl(kv.cen_80);
  byId("cen_90").textContent = brl(kv.cen_90);

  // Alerts (sim para tudo)
  const okTicket = ticket >= 370;
  const okC = turmasC <= 11;
  const ok30 = minAlunos >= 30;
  const okOcup = ocup >= 0.70;

  setBadge(byId("alert_ticket"), okTicket, "Ticket OK", "Ticket < 370");
  setBadge(byId("alert_c"), okC, "Turmas C OK", "C > 11");
  setBadge(byId("alert_30"), ok30, "Turmas ≥ 30", "< 30 alunos");
  setBadge(byId("alert_ocup"), okOcup, "Ocupação OK", "< 70%");

  const msgs = [];
  if(!okTicket) msgs.push("⚠️ Ticket médio abaixo de R$ 370.");
  if(!okC) msgs.push("⚠️ Turmas Classe C acima de 11.");
  if(!ok30) msgs.push("⚠️ Existe turma com menos de 30 alunos.");
  if(!okOcup) msgs.push("⚠️ Ocupação abaixo de 70% (risco de lucro).");
  byId("alertas_texto").textContent = msgs.length ? msgs.join(" ") : "✅ Tudo dentro das regras.";
}

async function main(){
  const status = byId("status");
  if (CONFIG.SHEET_ID === "1D5w0o9nzhFjR2GQlMmN1hijASBcTYwXX"){ //1bxfnVdmo0tcv1IKtSDkGYKW5V2plpKyN0oSkGXM_hI0"
    status.textContent = "Falta configurar o SHEET_ID no app.js. Veja o README.";
    status.style.color = "#8a0000";
    return;
  }
  try{
    status.textContent = "Buscando dados do Google Sheets…";
    const kv = await fetchDados();
    render(kv);
    status.textContent = "Atualizado com sucesso ✅";
  } catch(e){
    console.error(e);
    status.textContent = "Não consegui ler o Google Sheets. Verifique publicação/compartilhamento e o SHEET_ID.";
    status.style.color = "#8a0000";
  }
}

main();
