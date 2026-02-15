/**
 * KPIs GT — GVIZ (sem WebApp)
 * - Unidade + Mês base + Modo tabela
 * - Cards com setas + cores
 * - Tabela em caixa com scroll interno
 * - Tabela agrupada por categorias
 * - Destaque mês base e ano anterior
 * - Persistência (localStorage)
 */

const CONFIG = {
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  UNIDADES: [
    { label: "Cajazeiras", tab: "Números Cajazeiras" },
    { label: "Camaçari", tab: "Números Camaçari" },
    { label: "São Cristóvão", tab: "Números São Cristóvão" },
  ],
  DEBUG: false,
};

// Cards (ajuste os nomes se necessário — se não bater, aparece "—" e não quebra)
const KPI_CONFIG = [
  { key: "FATURAMENTO T (R$)", label: "Faturamento", better: "up", fmt: "brl" },
  { key: "CUSTO TOTAL T (R$)", label: "Custo Total", better: "down", fmt: "brl" },
  { key: "ALUNOS ATIVOS T", label: "Ativos", better: "up", fmt: "int" },
  { key: "META DE MATRÍCULA", label: "Meta", better: "up", fmt: "int" },
  { key: "MATRÍCULAS REALIZADAS", label: "Matrículas", better: "up", fmt: "int" },
  { key: "CONVERSÃO (%)", label: "Conversão", better: "up", fmt: "pct" },
  { key: "INADIMPLÊNCIA (META 9,3%)", label: "Inadimplência", better: "down", fmt: "pct" },
  { key: "EVASÃO REAL (META 4,8%)", label: "Evasão", better: "down", fmt: "pct" },
];

// Grupos da tabela (heurística por nome do indicador)
const GROUPS = [
  { name: "COMERCIAL", match: /(MATR|META|CONVERS|LEADS|AGEND|CAPTA|VISIT|PROPOST)/i },
  { name: "ACADÊMICO", match: /(ATIV|EVAS|ALUNO|TURMA|FALT|RETEN|CARGA|AULA)/i },
  { name: "FINANCEIRO", match: /(FATUR|RECEIT|BOLETO|INADIMPL|CUSTO|DESP|MARGEM|CAIXA|PAGAM)/i },
  { name: "OPERACIONAL", match: /(PROF|SAL|RH|OCUP|CAPAC|ESTRUT|EQUIPE)/i },
];

let LAST_RAW = null;     // { header, rows }
let LAST_SERIES = null;  // series por mês (YYYY-MM)

function $(id){ return document.getElementById(id); }
function safeText(v){ return (v ?? "").toString(); }
function escapeHtml(str){
  return safeText(str)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}
function setStatus(msg, type="info"){
  const el = $("status");
  if (el){
    el.textContent = msg;
    el.dataset.type = type;
  }
  console.log(`[${type}] ${msg}`);
}

/* ---------------------------
   PARSE DE MÊS (suporta):
   - "janeiro/23", "março/2024"
   - "jan 2023", "dez 2025"
--------------------------- */

function parseMonthKeyFromHeader(label){
  const v = safeText(label).trim().toLowerCase();

  // Padrão 1: "janeiro/23"
  let m = v.match(/^([a-zçãáâàéêíóôõúûü]{3,12})\/(\d{2}|\d{4})$/);
  if (m){
    const nome = normalizeMonthName(m[1]);
    let ano = m[2];
    if (ano.length === 2) ano = "20" + ano;
    const mes = monthNameToNumber(nome);
    if (!mes) return null;
    return `${ano}-${String(mes).padStart(2,"0")}`;
  }

  // Padrão 2: "jan 2023"
  m = v.match(/^([a-zçãáâàéêíóôõúûü]{3,12})\s(\d{4})$/);
  if (m){
    const nome = normalizeMonthName(m[1]);
    const ano = m[2];
    const mes = monthNameToNumber(nome);
    if (!mes) return null;
    return `${ano}-${String(mes).padStart(2,"0")}`;
  }

  return null;
}

function normalizeMonthName(s){
  return safeText(s)
    .trim()
    .toLowerCase()
    .replaceAll("ç","c")
    .replaceAll("ã","a")
    .replaceAll("á","a")
    .replaceAll("â","a")
    .replaceAll("à","a")
    .replaceAll("é","e")
    .replaceAll("ê","e")
    .replaceAll("í","i")
    .replaceAll("ó","o")
    .replaceAll("ô","o")
    .replaceAll("õ","o")
    .replaceAll("ú","u")
    .replaceAll("ü","u");
}

function monthNameToNumber(nome){
  const map = {
    jan:1, janeiro:1,
    fev:2, fevereiro:2,
    mar:3, marco:3,
    abr:4, abril:4,
    mai:5, maio:5,
    jun:6, junho:6,
    jul:7, julho:7,
    ago:8, agosto:8,
    set:9, setembro:9,
    out:10, outubro:10,
    nov:11, novembro:11,
    dez:12, dezembro:12
  };
  return map[nome] || null;
}

function addMonths(monthKey, delta){
  const [y,m] = monthKey.split("-").map(Number);
  const d = new Date(y, m-1, 1);
  d.setMonth(d.getMonth()+delta);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}
function getYear(monthKey){ return Number(monthKey.split("-")[0]); }

function formatMonthKeyBR(monthKey){
  const [y,m] = monthKey.split("-").map(Number);
  const map = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
  return `${map[m-1]}/${String(y).slice(2)}`;
}

function toNumberSmart(v){
  let s = safeText(v).trim();
  if (!s) return null;
  s = s.replaceAll("R$","").replaceAll(/\s+/g,"");
  s = s.replaceAll("%","");
  s = s.replaceAll(".","").replaceAll(",",".");
  const n = Number(s);
  if (Number.isNaN(n)) return null;
  return n;
}

/* ---------------------------
   GVIZ
--------------------------- */

function gvizUrl(sheetId, tabName){
  const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq`;
  const params = new URLSearchParams({
    sheet: tabName,
    headers: "1",
    tq: "select *",
    tqx: "out:json",
  });
  return `${base}?${params.toString()}`;
}

async function fetchGvizTable(sheetId, tabName){
  const url = gvizUrl(sheetId, tabName);
  const resp = await fetch(url);
  const text = await resp.text();
  const match = text.match(/setResponse\(([\s\S]+)\);\s*$/);
  if (!match) throw new Error("Resposta GVIZ inesperada (sem setResponse). Verifique permissões.");
  const json = JSON.parse(match[1]);
  if (json.status !== "ok") throw new Error(json.errors?.[0]?.detailed_message || "Erro no GVIZ");
  return json.table;
}

function tableToMatrix(table){
  const header = table.cols.map(c => safeText(c.label));
  const rows = table.rows.map(r => (r.c || []).map(cell => safeText(cell?.f ?? cell?.v ?? "")));
  return { header, rows };
}

/* ---------------------------
   SÉRIE POR MÊS + FECHAMENTO
--------------------------- */

function matrixToSeriesByMonth(header, rows){
  const monthCols = [];
  for (let i=0;i<header.length;i++){
    const key = parseMonthKeyFromHeader(header[i]);
    if (key) monthCols.push({ idx:i, key });
  }

  const series = {};
  monthCols.forEach(c => { series[c.key] = {}; });

  rows.forEach(r => {
    const indicador = safeText(r[0]).trim();
    if (!indicador) return;

    monthCols.forEach(c => {
      const raw = r[c.idx];
      const num = toNumberSmart(raw);
      series[c.key][indicador] = (num !== null ? num : safeText(raw).trim());
    });
  });

  return { series };
}

function pickCurrentMonthKey(series){
  const keys = Object.keys(series).sort();
  for (let i=keys.length-1;i>=0;i--){
    const k = keys[i];
    const obj = series[k];
    const hasSomething = obj && Object.values(obj).some(v => v !== "" && v !== null && v !== undefined);
    if (hasSomething) return k;
  }
  return keys[keys.length-1] || null;
}

function buildFechamento(series, monthKeyCurrent=null){
  const mesAtual = monthKeyCurrent || pickCurrentMonthKey(series);
  if (!mesAtual) throw new Error("Não consegui determinar o mês base.");

  const ano = getYear(mesAtual);
  const m1 = addMonths(mesAtual,-1);
  const m2 = addMonths(mesAtual,-2);
  const m3 = addMonths(mesAtual,-3);

  const ultimos3AnoCorrente = [m1,m2,m3].filter(k => k && getYear(k)===ano && series[k]);
  const mesAnoAnterior = addMonths(mesAtual,-12);
  const anoAnteriorExiste = (mesAnoAnterior && series[mesAnoAnterior]) ? mesAnoAnterior : null;

  const variacoes = {};
  const indicadores = new Set(Object.keys(series[mesAtual] || {}));

  indicadores.forEach(ind => {
    const atual = series[mesAtual]?.[ind] ?? null;
    const ant = (m1 && series[m1]) ? (series[m1][ind] ?? null) : null;
    const aa  = (anoAnteriorExiste) ? (series[anoAnteriorExiste][ind] ?? null) : null;

    const isNumA = typeof atual === "number";
    const isNumB = typeof ant === "number";
    const isNumC = typeof aa === "number";

    variacoes[ind] = {
      atual,
      mesAnterior: ant,
      anoAnterior: aa,
      deltaMesAnterior: (isNumA && isNumB) ? (atual-ant) : null,
      deltaAnoAnterior: (isNumA && isNumC) ? (atual-aa) : null,
    };
  });

  return { mesAtual, ultimos3AnoCorrente, mesAnoAnterior: anoAnteriorExiste, variacoes };
}

/* ---------------------------
   FILTRO DE COLUNAS (Fechamento)
--------------------------- */

function unique(arr){ return Array.from(new Set(arr.filter(Boolean))); }

function buildMonthsFocus(mesBase){
  return unique([
    mesBase,
    addMonths(mesBase,-1),
    addMonths(mesBase,-2),
    addMonths(mesBase,-3),
    addMonths(mesBase,-12),
  ]);
}

function filterMatrixByMonthKeys(raw, monthKeysWanted){
  const wanted = new Set(monthKeysWanted);
  const keepIdx = [0]; // indicador
  for (let i=1;i<raw.header.length;i++){
    const mk = parseMonthKeyFromHeader(raw.header[i]);
    if (mk && wanted.has(mk)) keepIdx.push(i);
  }
  const header = keepIdx.map(i => raw.header[i]);
  const rows = raw.rows.map(r => keepIdx.map(i => r[i] ?? ""));
  return { header, rows };
}

/* ---------------------------
   GRUPOS TABELA
--------------------------- */

function getGroupName(indicador){
  const t = safeText(indicador);
  for (const g of GROUPS){
    if (g.match.test(t)) return g.name;
  }
  return "OUTROS";
}

/* ---------------------------
   RENDER
--------------------------- */

function fmtValue(val, fmt){
  if (val === null || val === undefined || val === "") return "—";
  if (typeof val === "number"){
    if (fmt === "brl") return val.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
    if (fmt === "pct") return `${val.toLocaleString("pt-BR", { maximumFractionDigits:2 })}%`;
    if (fmt === "int") return val.toLocaleString("pt-BR", { maximumFractionDigits:0 });
    return val.toLocaleString("pt-BR", { maximumFractionDigits:2 });
  }
  return safeText(val);
}

function trendMeta(delta, better){
  if (delta === null || delta === undefined) return { arrow:"→", cls:"neutral", sign:"" };
  if (delta === 0) return { arrow:"→", cls:"neutral", sign:"" };

  const isUp = delta > 0;
  const good = (better === "up") ? isUp : !isUp;

  return { arrow: isUp ? "▲" : "▼", cls: good ? "good" : "bad", sign: delta>0 ? "+" : "" };
}

function renderFechamentoText(fechamento){
  const box = $("fechamento");
  if (!box) return;

  const tab = $("unidade")?.value || "";
  const labelUnidade = (CONFIG.UNIDADES.find(u => u.tab === tab)?.label) || tab;

  const ult3 = fechamento.ultimos3AnoCorrente.map(formatMonthKeyBR).join(" | ");

  box.textContent =
    `Unidade: ${labelUnidade}\n` +
    `Mês base: ${formatMonthKeyBR(fechamento.mesAtual)}\n` +
    `Últimos 3 (ano corrente): ${ult3 || "—"}\n` +
    `Mesmo mês ano anterior: ${fechamento.mesAnoAnterior ? formatMonthKeyBR(fechamento.mesAnoAnterior) : "—"}\n`;
}

function renderKPICards(fechamento){
  const wrap = $("cardsKPI");
  if (!wrap) return;

  wrap.innerHTML = "";

  const base = fechamento.mesAtual;
  const m1 = addMonths(base,-1);
  const m12 = fechamento.mesAnoAnterior;

  KPI_CONFIG.forEach(kpi => {
    const v = fechamento.variacoes?.[kpi.key] || {};
    const atual = v.atual;
    const dMes = v.deltaMesAnterior;
    const dAno = v.deltaAnoAnterior;

    const tMes = trendMeta(dMes, kpi.better);
    const tAno = trendMeta(dAno, kpi.better);

    const card = document.createElement("div");
    card.className = "kpiCard";

    const title = document.createElement("div");
    title.className = "kpiTitle";
    title.textContent = kpi.label;

    const value = document.createElement("div");
    value.className = "kpiValue";
    value.textContent = fmtValue(atual, kpi.fmt);

    const row1 = document.createElement("div");
    row1.className = "kpiRow";
    row1.innerHTML = `<span>vs ${m1 ? formatMonthKeyBR(m1) : "M-1"}</span>`;

    const d1 = document.createElement("span");
    d1.className = `kpiDelta ${tMes.cls}`;
    d1.textContent = `${tMes.arrow} ${tMes.sign}${fmtValue(dMes, kpi.fmt === "pct" ? "pct" : "num")}`;
    row1.appendChild(d1);

    const row2 = document.createElement("div");
    row2.className = "kpiRow";
    row2.innerHTML = `<span>vs ${m12 ? formatMonthKeyBR(m12) : "ano ant."}</span>`;

    const d2 = document.createElement("span");
    d2.className = `kpiDelta ${tAno.cls}`;
    d2.textContent = `${tAno.arrow} ${tAno.sign}${fmtValue(dAno, kpi.fmt === "pct" ? "pct" : "num")}`;
    row2.appendChild(d2);

    card.appendChild(title);
    card.appendChild(value);
    card.appendChild(row1);
    card.appendChild(row2);

    wrap.appendChild(card);
  });
}

function renderTableGrouped(header, rows, fechamento=null){
  const el = $("preview");
  if (!el) return;

  let idxMesBase = -1;
  let idxAnoAnterior = -1;

  if (fechamento){
    for (let i=0;i<header.length;i++){
      const mk = parseMonthKeyFromHeader(header[i]);
      if (mk && mk === fechamento.mesAtual) idxMesBase = i;
      if (mk && fechamento.mesAnoAnterior && mk === fechamento.mesAnoAnterior) idxAnoAnterior = i;
    }
  }

  // buckets
  const buckets = {};
  for (const r of rows){
    const ind = r?.[0] ?? "";
    const g = getGroupName(ind);
    if (!buckets[g]) buckets[g] = [];
    buckets[g].push(r);
  }

  const order = ["COMERCIAL","ACADÊMICO","FINANCEIRO","OPERACIONAL","OUTROS"];

  let html = `<table><thead><tr>`;
  header.forEach((h, idx) => {
    const cls = `${idx===idxMesBase ? "col-mesbase" : ""}${idx===idxAnoAnterior ? " col-anoanterior" : ""}`.trim();
    html += `<th class="${cls}">${escapeHtml(h)}</th>`;
  });
  html += `</tr></thead><tbody>`;

  order.forEach(groupName => {
    const list = buckets[groupName];
    if (!list || !list.length) return;

    html += `<tr class="group-row"><td colspan="${header.length}">${escapeHtml(groupName)}</td></tr>`;

    list.forEach(r => {
      html += `<tr>`;
      for (let i=0;i<header.length;i++){
        const v = r[i] ?? "";
        const cls = `${i===idxMesBase ? "col-mesbase" : ""}${i===idxAnoAnterior ? " col-anoanterior" : ""}`.trim();
        html += `<td class="${cls}">${escapeHtml(v)}</td>`;
      }
      html += `</tr>`;
    });
  });

  html += `</tbody></table>`;
  el.innerHTML = html;
}

function renderJSON(obj){
  if (!CONFIG.DEBUG) return;
  const el = $("json");
  if (!el) return;
  el.style.display = "block";
  el.textContent = JSON.stringify(obj, null, 2);
}

/* ---------------------------
   UI: selects + persistência
--------------------------- */

function fillUnidadeOptions(){
  const sel = $("unidade");
  if (!sel) return;

  sel.innerHTML = "";
  CONFIG.UNIDADES.forEach(u => {
    const opt = document.createElement("option");
    opt.value = u.tab;
    opt.textContent = u.label;
    sel.appendChild(opt);
  });

  const saved = localStorage.getItem("kpisgt_unidade_tab");
  const camacariTab = (CONFIG.UNIDADES.find(u => u.label.toLowerCase().includes("cama")) || {}).tab;

  sel.value = saved || camacariTab || CONFIG.UNIDADES[0]?.tab || sel.value;
}

function fillMesBaseOptions(series){
  const sel = $("mesBase");
  if (!sel) return;

  const keys = Object.keys(series).sort();
  sel.innerHTML = "";

  keys.forEach(k => {
    const opt = document.createElement("option");
    opt.value = k;
    opt.textContent = formatMonthKeyBR(k);
    sel.appendChild(opt);
  });

  const savedMes = localStorage.getItem("kpisgt_mes_base");
  const current = pickCurrentMonthKey(series) || keys[keys.length-1];

  sel.value = (savedMes && keys.includes(savedMes)) ? savedMes : current;
}

/* ---------------------------
   AÇÕES
--------------------------- */

async function carregarDados(){
  try{
    const tab = $("unidade")?.value || CONFIG.UNIDADES[0]?.tab;
    localStorage.setItem("kpisgt_unidade_tab", tab);

    setStatus("Carregando dados...", "info");
    const table = await fetchGvizTable(CONFIG.SHEET_ID, tab);

    LAST_RAW = tableToMatrix(table);

    const { series } = matrixToSeriesByMonth(LAST_RAW.header, LAST_RAW.rows);
    LAST_SERIES = series;

    fillMesBaseOptions(series);

    setStatus("Dados carregados.", "success");
    aplicar();
  }catch(err){
    console.error(err);
    setStatus(`ERRO: ${err.message}`, "error");
  }
}

function aplicar(){
  if (!LAST_RAW || !LAST_SERIES){
    setStatus("Clique em Carregar primeiro.", "warn");
    return;
  }

  const mesEscolhido = $("mesBase")?.value || null;
  if (mesEscolhido) localStorage.setItem("kpisgt_mes_base", mesEscolhido);

  const fechamento = buildFechamento(LAST_SERIES, mesEscolhido);
  renderFechamentoText(fechamento);
  renderKPICards(fechamento);

  const modo = $("modoTabela")?.value || "fechamento";

  if (modo === "completa"){
    renderTableGrouped(LAST_RAW.header, LAST_RAW.rows, fechamento);
  } else {
    const monthsFocus = buildMonthsFocus(fechamento.mesAtual);
    const filtered = filterMatrixByMonthKeys(LAST_RAW, monthsFocus);
    renderTableGrouped(filtered.header, filtered.rows, fechamento);
  }

  renderJSON({ fechamento });
  setStatus("Pronto.", "success");
}

/* ---------------------------
   INIT
--------------------------- */

document.addEventListener("DOMContentLoaded", () => {
  fillUnidadeOptions();

  $("btnCarregar")?.addEventListener("click", carregarDados);
  $("btnAplicar")?.addEventListener("click", aplicar);

  $("unidade")?.addEventListener("change", carregarDados);
  $("mesBase")?.addEventListener("change", aplicar);
  $("modoTabela")?.addEventListener("change", aplicar);

  setStatus("Pronto. Clique em Carregar.", "info");
});

// expor opcional
window.KPIsGT = { carregarDados, aplicar };
