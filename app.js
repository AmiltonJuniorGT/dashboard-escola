/**
 * KPIs GT — Mensal (GVIZ) — 3 abas (3 unidades)
 * - Período dinâmico (início/fim) para filtrar colunas do mês
 * - Mês base dos cards = mês do Período Fim
 * - Cards por blocos + Card Ritmo (mensal) com dias úteis
 * - Tabela navegável em caixa (scroll lateral/vertical), coluna A fixa, cabeçalho fixo
 * - Mapeamento por indicador (coluna A) salvo por unidade/aba
 */

const CONFIG = {
  SHEET_ID: "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go",
  UNIDADES: [
    { label: "Cajazeiras", tab: "Números Cajazeiras" },
    { label: "Camaçari", tab: "Números Camaçari" },
    { label: "São Cristóvão", tab: "Números São Cristóvão" },
  ],
};

// =======================
// Seções de cards (ordem executiva)
// =======================
const KPI_SECTIONS = [
  {
    title: "FINANCEIRO — RESULTADO",
    items: [
      { key: "FATURAMENTO", label: "Faturamento", better: "up", fmt: "brl" },
      { key: "CUSTOS", label: "Custos", better: "down", fmt: "brl" },
      { calc: "RESULTADO", label: "Resultado", better: "up", fmt: "brl" },
    ],
  },
  {
    title: "VOLUME — ABSOLUTOS",
    items: [
      { key: "ATIVOS", label: "Ativos", better: "up", fmt: "int" },
      { key: "VENDAS", label: "Vendas (Matrículas)", better: "up", fmt: "int" },
      { key: "PARCELAS_RECEBIDAS", label: "Parcelas recebidas", better: "up", fmt: "int" },
    ],
  },
  {
    title: "PERDAS — SAÚDE",
    items: [
      { key: "INADIMPLENCIA", label: "Inadimplência", better: "down", fmt: "pct" },
      { key: "EVASAO", label: "Evasão", better: "down", fmt: "pct" },
      { key: "FORMATURA", label: "Formatura", better: "up", fmt: "int" },
    ],
  },
  {
    title: "ATIVOS FINAL — FECHAMENTO",
    items: [
      { key: "ATIVOS_FINAL", label: "Ativos final", better: "up", fmt: "int" },
    ],
  },
  {
    title: "COMERCIAL — RITMO DO MÊS (MENSAL)",
    items: [
      { calc: "RITMO_META", label: "Ritmo p/ Meta", better: "up", fmt: "ritmo" },
    ],
  },
];

// =======================
// Grupos para tabela
// =======================
const GROUPS = [
  { name: "COMERCIAL", match: /(MATR|META|CONVERS|LEADS|AGEND|CAPTA|VISIT|PROPOST|VENDA)/i },
  { name: "ACADÊMICO", match: /(ATIV|EVAS|ALUNO|TURMA|FALT|RETEN|CARGA|AULA|FORMAT)/i },
  { name: "FINANCEIRO", match: /(FATUR|RECEIT|BOLETO|INADIMPL|CUSTO|DESP|MARGEM|CAIXA|PAGAM|PARCEL)/i },
  { name: "OPERACIONAL", match: /(PROF|SAL|RH|OCUP|CAPAC|ESTRUT|EQUIPE)/i },
];

let LAST_RAW = null;     // { header, rows }
let LAST_SERIES = null;  // { "YYYY-MM": { indicador: valorNumberOuTexto } }
let LAST_MONTH_COLS = null; // ["YYYY-MM", ...]
let LAST_FECHAMENTO = null;

// =======================
// Utils DOM
// =======================
function $(id){ return document.getElementById(id); }
function safeText(v){ return (v ?? "").toString(); }

function setStatus(msg, type="info"){
  const el = $("status");
  if (el){
    el.textContent = msg;
    el.dataset.type = type;
  }
  console.log(`[${type}] ${msg}`);
}

function escapeHtml(str){
  return safeText(str)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function isNum(x){ return typeof x === "number" && !Number.isNaN(x); }

// =======================
// Datas
// =======================
function toISODate(d){
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const day = String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}
function todayISO(){ return toISODate(new Date()); }
function startOfCurrentMonthISO(){
  const d = new Date();
  d.setDate(1);
  return toISODate(d);
}
function isoDateToMonthKey(iso){
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
  return iso.slice(0,7);
}

function parseMonthKey(monthKey){
  const [y,m] = monthKey.split("-").map(Number);
  return new Date(y, m-1, 1);
}
function addMonths(monthKey, delta){
  const d = parseMonthKey(monthKey);
  d.setMonth(d.getMonth()+delta);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}
function getYear(monthKey){ return Number(monthKey.split("-")[0]); }

function formatMonthKeyBR(monthKey){
  const [y,m] = monthKey.split("-").map(Number);
  const map = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
  return `${map[m-1]}/${String(y).slice(2)}`;
}

// dias úteis (sem feriados por enquanto)
function getMonthBusinessDays(monthKey){
  const d0 = parseMonthKey(monthKey);
  const y = d0.getFullYear(), m = d0.getMonth();
  const last = new Date(y, m+1, 0);
  let c = 0;
  for (let d = new Date(y,m,1); d <= last; d.setDate(d.getDate()+1)){
    const wd = d.getDay();
    if (wd !== 0 && wd !== 6) c++;
  }
  return c;
}
function getBusinessDaysElapsed(monthKey, refDateISO){
  const d0 = parseMonthKey(monthKey);
  const y = d0.getFullYear(), m = d0.getMonth();
  const ref = new Date(refDateISO + "T00:00:00");
  const end = new Date(y, m+1, 0);
  const stop = ref < end ? ref : end;

  let c = 0;
  for (let d = new Date(y,m,1); d <= stop; d.setDate(d.getDate()+1)){
    const wd = d.getDay();
    if (wd !== 0 && wd !== 6) c++;
  }
  return c;
}

// =======================
// Parse meses no cabeçalho (aceita "jan 2023" e "janeiro/23")
// =======================
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

function parseMonthKeyFromHeader(label){
  const v = safeText(label).trim().toLowerCase();

  // "janeiro/23"
  let m = v.match(/^([a-zçãáâàéêíóôõúûü]{3,12})\/(\d{2}|\d{4})$/);
  if (m){
    const nome = normalizeMonthName(m[1]);
    let ano = m[2];
    if (ano.length === 2) ano = "20" + ano;
    const mes = monthNameToNumber(nome);
    if (!mes) return null;
    return `${ano}-${String(mes).padStart(2,"0")}`;
  }

  // "jan 2023"
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

function monthKeyBetween(mk, startMk, endMk){
  return mk >= startMk && mk <= endMk;
}

// =======================
// Conversão numérica BR
// =======================
function toNumberSmart(v){
  let s = safeText(v).trim();
  if (!s) return null;

  s = s.replaceAll("R$","").replaceAll(/\s+/g,"");
  const isPct = s.includes("%");
  s = s.replaceAll("%","");

  s = s.replaceAll(".","").replaceAll(",",".");
  const n = Number(s);
  if (Number.isNaN(n)) return null;

  // porcentagem: guardamos o valor numérico "puro"
  // (renderização decide se coloca %)
  return n;
}

// =======================
// GVIZ
// =======================
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
  if (json.status !== "ok") throw new Error(json.errors?.[0]?.detailed_message || "Erro no GVIZ.");
  return json.table;
}

function tableToMatrix(table){
  const header = table.cols.map(c => safeText(c.label));
  const rows = table.rows.map(r => (r.c || []).map(cell => safeText(cell?.f ?? cell?.v ?? "")));
  return { header, rows };
}

// =======================
// Série por mês: { "YYYY-MM": { Indicador: valor } }
// =======================
function matrixToSeriesByMonth(header, rows){
  const monthCols = []; // {idx, key}
  for (let i=0;i<header.length;i++){
    const key = parseMonthKeyFromHeader(header[i]);
    if (key) monthCols.push({ idx:i, key });
  }

  const series = {};
  const monthKeys = monthCols.map(c => c.key);
  monthKeys.forEach(k => { series[k] = {}; });

  rows.forEach(r => {
    const indicador = safeText(r?.[0]).trim();
    if (!indicador) return;

    monthCols.forEach(c => {
      const raw = r[c.idx];
      const num = toNumberSmart(raw);
      series[c.key][indicador] = (num !== null ? num : safeText(raw).trim());
    });
  });

  return { series, monthKeys };
}

// =======================
// Fechamento mensal (M-1/M-2/M-3 + AA)
// =======================
function buildFechamento(series, mesBase){
  if (!mesBase || !series[mesBase]) throw new Error("Mês base inválido/fora da base.");

  const ano = getYear(mesBase);
  const m1 = addMonths(mesBase, -1);
  const m2 = addMonths(mesBase, -2);
  const m3 = addMonths(mesBase, -3);

  const ultimos3AnoCorrente = [m1,m2,m3].filter(k => k && getYear(k)===ano && series[k]);

  const mesAnoAnterior = addMonths(mesBase, -12);
  const aa = (mesAnoAnterior && series[mesAnoAnterior]) ? mesAnoAnterior : null;

  // variacoes por indicador (para usar em cards)
  const variacoes = {};
  const indicadores = new Set(Object.keys(series[mesBase] || {}));

  indicadores.forEach(ind => {
    const atual = series[mesBase]?.[ind] ?? null;
    const ant = (m1 && series[m1]) ? (series[m1][ind] ?? null) : null;
    const anoAnt = aa ? (series[aa][ind] ?? null) : null;

    const dMes = (isNum(atual) && isNum(ant)) ? (atual - ant) : null;
    const dAno = (isNum(atual) && isNum(anoAnt)) ? (atual - anoAnt) : null;

    variacoes[ind] = { atual, mesAnterior: ant, anoAnterior: anoAnt, dMes, dAno };
  });

  return { mesAtual: mesBase, ultimos3AnoCorrente, mesAnoAnterior: aa, variacoes };
}

// =======================
// Render helpers
// =======================
function fmtValue(val, fmt){
  if (val === null || val === undefined || val === "") return "—";

  if (fmt === "ritmo") return safeText(val);

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

// =======================
// Tabela — agrupamento
// =======================
function getGroupName(indicador){
  const t = safeText(indicador);
  for (const g of GROUPS){
    if (g.match.test(t)) return g.name;
  }
  return "OUTROS";
}

function filterMatrixByMonthRange(raw, startMk, endMk){
  const keepIdx = [];

  // sempre manter a coluna 0 (indicador)
  keepIdx.push(0);

  // manter colunas não-mês (exceto 0) e colunas de mês dentro do intervalo
  for (let i=1;i<raw.header.length;i++){
    const mk = parseMonthKeyFromHeader(raw.header[i]);
    if (!mk){
      keepIdx.push(i); // não é mês, mantém
      continue;
    }
    if (monthKeyBetween(mk, startMk, endMk)) keepIdx.push(i);
  }

  const header = keepIdx.map(i => raw.header[i]);
  const rows = raw.rows.map(r => keepIdx.map(i => r[i] ?? ""));
  return { header, rows };
}

function renderTableGrouped(header, rows, fechamento){
  const el = $("preview");
  if (!el) return;

  // achar índice do mês base e do AA na matriz atual
  let idxMesBase = -1;
  let idxAnoAnterior = -1;

  for (let i=0;i<header.length;i++){
    const mk = parseMonthKeyFromHeader(header[i]);
    if (mk && mk === fechamento.mesAtual) idxMesBase = i;
    if (mk && fechamento.mesAnoAnterior && mk === fechamento.mesAnoAnterior) idxAnoAnterior = i;
  }

  const buckets = {};
  rows.forEach(r => {
    const ind = r?.[0] ?? "";
    const g = getGroupName(ind);
    if (!buckets[g]) buckets[g] = [];
    buckets[g].push(r);
  });

  const order = ["FINANCEIRO","COMERCIAL","ACADÊMICO","OPERACIONAL","OUTROS"];

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

// =======================
// Mapeamento por indicador (coluna A)
// =======================
const MAP_KEYS = [
  { key: "FATURAMENTO", label: "Faturamento (R$)" },
  { key: "CUSTOS", label: "Custos (R$)" },
  { key: "VENDAS", label: "Vendas / Matrículas" },
  { key: "META_VENDAS", label: "Meta (Matrículas)" },
  { key: "ATIVOS", label: "Ativos" },
  { key: "ATIVOS_FINAL", label: "Ativos final" },
  { key: "PARCELAS_RECEBIDAS", label: "Parcelas recebidas" },
  { key: "INADIMPLENCIA", label: "Inadimplência (%)" },
  { key: "EVASAO", label: "Evasão (%)" },
  { key: "FORMATURA", label: "Formatura" },
];

function mapStorageKey(tabName){
  return `kpisgt_map_${CONFIG.SHEET_ID}_${tabName}`;
}

function loadMapping(tabName){
  try{
    const raw = localStorage.getItem(mapStorageKey(tabName));
    return raw ? JSON.parse(raw) : {};
  }catch(_){
    return {};
  }
}

function saveMapping(tabName, mapping){
  localStorage.setItem(mapStorageKey(tabName), JSON.stringify(mapping));
}

function getIndicatorNamesFromRows(rows){
  return rows
    .map(r => safeText(r?.[0]).trim())
    .filter(v => v && v.toLowerCase() !== "indicador");
}

function buildMappingUI(indicadores, tabName){
  const wrap = $("mapWrap");
  if (!wrap) return;

  wrap.innerHTML = "";
  const current = loadMapping(tabName);

  const opts = [`<option value="">— selecione —</option>`]
    .concat(indicadores.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`))
    .join("");

  MAP_KEYS.forEach(item => {
    const div = document.createElement("div");
    div.className = "field";
    div.innerHTML = `
      <label>${item.label}</label>
      <select data-mapkey="${item.key}">
        ${opts}
      </select>
    `;
    wrap.appendChild(div);

    const sel = div.querySelector("select");
    const saved = current[item.key] || "";
    if (saved) sel.value = saved;
  });

  const btn = $("btnSalvarMap");
  if (btn){
    btn.onclick = () => {
      const mapping = {};
      wrap.querySelectorAll("select[data-mapkey]").forEach(sel => {
        const k = sel.getAttribute("data-mapkey");
        const v = sel.value || "";
        if (v) mapping[k] = v;
      });
      saveMapping(tabName, mapping);
      window.__KPI_MAP_CURRENT = mapping;
      setStatus("Mapeamento salvo. Dados carregados.", "success");
      aplicar(); // redesenha cards/tabela com o mapa novo
    };
  }
}

function getMonthlyValue(series, monthKey, indicadorNome){
  if (!series?.[monthKey] || !indicadorNome) return null;
  const v = series[monthKey][indicadorNome];
  return (v === undefined ? null : v);
}

function buildMappedValues(series, monthKey, tabName){
  const map = loadMapping(tabName);
  window.__KPI_MAP_CURRENT = map;

  const out = {};
  MAP_KEYS.forEach(k => {
    const real = map[k.key];
    out[k.key] = getMonthlyValue(series, monthKey, real);
  });

  return { values: out, map };
}

// =======================
// Cards (seções + Ritmo)
// =======================
function fmtRitmoText(obj){
  const meta = isNum(obj.meta) ? obj.meta : null;
  const real = isNum(obj.realizado) ? obj.realizado : null;

  const duMes = obj.diasUteisMes || 0;
  const duElap = obj.diasUteisElap || 0;
  const duRest = Math.max(duMes - duElap, 0);

  const pct = (meta && meta !== 0 && real !== null) ? (real / meta) : null;

  return [
    `Meta: ${meta!==null ? meta.toLocaleString("pt-BR") : "—"} | Realizado: ${real!==null ? real.toLocaleString("pt-BR") : "—"} | %: ${pct!==null ? (pct*100).toFixed(1)+"%" : "—"}`,
    `Dias úteis: ${duElap}/${duMes} | Restantes: ${duRest}`,
    `Projeção (ritmo atual): ${isNum(obj.proj) ? obj.proj.toLocaleString("pt-BR") : "—"} | Gap p/ meta: ${isNum(obj.gap) ? obj.gap.toLocaleString("pt-BR") : "—"}`,
    `Necessário/dia útil (restante): ${isNum(obj.necessarioDia) ? obj.necessarioDia.toFixed(2) : "—"}`
  ].join("\n");
}

function renderFechamentoText(tabName, startMk, endMk, fechamento){
  const box = $("fechamento");
  if (!box) return;

  const labelUnidade = (CONFIG.UNIDADES.find(u => u.tab === tabName)?.label) || tabName;
  const ult3 = fechamento.ultimos3AnoCorrente.map(formatMonthKeyBR).join(" | ");

  box.textContent =
    `Unidade: ${labelUnidade}\n` +
    `Período: ${formatMonthKeyBR(startMk)} → ${formatMonthKeyBR(endMk)}\n` +
    `Mês base (cards): ${formatMonthKeyBR(fechamento.mesAtual)}\n` +
    `Últimos 3 (ano corrente): ${ult3 || "—"}\n` +
    `Mesmo mês ano anterior: ${fechamento.mesAnoAnterior ? formatMonthKeyBR(fechamento.mesAnoAnterior) : "—"}\n`;
}

function renderKPICardsMensal(tabName, fechamento, mappedValues, periodoFimISO){
  const wrap = $("cardsKPI");
  if (!wrap) return;
  wrap.innerHTML = "";

  const base = fechamento.mesAtual;
  const m1 = addMonths(base,-1);
  const aa = fechamento.mesAnoAnterior;

  function deltaOf(realName){
    if (!realName) return { atual:null, dMes:null, dAno:null };

    const atual = getMonthlyValue(LAST_SERIES, base, realName);
    const ant  = (m1 && LAST_SERIES[m1]) ? getMonthlyValue(LAST_SERIES, m1, realName) : null;
    const anoA = aa ? getMonthlyValue(LAST_SERIES, aa, realName) : null;

    const dMes = (isNum(atual) && isNum(ant)) ? (atual - ant) : null;
    const dAno = (isNum(atual) && isNum(anoA)) ? (atual - anoA) : null;

    return { atual, dMes, dAno };
  }

  const map = loadMapping(tabName);

  KPI_SECTIONS.forEach(section => {
    const title = document.createElement("div");
    title.className = "kpiSectionTitle";
    title.textContent = section.title;
    wrap.appendChild(title);

    const grid = document.createElement("div");
    grid.className = "kpiSectionGrid";
    wrap.appendChild(grid);

    section.items.forEach(item => {
      // Ritmo (mensal)
      if (item.calc === "RITMO_META"){
        const meta = mappedValues["META_VENDAS"];
        const vendas = mappedValues["VENDAS"];

        const diasUteisMes = getMonthBusinessDays(base);
        const diasUteisElap = getBusinessDaysElapsed(base, periodoFimISO);
        const diasRest = Math.max(diasUteisMes - diasUteisElap, 0);

        const ritmoAtualDia = (isNum(vendas) && diasUteisElap > 0) ? (vendas / diasUteisElap) : null;
        const proj = (isNum(ritmoAtualDia) && diasUteisMes > 0) ? (ritmoAtualDia * diasUteisMes) : null;

        const gap = (isNum(meta) && isNum(proj)) ? (meta - proj) : null;

        const restante = (isNum(meta) && isNum(vendas)) ? (meta - vendas) : null;
        const necessarioDia = (isNum(restante) && diasRest > 0) ? (restante / diasRest) : null;

        const card = document.createElement("div");
        card.className = "kpiCard";

        const t = document.createElement("div");
        t.className = "kpiTitle";
        t.textContent = item.label;

        const v = document.createElement("div");
        v.className = "kpiValue";
        v.textContent = "Correção comercial (mensal)";

        const pre = document.createElement("pre");
        pre.className = "kpiRitmoPre";
        pre.textContent = fmtRitmoText({ meta, realizado: vendas, diasUteisMes, diasUteisElap, proj, gap, necessarioDia });

        card.appendChild(t);
        card.appendChild(v);
        card.appendChild(pre);

        grid.appendChild(card);
        return;
      }

      // Resultado (calc) com deltas calculados
      if (item.calc === "RESULTADO"){
        const fatName = map.FATURAMENTO;
        const cusName = map.CUSTOS;

        const fat = getMonthlyValue(LAST_SERIES, base, fatName);
        const cus = getMonthlyValue(LAST_SERIES, base, cusName);
        const atual = (isNum(fat) && isNum(cus)) ? (fat - cus) : null;

        const fat_m1 = (m1 && LAST_SERIES[m1]) ? getMonthlyValue(LAST_SERIES, m1, fatName) : null;
        const cus_m1 = (m1 && LAST_SERIES[m1]) ? getMonthlyValue(LAST_SERIES, m1, cusName) : null;
        const ant = (isNum(fat_m1) && isNum(cus_m1)) ? (fat_m1 - cus_m1) : null;

        const fat_aa = aa ? getMonthlyValue(LAST_SERIES, aa, fatName) : null;
        const cus_aa = aa ? getMonthlyValue(LAST_SERIES, aa, cusName) : null;
        const anoA = (isNum(fat_aa) && isNum(cus_aa)) ? (fat_aa - cus_aa) : null;

        const dMes = (isNum(atual) && isNum(ant)) ? (atual - ant) : null;
        const dAno = (isNum(atual) && isNum(anoA)) ? (atual - anoA) : null;

        const tMes = trendMeta(dMes, item.better);
        const tAno = trendMeta(dAno, item.better);

        const card = document.createElement("div");
        card.className = "kpiCard";

        const titleEl = document.createElement("div");
        titleEl.className = "kpiTitle";
        titleEl.textContent = item.label;

        const valueEl = document.createElement("div");
        valueEl.className = "kpiValue";
        valueEl.textContent = fmtValue(atual, item.fmt);

        const row1 = document.createElement("div");
        row1.className = "kpiRow";
        row1.innerHTML = `<span>vs ${m1 ? formatMonthKeyBR(m1) : "M-1"}</span>`;
        const d1 = document.createElement("span");
        d1.className = `kpiDelta ${tMes.cls}`;
        d1.textContent = `${tMes.arrow} ${tMes.sign}${fmtValue(dMes, item.fmt)}`;
        row1.appendChild(d1);

        const row2 = document.createElement("div");
        row2.className = "kpiRow";
        row2.innerHTML = `<span>vs ${aa ? formatMonthKeyBR(aa) : "AA"}</span>`;
        const d2 = document.createElement("span");
        d2.className = `kpiDelta ${tAno.cls}`;
        d2.textContent = `${tAno.arrow} ${tAno.sign}${fmtValue(dAno, item.fmt)}`;
        row2.appendChild(d2);

        card.appendChild(titleEl);
        card.appendChild(valueEl);
        card.appendChild(row1);
        card.appendChild(row2);

        grid.appendChild(card);
        return;
      }

      // Card normal (mapeado)
      const realName = map[item.key];
      const { atual, dMes, dAno } = deltaOf(realName);

      const tMes = trendMeta(dMes, item.better);
      const tAno = trendMeta(dAno, item.better);

      const card = document.createElement("div");
      card.className = "kpiCard";

      const titleEl = document.createElement("div");
      titleEl.className = "kpiTitle";
      titleEl.textContent = item.label;

      const valueEl = document.createElement("div");
      valueEl.className = "kpiValue";
      valueEl.textContent = fmtValue(atual, item.fmt);

      const row1 = document.createElement("div");
      row1.className = "kpiRow";
      row1.innerHTML = `<span>vs ${m1 ? formatMonthKeyBR(m1) : "M-1"}</span>`;
      const d1 = document.createElement("span");
      d1.className = `kpiDelta ${tMes.cls}`;
      d1.textContent = `${tMes.arrow} ${tMes.sign}${fmtValue(dMes, item.fmt)}`;
      row1.appendChild(d1);

      const row2 = document.createElement("div");
      row2.className = "kpiRow";
      row2.innerHTML = `<span>vs ${aa ? formatMonthKeyBR(aa) : "AA"}</span>`;
      const d2 = document.createElement("span");
      d2.className = `kpiDelta ${tAno.cls}`;
      d2.textContent = `${tAno.arrow} ${tAno.sign}${fmtValue(dAno, item.fmt)}`;
      row2.appendChild(d2);

      card.appendChild(titleEl);
      card.appendChild(valueEl);
      card.appendChild(row1);
      card.appendChild(row2);

      grid.appendChild(card);
    });
  });
}

// =======================
// UI: Unidade + default datas
// =======================
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

function setDefaultDatesIfEmpty(){
  const ini = $("periodoInicio");
  const fim = $("periodoFim");
  if (!ini || !fim) return;

  const savedIni = localStorage.getItem("kpisgt_periodo_inicio");
  const savedFim = localStorage.getItem("kpisgt_periodo_fim");

  ini.value = savedIni || ini.value || startOfCurrentMonthISO();
  fim.value = savedFim || fim.value || todayISO();
}

function persistDates(){
  const ini = $("periodoInicio")?.value;
  const fim = $("periodoFim")?.value;
  if (ini) localStorage.setItem("kpisgt_periodo_inicio", ini);
  if (fim) localStorage.setItem("kpisgt_periodo_fim", fim);
}

// =======================
// Ações
// =======================
async function carregarDados(){
  try{
    const tab = $("unidade")?.value || CONFIG.UNIDADES[0]?.tab;
    localStorage.setItem("kpisgt_unidade_tab", tab);

    setStatus("Carregando dados...", "info");

    const table = await fetchGvizTable(CONFIG.SHEET_ID, tab);
    LAST_RAW = tableToMatrix(table);

    const { series, monthKeys } = matrixToSeriesByMonth(LAST_RAW.header, LAST_RAW.rows);
    LAST_SERIES = series;
    LAST_MONTH_COLS = monthKeys;

    // monta mapeamento UI com base na coluna A
    const indicadores = getIndicatorNamesFromRows(LAST_RAW.rows);
    buildMappingUI(indicadores, tab);

    setStatus("Dados carregados.", "success");

    aplicar();
  }catch(err){
    console.error(err);
    setStatus(`ERRO: ${err.message}`, "error");
  }
}

function aplicar(){
  try{
    if (!LAST_RAW || !LAST_SERIES){
      setStatus("Clique em Carregar primeiro.", "warn");
      return;
    }

    persistDates();

    const tab = $("unidade")?.value || CONFIG.UNIDADES[0]?.tab;
    const startMk = isoDateToMonthKey($("periodoInicio")?.value);
    const endMk = isoDateToMonthKey($("periodoFim")?.value);

    if (!startMk || !endMk){
      setStatus("Informe Período Início e Fim.", "warn");
      return;
    }
    if (startMk > endMk){
      setStatus("Período inválido: Início maior que Fim.", "warn");
      return;
    }

    const mesBase = endMk;
    if (!LAST_SERIES[mesBase]){
      setStatus("O mês do Período Fim não existe na base (sem coluna correspondente).", "warn");
      return;
    }

    const fechamento = buildFechamento(LAST_SERIES, mesBase);
    LAST_FECHAMENTO = fechamento;

    const periodoFimISO = $("periodoFim")?.value || todayISO();

    const { values: mappedValues } = buildMappedValues(LAST_SERIES, mesBase, tab);

    renderFechamentoText(tab, startMk, endMk, fechamento);
    renderKPICardsMensal(tab, fechamento, mappedValues, periodoFimISO);

    const modo = $("modoTabela")?.value || "intervalo";
    if (modo === "completa"){
      renderTableGrouped(LAST_RAW.header, LAST_RAW.rows, fechamento);
    }else{
      const filtered = filterMatrixByMonthRange(LAST_RAW, startMk, endMk);
      renderTableGrouped(filtered.header, filtered.rows, fechamento);
    }

    setStatus("Dados carregados.", "success");
  }catch(err){
    console.error(err);
    setStatus(`ERRO: ${err.message}`, "error");
  }
}

// =======================
// init
// =======================
document.addEventListener("DOMContentLoaded", () => {
  fillUnidadeOptions();
  setDefaultDatesIfEmpty();

  $("btnCarregar")?.addEventListener("click", carregarDados);
  $("btnAplicar")?.addEventListener("click", aplicar);

  $("unidade")?.addEventListener("change", carregarDados);
  $("periodoInicio")?.addEventListener("change", aplicar);
  $("periodoFim")?.addEventListener("change", aplicar);
  $("modoTabela")?.addEventListener("change", aplicar);

  setStatus("Pronto. Clique em Carregar.", "info");
});
