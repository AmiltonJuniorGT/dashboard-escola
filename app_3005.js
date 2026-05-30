const SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const UNITS = [
  "Números Cajazeiras",
  "Números São Cristóvão",
  "Números Camaçari"
];

const FINANCE_KEYWORDS = [
  "faturamento",
  "receita",
  "despesa",
  "custo",
  "custos",
  "lucro",
  "resultado",
  "inadimpl",
  "mensalidade",
  "matricula",
  "matrícula",
  "financeiro",
  "caixa"
];

let DATA = {};
let CURRENT = {
  unit: UNITS[0],
  indicator: "",
  start: 0,
  end: 0
};

const $ = (id) => document.getElementById(id);

boot();

async function boot() {
  document.querySelector("main").insertAdjacentHTML(
    "afterbegin",
    `<div id="loading" class="loading">Carregando dados da planilha...</div>`
  );

  try {
    for (const unit of UNITS) {
      DATA[unit] = await loadSheet(unit);
    }

    $("loading")?.remove();
    setupFilters();
    render();
  } catch (err) {
    document.querySelector("main").innerHTML =
      `<div class="error">Erro ao carregar a base: ${err.message}</div>`;
  }
}

async function loadSheet(sheetName) {
  const url =
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;

  const res = await fetch(url);
  const text = await res.text();

  const json = JSON.parse(text.substring(text.indexOf("{"), text.lastIndexOf("}") + 1));
  const cols = json.table.cols.map(c => clean(c.label || c.id || ""));
  const months = cols.slice(1).map(parseMonthLabel);

  const rows = json.table.rows.map(r => {
    const cells = r.c || [];
    const indicator = clean(getCell(cells[0]));
    const values = cells.slice(1).map(c => parseValue(getCell(c)));

    return { indicator, values };
  }).filter(r => r.indicator);

  return { months, rows };
}

function getCell(cell) {
  if (!cell) return "";
  return cell.f ?? cell.v ?? "";
}

function clean(v) {
  return String(v || "").trim();
}

function parseValue(v) {
  if (typeof v === "number") return v;

  let s = String(v || "")
    .replace(/\s/g, "")
    .replace("R$", "")
    .replace("%", "")
    .replace(/\./g, "")
    .replace(",", ".");

  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function parseMonthLabel(label) {
  const raw = clean(label);
  return {
    raw,
    label: raw.charAt(0).toUpperCase() + raw.slice(1)
  };
}

function setupFilters() {
  const unitSelect = $("unitSelect");
  const indicatorSelect = $("indicatorSelect");
  const startMonth = $("startMonth");
  const endMonth = $("endMonth");

  unitSelect.innerHTML = UNITS.map(u => `<option>${u}</option>`).join("");
  unitSelect.value = CURRENT.unit;

  fillIndicators();

  unitSelect.onchange = () => {
    CURRENT.unit = unitSelect.value;
    fillIndicators();
    render();
  };

  indicatorSelect.onchange = () => {
    CURRENT.indicator = indicatorSelect.value;
    render();
  };

  startMonth.onchange = () => {
    CURRENT.start = Number(startMonth.value);
    if (CURRENT.start > CURRENT.end) CURRENT.end = CURRENT.start;
    renderMonths();
    render();
  };

  endMonth.onchange = () => {
    CURRENT.end = Number(endMonth.value);
    if (CURRENT.end < CURRENT.start) CURRENT.start = CURRENT.end;
    renderMonths();
    render();
  };
}

function fillIndicators() {
  const rows = DATA[CURRENT.unit].rows;

  let indicators = rows
    .map(r => r.indicator)
    .filter(name => FINANCE_KEYWORDS.some(k => normalize(name).includes(k)));

  if (!indicators.length) indicators = rows.map(r => r.indicator);

  $("indicatorSelect").innerHTML = indicators
    .map(i => `<option>${i}</option>`)
    .join("");

  CURRENT.indicator = indicators[0] || "";
  $("indicatorSelect").value = CURRENT.indicator;

  const last = DATA[CURRENT.unit].months.length - 1;
  CURRENT.start = Math.max(0, last - 5);
  CURRENT.end = last;

  renderMonths();
}

function renderMonths() {
  const months = DATA[CURRENT.unit].months;

  $("startMonth").innerHTML = months.map((m, i) =>
    `<option value="${i}" ${i === CURRENT.start ? "selected" : ""}>${m.label}</option>`
  ).join("");

  $("endMonth").innerHTML = months.map((m, i) =>
    `<option value="${i}" ${i === CURRENT.end ? "selected" : ""}>${m.label}</option>`
  ).join("");
}

function render() {
  const base = DATA[CURRENT.unit];
  const row = base.rows.find(r => r.indicator === CURRENT.indicator);
  if (!row) return;

  const values = row.values;
  const slice = values.slice(CURRENT.start, CURRENT.end + 1);
  const months = base.months.slice(CURRENT.start, CURRENT.end + 1);

  const total = sum(slice);
  const average = slice.length ? total / slice.length : 0;

  const prevMonthValue = values[CURRENT.start - 1] ?? 0;
  const lastSelectedValue = values[CURRENT.end] ?? 0;

  const previousYearIndex = CURRENT.end - 12;
  const previousYearValue = previousYearIndex >= 0 ? values[previousYearIndex] : 0;

  const pctPrevMonth = percent(lastSelectedValue, prevMonthValue);
  const pctPrevYear = percent(lastSelectedValue, previousYearValue);

  renderCards({
    total,
    average,
    lastSelectedValue,
    prevMonthValue,
    previousYearValue,
    pctPrevMonth,
    pctPrevYear
  });

  drawLineChart("lineChart", months.map(m => m.label), slice);
  drawBarChart("barChart", [
    "Mês atual",
    "Mês anterior",
    "Ano anterior",
    "Média"
  ], [
    lastSelectedValue,
    prevMonthValue,
    previousYearValue,
    average
  ]);

  renderTable(months, slice);
}

function renderCards(data) {
  $("cards").innerHTML = `
    ${card("Total período", money(data.total), "Soma do período selecionado")}
    ${card("Média mensal", money(data.average), "Média dos meses selecionados")}
    ${card("Mês anterior", money(data.prevMonthValue), formatPercent(data.pctPrevMonth), data.pctPrevMonth)}
    ${card("Ano anterior", money(data.previousYearValue), formatPercent(data.pctPrevYear), data.pctPrevYear)}
  `;
}

function card(title, value, detail, pct = 0) {
  const cls = pct < 0 ? "neg" : "";
  return `
    <div class="card">
      <small>${title}</small>
      <strong>${value}</strong>
      <span class="${cls}">${detail}</span>
    </div>
  `;
}

function renderTable(months, values) {
  $("tableWrap").innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Mês</th>
          <th>Valor</th>
        </tr>
      </thead>
      <tbody>
        ${months.map((m, i) => `
          <tr>
            <td>${m.label}</td>
            <td>${money(values[i])}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function drawLineChart(id, labels, values) {
  const canvas = $(id);
  const ctx = canvas.getContext("2d");
  const w = canvas.width = canvas.offsetWidth;
  const h = canvas.height;

  ctx.clearRect(0, 0, w, h);

  const pad = 36;
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const range = max - min || 1;

  ctx.strokeStyle = "#d7eadf";
  ctx.lineWidth = 1;

  for (let i = 0; i < 5; i++) {
    const y = pad + (h - pad * 2) * i / 4;
    ctx.beginPath();
    ctx.moveTo(pad, y);
    ctx.lineTo(w - pad, y);
    ctx.stroke();
  }

  ctx.strokeStyle = "#0f7a43";
  ctx.lineWidth = 3;
  ctx.beginPath();

  values.forEach((v, i) => {
    const x = pad + ((w - pad * 2) / Math.max(values.length - 1, 1)) * i;
    const y = h - pad - ((v - min) / range) * (h - pad * 2);

    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });

  ctx.stroke();

  ctx.fillStyle = "#075c31";
  values.forEach((v, i) => {
    const x = pad + ((w - pad * 2) / Math.max(values.length - 1, 1)) * i;
    const y = h - pad - ((v - min) / range) * (h - pad * 2);

    ctx.beginPath();
    ctx.arc(x, y, 4, 0, Math.PI * 2);
    ctx.fill();

    ctx.font = "11px Arial";
    ctx.fillText(labels[i].slice(0, 8), x - 18, h - 10);
  });
}

function drawBarChart(id, labels, values) {
  const canvas = $(id);
  const ctx = canvas.getContext("2d");
  const w = canvas.width = canvas.offsetWidth;
  const h = canvas.height;

  ctx.clearRect(0, 0, w, h);

  const pad = 36;
  const max = Math.max(...values, 1);
  const barW = (w - pad * 2) / values.length - 18;

  values.forEach((v, i) => {
    const x = pad + i * ((w - pad * 2) / values.length) + 9;
    const barH = (v / max) * (h - pad * 2);
    const y = h - pad - barH;

    ctx.fillStyle = "#0f7a43";
    ctx.fillRect(x, y, barW, barH);

    ctx.fillStyle = "#18352a";
    ctx.font = "11px Arial";
    ctx.fillText(labels[i], x - 4, h - 10);
  });
}

function sum(arr) {
  return arr.reduce((a, b) => a + Number(b || 0), 0);
}

function percent(current, previous) {
  if (!previous) return 0;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function formatPercent(v) {
  const sign = v > 0 ? "+" : "";
  return `${sign}${v.toFixed(1).replace(".", ",")}%`;
}

function money(v) {
  return Number(v || 0).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL"
  });
}

function normalize(v) {
  return String(v || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}
