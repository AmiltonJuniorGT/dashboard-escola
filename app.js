/* =========================================================
   Dashboard – Fechamento (3 Unidades) – versão estável
   (meses em colunas e cabeçalho em qualquer linha)

   Fonte: Planilha "Plano de Ação Co Gestão"
   ID: 1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go

   Estrutura esperada nas abas (unidades):
   - Coluna A: nome do índice/indicador
   - Alguma linha (não necessariamente a 1ª) contém os meses nas colunas (ex: janeiro.23)
   - Abas = unidades (Cajazeiras / Camaçari / São Cristóvão)

   Esta versão:
   - encontra automaticamente a linha que contém os meses
   - aceita: "janeiro.23", "jan.23", "jan/23", "janeiro 2023" etc
   - não quebra se faltar algum mês (M-1/M-2/M-3/Ano-1)
   - separa blocos "TÉCNICO" e "PROFISSIONALIZANTE"
   ========================================================= */

(() => {
  // ================= CONFIG =================
  const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

  // chaves SEM acento/espaço (evita erros de sintaxe e de edição)
  const TABS_UNIDADES = {
    Cajazeiras: "Números Cajazeiras",
    Camacari: "Números Camaçari",
    SaoCristovao: "Números Sao Cristovao",
  };

  const UNIDADE_LABEL = {
    Cajazeiras: "Cajazeiras",
    Camacari: "Camaçari",
    SaoCristovao: "São Cristóvão",
  };

  const state = {
    unidade: "TODAS", // TODAS | Cajazeiras | Camacari | SaoCristovao
    monthKey: null,  // {y: 2025, m: 12}
  };

  const $ = (id) => document.getElementById(id);

  // ================= Helpers =================
  function stripAccents(s) {
    return String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }
  function norm(s) {
    return stripAccents(String(s ?? "")).trim().toUpperCase();
  }

  // BR parse (R$ / milhar / decimal / %)
  function parseBR(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === "number") return v;
    const s0 = String(v).trim();
    if (!s0) return null;

    const s = s0
      .replace(/R\$\s?/i, "")
      .replace(/\s/g, "")
      .replace(/\./g, "")
      .replace(",", ".")
      .replace("%", "");

    const n = Number(s);
    return Number.isNaN(n) ? null : n;
  }

  function fmtCurrency(n) {
    if (n === null || n === undefined) return "—";
    return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  }
  function fmtInt(n) {
    if (n === null || n === undefined) return "—";
    return Math.round(n).toLocaleString("pt-BR");
  }
  function fmtPct(n) {
    if (n === null || n === undefined) return "—";
    return `${n.toLocaleString("pt-BR", { maximumFractionDigits: 2 })}%`;
  }
  function fmtValue(n, type) {
    if (type === "currency") return fmtCurrency(n);
    if (type === "int") return fmtInt(n);
    if (type === "pct") return fmtPct(n);
    return n ?? "—";
  }
  function fmtDelta(n, type) {
    if (n === null || n === undefined) return "—";
    const up = n >= 0;
    const arrow = up ? "↑" : "↓";
    const cls = up ? "delta up" : "delta down";
    return `<span class="${cls}">${arrow} ${fmtValue(n, type)}</span>`;
  }

  // ================= GViz fetch =================
  async function fetchGviz(sheetId, tab) {
    // headers=0 => NÃO assume cabeçalho na 1ª linha (sua planilha tem título acima)
    const url =
      `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?sheet=${encodeURIComponent(tab)}&headers=0`;

    const res = await fetch(url);
    const txt = await res.text();
    const json = JSON.parse(txt.substring(txt.indexOf("{"), txt.lastIndexOf("}") + 1));
    const rows = json.table.rows.map((r) => r.c.map((cell) => (cell ? cell.v : null)));
    return { rows };
  }

  // ================= Meses (colunas) =================
  const PT_MONTHS = {
    JANEIRO: 1, JAN: 1,
    FEVEREIRO: 2, FEV: 2,
    MARCO: 3, MARÇO: 3, MAR: 3,
    ABRIL: 4, ABR: 4,
    MAIO: 5, MAI: 5,
    JUNHO: 6, JUN: 6,
    JULHO: 7, JUL: 7,
    AGOSTO: 8, AGO: 8,
    SETEMBRO: 9, SET: 9,
    OUTUBRO: 10, OUT: 10,
    NOVEMBRO: 11, NOV: 11,
    DEZEMBRO: 12, DEZ: 12,
  };

  function parseMonthLabel(label) {
    const raw = String(label ?? "").trim();
    if (!raw) return null;

    const cleaned = stripAccents(raw).toUpperCase();
    // aceita: "janeiro.23", "janeiro 23", "janeiro/23", "jan.23", "jan/23", "jan 23", "janeiro.2023"
    const m = cleaned.match(/^([A-ZÇ]+)[\.\s\/-]?(\d{2}|\d{4})/);
    if (!m) return null;

    const name = m[1];
    const yy = m[2];

    const month = PT_MONTHS[name];
    if (!month) return null;

    const year = yy.length === 4 ? Number(yy) : Number("20" + yy);
    if (!year || year < 2000 || year > 2100) return null;

    return { y: year, m: month };
  }

  function mkToKeyNum(mk) { return mk.y * 100 + mk.m; }

  function mkToShortLabel(mk) {
    const names = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    return `${names[mk.m]}.${String(mk.y).slice(-2)}`;
  }

  function shiftMonth(mk, delta) {
    let y = mk.y;
    let m = mk.m + delta;
    while (m <= 0) { m += 12; y -= 1; }
    while (m > 12) { m -= 12; y += 1; }
    return { y, m };
  }

  // Encontra a "linha de meses" (onde várias colunas parecem mês)
  function findMonthHeaderRow(rows) {
    let best = { idx: -1, hits: 0, header: null };

    for (let i = 0; i < Math.min(rows.length, 30); i++) {
      const r = rows[i] || [];
      let hits = 0;
      for (let j = 0; j < r.length; j++) {
        if (parseMonthLabel(r[j])) hits++;
      }
      if (hits > best.hits) best = { idx: i, hits, header: r };
    }

    // precisa ter pelo menos 2 meses para considerar cabeçalho válido
    if (best.idx >= 0 && best.hits >= 2) return best;
    return null;
  }

  function buildMonthColsFromHeaderRow(headerRow) {
    const monthCols = [];
    for (let j = 0; j < headerRow.length; j++) {
      const mk = parseMonthLabel(headerRow[j]);
      if (mk) monthCols.push({ idx: j, mk, keyNum: mkToKeyNum(mk), raw: String(headerRow[j]) });
    }
    monthCols.sort((a, b) => a.keyNum - b.keyNum);
    const byKey = new Map(monthCols.map((c) => [c.keyNum, c.idx]));
    return { monthCols, byKey };
  }

  // ================= Seções =================
  function findSectionRange(rows, needle) {
    const n = norm(needle);
    let start = -1;

    for (let i = 0; i < rows.length; i++) {
      const a = norm(rows[i]?.[0]);
      if (a.includes(n)) { start = i + 1; break; }
    }
    if (start < 0) return null;

    let end = rows.length;
    for (let i = start; i < rows.length; i++) {
      const a = norm(rows[i]?.[0]);
      if (a.includes("PLANILHA DE INDICADORES") && i > start) { end = i; break; }
      if (a.includes("FACULDADE")) { end = i; break; }
      if (a.includes("GRAU EDUCACIONAL") && i > start) { end = i; break; }
    }
    return { start, end };
  }

  function findRowIndexInRange(rows, range, needles) {
    const ns = needles.map(norm);
    for (let i = range.start; i < range.end; i++) {
      const a = norm(rows[i]?.[0]);
      if (!a) continue;
      for (const nn of ns) {
        if (a === nn || a.includes(nn)) return i;
      }
    }
    return -1;
  }

  // ================= Métricas =================
  const METRICS = [
    { label: "Faturamento", type: "currency", needles: ["FATURAMENTO"] },
    { label: "Custo Operacional", type: "currency", needles: ["CUSTO OPERACIONAL"] },
    { label: "Custo Total", type: "currency", needles: ["CUSTO TOTAL"] },
    { label: "Ticket Médio", type: "currency", needles: ["TICKET MEDIO", "TICKET MÉDIO"] },
    { label: "Ativos", type: "int", needles: ["ALUNOS ATIVOS", "ATIVOS", "NUMERO DE ATIVOS", "NÚMERO DE ATIVOS"] },
    { label: "Matrículas", type: "int", needles: ["MATRICULAS REALIZADAS", "MATRÍCULAS REALIZADAS", "MATRICULAS"] },
    { label: "Meta", type: "int", needles: ["META DE MATRICULA", "META DE MATRÍCULA", "META"] },
    { label: "Evasão Real", type: "int", needles: ["EVASAO REAL", "EVASÃO REAL"] },
    { label: "Inadimplência (%)", type: "pct", needles: ["INADIMPLENCIA", "INADIMPLÊNCIA"] },
  ];

  function getMetric(rows, range, metric, colIdx) {
    const rowIdx = findRowIndexInRange(rows, range, metric.needles);
    if (rowIdx < 0) return null;
    return parseBR(rows[rowIdx]?.[colIdx]);
  }

  function buildComparative(rows, range, byKey, mkCur) {
    const mkM1 = shiftMonth(mkCur, -1);
    const mkM2 = shiftMonth(mkCur, -2);
    const mkM3 = shiftMonth(mkCur, -3);
    const mkY1 = shiftMonth(mkCur, -12);

    const idxCur = byKey.get(mkToKeyNum(mkCur));
    const idxM1  = byKey.get(mkToKeyNum(mkM1));
    const idxM2  = byKey.get(mkToKeyNum(mkM2));
    const idxM3  = byKey.get(mkToKeyNum(mkM3));
    const idxY1  = byKey.get(mkToKeyNum(mkY1));

    return METRICS.map((m) => {
      const cur = idxCur != null ? getMetric(rows, range, m, idxCur) : null;
      const m1  = idxM1  != null ? getMetric(rows, range, m, idxM1)  : null;
      const m2  = idxM2  != null ? getMetric(rows, range, m, idxM2)  : null;
      const m3  = idxM3  != null ? getMetric(rows, range, m, idxM3)  : null;
      const y1  = idxY1  != null ? getMetric(rows, range, m, idxY1)  : null;

      const vm = (cur != null && m1 != null) ? (cur - m1) : null;
      const vy = (cur != null && y1 != null) ? (cur - y1) : null;

      return { metric: m, values: { cur, vm, vy, m1, m2, m3, y1 } };
    });
  }

  // ================= UI (inline styles) =================
  function ensureInlineStyles() {
    if (document.getElementById("fechamento-style")) return;
    const st = document.createElement("style");
    st.id = "fechamento-style";
    st.textContent = `
      .fech-topbar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:10px 0 14px 0}
      .fech-topbar select,.fech-topbar button{padding:6px 10px;border-radius:10px;border:1px solid #cfd8cf;background:#fff}
      .fech-topbar button{cursor:pointer}
      .fech-card{background:#fff;border:1px solid #dfe7df;border-radius:14px;padding:12px 14px;margin:12px 0}
      .fech-muted{opacity:.75;font-size:12px;margin-top:6px}
      .fech-table-wrap{overflow:auto}
      table.fech{width:100%;border-collapse:collapse}
      table.fech th, table.fech td{border-bottom:1px solid #e9efe9;padding:8px 10px;text-align:right;white-space:nowrap}
      table.fech th:first-child, table.fech td:first-child{text-align:left}
      .delta{font-weight:700}
      .delta.up{color:#1b7f3a}
      .delta.down{color:#b3261e}
      .block-title{font-weight:800;margin:0 0 8px 0}
      .fech-error{background:#fff3f3;border:1px solid #ffd1d1;color:#8a1f1f;border-radius:14px;padding:12px 14px;margin:12px 0}
    `;
    document.head.appendChild(st);
  }

  function renderBlock(title, rowsBuilt, mkCur) {
    const head = `
      <thead>
        <tr>
          <th>Indicador</th>
          <th>${mkToShortLabel(mkCur)}</th>
          <th>Δ M-1</th>
          <th>Δ Ano-1</th>
          <th>M-1</th>
          <th>M-2</th>
          <th>M-3</th>
          <th>Ano-1</th>
        </tr>
      </thead>
    `;

    const body = `
      <tbody>
        ${rowsBuilt.map((r) => {
          const t = r.metric.type;
          const v = r.values;
          return `
            <tr>
              <td>${r.metric.label}</td>
              <td>${fmtValue(v.cur, t)}</td>
              <td>${fmtDelta(v.vm, t)}</td>
              <td>${fmtDelta(v.vy, t)}</td>
              <td>${fmtValue(v.m1, t)}</td>
              <td>${fmtValue(v.m2, t)}</td>
              <td>${fmtValue(v.m3, t)}</td>
              <td>${fmtValue(v.y1, t)}</td>
            </tr>
          `;
        }).join("")}
      </tbody>
    `;

    return `
      <div class="fech-card">
        <div class="block-title">${title}</div>
        <div class="fech-table-wrap">
          <table class="fech">
            ${head}
            ${body}
          </table>
        </div>
      </div>
    `;
  }

  // ================= RENDER =================
  async function resolveMonthOptions() {
    // usa Cajazeiras apenas para descobrir todos os meses disponíveis
    const { rows } = await fetchGviz(DATA_SHEET_ID, TABS_UNIDADES.Cajazeiras);
    const hdr = findMonthHeaderRow(rows);
    if (!hdr) return { monthCols: [], byKey: new Map() };

    const { monthCols, byKey } = buildMonthColsFromHeaderRow(hdr.header);
    return { monthCols, byKey };
  }

  async function renderFechamento() {
    const container = $("fechamentoView");
    if (!container) return;

    ensureInlineStyles();
    container.innerHTML = "Carregando…";

    const { monthCols } = await resolveMonthOptions();
    if (!monthCols.length) {
      container.innerHTML = `<div class="fech-error">Não encontrei colunas de mês (ex: janeiro.23). Confirme se os meses estão em alguma linha (não no topo do arquivo).</div>`;
      return;
    }

    if (!state.monthKey) {
      state.monthKey = monthCols[monthCols.length - 1].mk; // mais recente
    }
    const mkCur = state.monthKey;

    const monthOptions = monthCols.map((c) => {
      const selected = (mkToKeyNum(mkCur) === c.keyNum) ? "selected" : "";
      return `<option value="${c.keyNum}" ${selected}>${mkToShortLabel(c.mk)} (${c.raw})</option>`;
    }).join("");

    const unitOptions = [
      `<option value="TODAS" ${state.unidade==="TODAS"?"selected":""}>Todas</option>`,
      ...Object.keys(TABS_UNIDADES).map((k) => {
        const sel = state.unidade === k ? "selected" : "";
        return `<option value="${k}" ${sel}>${UNIDADE_LABEL[k]}</option>`;
      })
    ].join("");

    const topbar = `
      <div class="fech-topbar">
        <b>Unidade:</b>
        <select id="fech_unidade">${unitOptions}</select>

        <b>Mês:</b>
        <select id="fech_mes">${monthOptions}</select>

        <button id="fech_aplicar">Aplicar</button>
        <button id="fech_ultimo">Mês mais recente</button>
      </div>
    `;

    const keys = state.unidade === "TODAS"
      ? Object.keys(TABS_UNIDADES)
      : [state.unidade];

    let content = "";
    for (const k of keys) {
      const tab = TABS_UNIDADES[k];
      const label = UNIDADE_LABEL[k] ?? k;

      const { rows } = await fetchGviz(DATA_SHEET_ID, tab);

      const hdr = findMonthHeaderRow(rows);
      if (!hdr) {
        content += `<div class="fech-error"><b>${label}</b>: não consegui localizar a linha com os meses (ex: janeiro.23).</div>`;
        continue;
      }

      const { byKey } = buildMonthColsFromHeaderRow(hdr.header);

      const rgTec = findSectionRange(rows, "TECNICO");
      const rgPro = findSectionRange(rows, "PROFISSIONALIZANTE");

      content += `
        <div class="fech-card">
          <div style="font-weight:900;font-size:18px">${label}</div>
          <div class="fech-muted">Aba: ${tab} • Mês: ${mkToShortLabel(mkCur)}</div>
        </div>
      `;

      if (!rgTec && !rgPro) {
        content += `<div class="fech-error">Não encontrei os blocos TÉCNICO/PROFISSIONALIZANTE nessa aba.</div>`;
        continue;
      }

      if (rgTec) content += renderBlock("Grau Técnico", buildComparative(rows, rgTec, byKey, mkCur), mkCur);
      else content += `<div class="fech-error">Bloco TÉCNICO não encontrado.</div>`;

      if (rgPro) content += renderBlock("Profissionalizante", buildComparative(rows, rgPro, byKey, mkCur), mkCur);
      else content += `<div class="fech-error">Bloco PROFISSIONALIZANTE não encontrado.</div>`;
    }

    container.innerHTML = topbar + content;

    const selU = $("fech_unidade");
    const selM = $("fech_mes");
    const btnA = $("fech_aplicar");
    const btnLast = $("fech_ultimo");

    if (btnA) btnA.onclick = async () => {
      state.unidade = selU ? selU.value : state.unidade;
      if (selM) {
        const keyNum = Number(selM.value);
        const y = Math.floor(keyNum / 100);
        const m = keyNum % 100;
        state.monthKey = { y, m };
      }
      await renderFechamento();
    };

    if (btnLast) btnLast.onclick = async () => {
      state.unidade = selU ? selU.value : state.unidade;
      const { monthCols: monthCols2 } = await resolveMonthOptions();
      state.monthKey = monthCols2.length ? monthCols2[monthCols2.length - 1].mk : state.monthKey;
      await renderFechamento();
    };
  }

  window.addEventListener("load", () => {
    renderFechamento().catch((err) => {
      console.error(err);
      const c = $("fechamentoView");
      if (c) c.innerHTML = `<div class="fech-error">Erro ao carregar: ${String(err)}</div>`;
    });
  });
})();
