// ================= CONFIG =================
const DATA_SHEET_ID = "1d4G--uvR-fjdn4gP8HM7r69SCHG_6bZNBpe_97Zx3Go";

const TABS_UNIDADES = {
  Cajazeiras: "Números Cajazeiras",
  Camacari: "Números Camaçari",
  SaoCristovao: "Números Sao Cristovao",
};

const state = {
  unidade: "TODAS",
  monthKey: null
};

const $ = (id) => document.getElementById(id);

// ================= HELPERS =================
function stripAccents(s) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function parseMonthLabel(label) {
  if (!label) return null;

  const m = label.toLowerCase().match(/([a-zç]{3,})\.?\s*(\d{2})/i);
  if (!m) return null;

  const monthMap = {
    jan:1, fev:2, mar:3, abr:4, mai:5, jun:6,
    jul:7, ago:8, set:9, out:10, nov:11, dez:12
  };

  const month = monthMap[stripAccents(m[1].slice(0,3))];
