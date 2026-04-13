/* visualisation_flux.js — Superposition multi-années PULSE */
"use strict";

/* =========================================================
   PALETTE ANNÉES
   ========================================================= */
const YEAR_COLORS = [
  "#4C7CF3", // bleu primaire
  "#32D583", // vert
  "#F59E0B", // ambre
  "#EC4899", // rose
  "#8B5CF6", // violet
  "#06B6D4", // cyan
  "#F97316", // orange
  "#10B981", // émeraude
  "#EF4444", // rouge
  "#A78BFA", // lavande
];

const MONTHS_LABELS = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"];

/* =========================================================
   STATE
   ========================================================= */
let catalogue    = {};     // { section: [flux_names] }
let currentData  = null;   // réponse brute de l'API
let activeYears  = new Set();
let chartType    = { weekly: "line", monthly: "line", annual: "bar" };
let charts       = { weekly: null, monthly: null, annual: null };

/* =========================================================
   HELPERS
   ========================================================= */
function qs(id) { return document.getElementById(id); }

function formatNum(v) {
  if (typeof v !== "number" || Number.isNaN(v)) return "—";
  return v.toLocaleString("fr-FR");
}

function yearColor(annee, allYears) {
  const idx = allYears.indexOf(annee);
  return YEAR_COLORS[idx % YEAR_COLORS.length];
}

/* =========================================================
   CATALOGUE
   ========================================================= */
async function loadCatalogue() {
  const res  = await fetch("/api/catalogue");
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  catalogue  = await res.json();

  const sel = qs("f-section");
  sel.innerHTML = '<option value="">Choisir…</option>';
  Object.keys(catalogue).sort().forEach(s => {
    const opt   = document.createElement("option");
    opt.value   = s;
    opt.textContent = s;
    sel.appendChild(opt);
  });
}

/* =========================================================
   MISE À JOUR SELECT FLUX
   ========================================================= */
function updateFluxSelect(section) {
  const sel  = qs("f-flux");
  sel.innerHTML = "";
  sel.disabled  = !section;

  if (!section) {
    sel.innerHTML = '<option value="">— sélectionner une section —</option>';
    return;
  }

  const fluxList = catalogue[section] || [];

  // Option "Tous les flux" en premier
  const all = document.createElement("option");
  all.value = "Tous les flux";
  all.textContent = "Tous les flux";
  sel.appendChild(all);

  fluxList.sort().forEach(f => {
    const opt     = document.createElement("option");
    opt.value     = f;
    opt.textContent = f;
    sel.appendChild(opt);
  });

  qs("btn-afficher").disabled = false;
}

/* =========================================================
   YEAR TOGGLES
   ========================================================= */
function buildYearToggles(annees) {
  const container = qs("year-toggles");
  container.innerHTML = "";

  if (!annees || annees.length === 0) {
    container.innerHTML = '<span style="font-size:11px;color:#475569;">Aucune année disponible</span>';
    return;
  }

  annees.forEach(y => {
    const color = yearColor(y, annees);
    const btn   = document.createElement("button");
    btn.className     = "year-btn";
    btn.dataset.year  = y;
    btn.textContent   = y;
    btn.style.cssText = `color:${color};border-color:${color}40;background:${color}18;`;

    if (activeYears.has(y)) {
      btn.style.cssText = `color:${color};border-color:${color};background:${color}28;box-shadow:0 0 8px ${color}40;`;
    } else {
      btn.classList.add("off");
    }

    btn.addEventListener("click", () => toggleYear(y, annees));
    container.appendChild(btn);
  });
}

function toggleYear(y, allYears) {
  if (activeYears.has(y)) {
    activeYears.delete(y);
  } else {
    activeYears.add(y);
  }
  buildYearToggles(allYears);
  if (currentData) renderCharts();
}

function setAllYears(annees, active) {
  activeYears.clear();
  if (active) annees.forEach(y => activeYears.add(y));
  buildYearToggles(annees);
  if (currentData) renderCharts();
}

/* =========================================================
   CHART.JS DEFAULTS
   ========================================================= */
Chart.defaults.color           = "#FFFFFF";
Chart.defaults.borderColor     = "rgba(255,255,255,.07)";
Chart.defaults.font.family     = "Inter, system-ui, sans-serif";

function commonOptions(ticksCallback) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: {
      legend: {
        position: "top",
        labels: { boxWidth: 12, padding: 14, font: { size: 11 }, color: "#94A3B8" }
      },
      tooltip: {
        backgroundColor: "#0e1420",
        borderColor: "rgba(255,255,255,.12)",
        borderWidth: 1,
        titleColor: "#F3F4F6",
        bodyColor: "#94A3B8",
        padding: 10,
        callbacks: {
          label: ctx => ` ${ctx.dataset.label} : ${formatNum(ctx.parsed.y)}`
        }
      }
    },
    scales: {
      x: {
        grid: { color: "rgba(255,255,255,.05)" },
        ticks: { color: "#FFFFFF", font: { size: 10, weight: "500" }, callback: ticksCallback },
        title: { display: true, text: "Période", color: "#FFFFFF", font: { size: 11, weight: "500" } }
      },
      y: {
        grid: { color: "rgba(255,255,255,.05)" },
        ticks: { color: "#FFFFFF", font: { size: 10, weight: "500" }, callback: v => formatNum(v) },
        title: { display: true, text: "Montant (k€)", color: "#FFFFFF", font: { size: 11, weight: "500" } }
      }
    }
  };
}

/* =========================================================
   CONSTRUCTION DATASETS
   ========================================================= */
function buildWeeklyDatasets(data, visibleYears) {
  // All ISO weeks present across any year
  const allWeeks = new Set();
  visibleYears.forEach(y => {
    Object.keys(data.weekly[y] || {}).forEach(w => allWeeks.add(Number(w)));
  });
  const weeks = Array.from(allWeeks).sort((a, b) => a - b);

  const datasets = visibleYears.map(y => {
    const color  = yearColor(y, data.annees);
    const values = weeks.map(w => data.weekly[y]?.[String(w)] ?? null);
    return {
      label:           String(y),
      data:            values,
      borderColor:     color,
      backgroundColor: color + "28",
      borderWidth:     2,
      pointRadius:     2,
      tension:         0.35,
    };
  });
  return { labels: weeks.map(w => `S${w}`), datasets };
}

function buildMonthlyDatasets(data, visibleYears) {
  // Month labels filtered to selected period
  const monthStart = parseInt(qs("f-mois-debut").value, 10);
  const monthEnd   = parseInt(qs("f-mois-fin").value, 10);
  const months     = [];
  for (let m = monthStart; m <= monthEnd; m++) months.push(m);

  const datasets = visibleYears.map(y => {
    const color  = yearColor(y, data.annees);
    const values = months.map(m => data.monthly[y]?.[String(m)] ?? null);
    return {
      label:           String(y),
      data:            values,
      borderColor:     color,
      backgroundColor: color + "28",
      borderWidth:     2,
      pointRadius:     3,
      tension:         0.3,
    };
  });
  return { labels: months.map(m => MONTHS_LABELS[m - 1]), datasets };
}

function buildAnnualDatasets(data, visibleYears) {
  const color = visibleYears.map(y => yearColor(y, data.annees));
  const values = visibleYears.map(y => data.annual[String(y)] ?? 0);
  return {
    labels: visibleYears.map(String),
    datasets: [{
      label:           "Total annuel",
      data:            values,
      backgroundColor: color.map(c => c + "55"),
      borderColor:     color,
      borderWidth:     2,
      borderRadius:    6,
    }]
  };
}

/* =========================================================
   RENDU CHARTS
   ========================================================= */
function destroyChart(key) {
  if (charts[key]) { charts[key].destroy(); charts[key] = null; }
}

function renderCharts() {
  if (!currentData) return;

  const data         = currentData;
  const visibleYears = data.annees.filter(y => activeYears.has(y));

  if (visibleYears.length === 0) {
    // Détruire tous les charts, afficher message dans la topbar
    ["weekly","monthly","annual"].forEach(destroyChart);
    return;
  }

  // ── Hebdomadaire ──
  {
    const { labels, datasets } = buildWeeklyDatasets(data, visibleYears);
    const type = chartType.weekly;
    const opts = commonOptions(null);
    // Pour bar, ne pas afficher tous les ticks si trop nombreux
    if (type === "bar") opts.scales.x.ticks.maxTicksLimit = 26;

    destroyChart("weekly");
    charts.weekly = new Chart(qs("chart-weekly"), {
      type,
      data: { labels, datasets },
      options: opts,
    });
  }

  // ── Mensuel ──
  {
    const { labels, datasets } = buildMonthlyDatasets(data, visibleYears);
    const type = chartType.monthly;
    destroyChart("monthly");
    charts.monthly = new Chart(qs("chart-monthly"), {
      type,
      data: { labels, datasets },
      options: commonOptions(null),
    });
  }

  // ── Annuel ──
  {
    const { labels, datasets } = buildAnnualDatasets(data, visibleYears);
    const type = chartType.annual;
    destroyChart("annual");
    charts.annual = new Chart(qs("chart-annual"), {
      type,
      data: { labels, datasets },
      options: {
        ...commonOptions(null),
        plugins: {
          ...commonOptions(null).plugins,
          legend: { display: false },
        }
      },
    });
  }
}

/* =========================================================
   KPI ROW
   ========================================================= */
function renderKpis(data) {
  const kpis = data.kpis;
  qs("kpi-nb-annees").textContent  = kpis.nb_annees;
  qs("kpi-peak-annee").textContent = kpis.annee_peak.annee;
  qs("kpi-peak-val").textContent   = formatNum(kpis.annee_peak.valeur);
  qs("kpi-trough-annee").textContent = kpis.annee_trough.annee;
  qs("kpi-trough-val").textContent   = formatNum(kpis.annee_trough.valeur);
  qs("kpi-nb-points").textContent  = formatNum(kpis.nb_points);
  qs("kpi-row").style.display = "grid";
}

/* =========================================================
   TOPBAR
   ========================================================= */
function updateTopbar(data) {
  qs("chart-title").textContent    = data.flux === "Tous les flux"
    ? `Tous les flux — ${data.section}`
    : `${data.flux}`;
  qs("chart-subtitle").textContent = `${data.section} · mois ${data.month_start}→${data.month_end}`;

  qs("chart-badges").innerHTML = `
    <span class="chart-badge chart-badge--section">${data.section}</span>
    <span class="chart-badge chart-badge--flux">${data.flux}</span>
  `;
}

/* =========================================================
   ÉTATS UI
   ========================================================= */
function showState(state) {
  ["state-empty","state-loading","state-error","state-result"].forEach(id => {
    const el = qs(id);
    if (el) el.hidden = true;
  });
  const target = qs(state);
  if (target) target.hidden = false;
}

/* =========================================================
   FETCH & RENDER
   ========================================================= */
async function fetchAndRender() {
  const section    = qs("f-section").value.trim();
  const flux       = qs("f-flux").value.trim();
  const monthStart = qs("f-mois-debut").value;
  const monthEnd   = qs("f-mois-fin").value;

  if (!section || !flux) return;

  showState("state-loading");
  qs("kpi-row").style.display = "none";

  const params = new URLSearchParams({
    section,
    flux,
    month_start: monthStart,
    month_end:   monthEnd,
  });

  try {
    const res = await fetch(`/api/visualisation_flux?${params}`);
    if (!res.ok) {
      const json = await res.json().catch(() => ({}));
      throw new Error(json.error || `HTTP ${res.status}`);
    }

    currentData = await res.json();

    // Normalise annees to numbers
    currentData.annees = (currentData.annees || []).map(Number);

    // Par défaut, toutes les années actives
    activeYears.clear();
    currentData.annees.forEach(y => activeYears.add(y));

    buildYearToggles(currentData.annees);
    updateTopbar(currentData);
    renderKpis(currentData);

    showState("state-result");
    renderCharts();

    const btnExcel = document.getElementById("btn-export-excel");
    if (btnExcel) btnExcel.disabled = false;

  } catch (err) {
    qs("error-msg").textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
}

/* =========================================================
   CHART TYPE TOGGLE (ligne/barres par chart)
   ========================================================= */
document.querySelectorAll(".ctype-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    const chartKey = btn.dataset.chart;
    const type     = btn.dataset.type;

    // Désactiver les autres boutons du même chart
    document.querySelectorAll(`.ctype-btn[data-chart="${chartKey}"]`).forEach(b => {
      b.classList.remove("active");
    });
    btn.classList.add("active");

    chartType[chartKey] = type;
    if (currentData) renderCharts();
  });
});

/* =========================================================
   EVENTS
   ========================================================= */
qs("f-section").addEventListener("change", () => {
  const section = qs("f-section").value;
  updateFluxSelect(section);
  currentData = null;
  activeYears.clear();
  qs("year-toggles").innerHTML = '<span style="font-size:11px;color:#475569;">Chargez des données…</span>';
  qs("kpi-row").style.display = "none";
  showState("state-empty");
  qs("chart-badges").innerHTML = "";
  qs("chart-title").textContent    = "Superposition multi-années";
  qs("chart-subtitle").textContent = "Sélectionnez une filiale et un flux";
});

qs("f-flux").addEventListener("change", () => {
  qs("btn-afficher").disabled = !qs("f-flux").value;
});

qs("btn-afficher").addEventListener("click", fetchAndRender);

document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
  window.pulsePDF("Superposition multi-années — PULSE");
});

document.getElementById("btn-export-excel")?.addEventListener("click", () => {
  const chart = charts.monthly || charts.weekly || charts.annual;
  if (chart) {
    const section = qs("f-section")?.value || "Section";
    const flux    = qs("f-flux")?.value    || "Flux";
    window.pulseExcelChart(chart, `superposition_${section}_${flux}`);
  } else {
    alert("Affichez d'abord un graphique.");
  }
});

qs("btn-year-all").addEventListener("click", () => {
  if (currentData) setAllYears(currentData.annees, true);
});

qs("btn-year-none").addEventListener("click", () => {
  if (currentData) setAllYears(currentData.annees, false);
});

/* =========================================================
   INIT
   ========================================================= */
(async () => {
  try {
    await loadCatalogue();
  } catch (err) {
    window.toast?.("Impossible de charger le catalogue : " + err.message, "error");
  }
})();
