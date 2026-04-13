/* tendance.js — Analyse avancée des flux (version web fidèle) */
"use strict";

/* =========================================================
 * CONFIG GLOBALE
 * ========================================================= */
Chart.defaults.color = "#8896B0";
Chart.defaults.borderColor = "#1A1F2E";
Chart.defaults.font.family = "Inter, system-ui, sans-serif";
Chart.defaults.font.size = 12;

const UI = {
  bg: "#0B0F17",
  surface: "#141A24",
  surface2: "#1A2230",
  surface3: "#212B3A",
  border: "#2B3647",
  borderSoft: "#212A38",
  text: "#F3F4F6",
  textSoft: "#D1D5DB",
  muted: "#9CA3AF",
  muted2: "#7C8798",
  accent: "#4C7CF3",
  yellow: "#FFCC00",
  green: "#27AE60",
  amber: "#D68910",
  red: "#C0392B",
  blue: "#3498DB",
  white: "#FFFFFF",
};

let catalogue = {};
let chartRegistry = [];

/* =========================================================
 * HELPERS FORMAT
 * ========================================================= */
const fmt = (v, digits = 2) =>
  typeof v === "number" && Number.isFinite(v)
    ? v.toLocaleString("fr-FR", {
        minimumFractionDigits: 0,
        maximumFractionDigits: digits,
      })
    : "—";

const fmtPct = (v, digits = 1) =>
  typeof v === "number" && Number.isFinite(v) ? `${fmt(v, digits)} %` : "—";

const fmtRange = (a, b, digits = 0) =>
  `[${fmt(a, digits)} ; ${fmt(b, digits)}]`;

const getOr = (obj, key, fallback = null) =>
  obj && Object.prototype.hasOwnProperty.call(obj, key) ? obj[key] : fallback;

/* =========================================================
 * ÉTATS
 * ========================================================= */
function showState(id) {
  ["state-empty", "state-loading", "state-error", "state-result"].forEach((s) => {
    const el = document.getElementById(s);
    if (el) el.hidden = s !== id;
  });
}

function destroyCharts() {
  for (const chart of chartRegistry) {
    try {
      chart.destroy();
    } catch (_) {}
  }
  chartRegistry = [];
}

/* =========================================================
 * CATALOGUE / FILTRES
 * ========================================================= */
async function loadCatalogue() {
  const res = await fetch("/api/catalogue");
  if (!res.ok) throw new Error(`Catalogue HTTP ${res.status}`);
  catalogue = await res.json();

  const selSection = document.getElementById("f-section");
  const selAnnee = document.getElementById("f-annee");

  selSection.innerHTML = `<option value="">Choisir…</option>`;
  for (const section of Object.keys(catalogue).sort()) {
    const opt = document.createElement("option");
    opt.value = section;
    opt.textContent = section;
    selSection.appendChild(opt);
  }

  try {
    const metaRes = await fetch("/api/accueil");
    if (metaRes.ok) {
      const meta = await metaRes.json();
      const annees = (meta.annees || []).slice().sort((a, b) => a - b);

      selAnnee.innerHTML = `<option value="">Toutes</option>`;
      for (const y of annees) {
        const opt = document.createElement("option");
        opt.value = String(y);
        opt.textContent = String(y);
        selAnnee.appendChild(opt);
      }

      if (annees.length) {
        const currentYear = new Date().getFullYear();
        const best =
          annees.filter((y) => y < currentYear).pop() ?? annees[annees.length - 1];
        selAnnee.value = String(best);
      }
    }
  } catch (_) {
    /* fallback silencieux */
  }

  selSection.addEventListener("change", onSectionChange);
  document.getElementById("f-flux").addEventListener("change", checkReady);
  document.getElementById("btn-analyser").addEventListener("click", analyser);

  /* Export PDF */
  document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
    window.pulsePDF("Tendance flux — PULSE");
  });

  /* Export Excel : exporte les données de la première série (dates + réels) */
  document.getElementById("btn-export-excel")?.addEventListener("click", () => {
    const btnExcel = document.getElementById("btn-export-excel");
    const body = btnExcel?._tendanceBody;
    if (!body) { alert("Lancez d'abord une analyse."); return; }

    const section = document.getElementById("f-section")?.value || "Section";
    const flux    = document.getElementById("f-flux")?.value    || "Flux";

    /* Utiliser le premier graphique de chartRegistry si disponible */
    if (chartRegistry.length > 0) {
      window.pulseExcelChart(chartRegistry[0], `tendance_${section}_${flux}`);
    } else if (body.dates && body.reel) {
      const headers = ["Date", "Réel"];
      const rows = body.dates.map((d, i) => [d, body.reel[i] ?? ""]);
      window.pulseExcelData(headers, rows, `tendance_${section}_${flux}`);
    } else {
      alert("Aucune donnée à exporter.");
    }
  });
}

function onSectionChange() {
  const section = document.getElementById("f-section").value;
  const selFlux = document.getElementById("f-flux");

  selFlux.innerHTML = "";
  selFlux.disabled = !section;

  if (!section) {
    selFlux.innerHTML = "<option value=''>— sélectionner une section d'abord —</option>";
    checkReady();
    return;
  }

  const flux = catalogue[section] || [];
  for (const f of flux) {
    const opt = document.createElement("option");
    opt.value = f;
    opt.textContent = f;
    selFlux.appendChild(opt);
  }

  checkReady();
}

function checkReady() {
  const section = document.getElementById("f-section").value;
  const flux = document.getElementById("f-flux").value;
  document.getElementById("btn-analyser").disabled = !(section && flux);
}

/* =========================================================
 * LAYOUT DYNAMIQUE
 * ========================================================= */
function ensureResultLayout() {
  const stateResult = document.getElementById("state-result");

  stateResult.innerHTML = `
    <div class="stats-row" id="stats-row" style="margin-bottom:16px;"></div>

    <div id="analysis-meta" style="margin-bottom:16px;"></div>

    <div id="graph-stack" style="display:flex;flex-direction:column;gap:16px;"></div>

    <div class="chart-container" style="margin-top:20px;">
      <div class="chart-header">
        <div>
          <div class="chart-title">Tableau détaillé des tendances</div>
          <div class="chart-subtitle">Vue analytique hebdomadaire, K-means, mensuelle et radar</div>
        </div>
      </div>
      <div style="overflow:auto;">
        <table id="tendance-table" style="width:100%;border-collapse:collapse;font-size:13px;">
          <thead></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  `;
}

function createGraphCard({ title, subtitle = "", note = "", canvasId, height = 360 }) {
  const stack = document.getElementById("graph-stack");
  const card = document.createElement("div");
  card.className = "chart-container";
  card.innerHTML = `
    <div class="chart-header" style="align-items:flex-start;">
      <div>
        <div class="chart-title">${title}</div>
        <div class="chart-subtitle">${subtitle}</div>
      </div>
    </div>
    ${
      note
        ? `<div style="color:${UI.textSoft};font-size:12px;line-height:1.45;margin:0 0 12px 0;">${note}</div>`
        : ""
    }
    <div class="chart-canvas-wrap" style="height:${height}px;">
      <canvas id="${canvasId}"></canvas>
    </div>
  `;
  stack.appendChild(card);
  return document.getElementById(canvasId);
}

function createTripleRadarRow() {
  const stack = document.getElementById("graph-stack");
  const wrapper = document.createElement("div");
  wrapper.className = "chart-container";
  wrapper.innerHTML = `
    <div class="chart-header" style="align-items:flex-start;">
      <div>
        <div class="chart-title">Indices de saisonnalité — vue radar</div>
        <div class="chart-subtitle">Mensuel, hebdomadaire et intra-mensuel</div>
      </div>
    </div>
    <div style="color:${UI.textSoft};font-size:12px;line-height:1.45;margin:0 0 12px 0;">
      Lecture : la base 100 correspond à la moyenne de référence. Un indice supérieur à 100 traduit une période forte,
      un indice inférieur à 100 une période faible.
    </div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;">
      <div class="chart-container" style="margin:0;">
        <div class="chart-header"><div><div class="chart-title" style="font-size:15px;">Radar mensuel</div></div></div>
        <div class="chart-canvas-wrap" style="height:420px;"><canvas id="chart-radar-month"></canvas></div>
      </div>
      <div class="chart-container" style="margin:0;">
        <div class="chart-header"><div><div class="chart-title" style="font-size:15px;">Radar hebdomadaire</div></div></div>
        <div class="chart-canvas-wrap" style="height:420px;"><canvas id="chart-radar-week"></canvas></div>
      </div>
      <div class="chart-container" style="margin:0;">
        <div class="chart-header"><div><div class="chart-title" style="font-size:15px;">Radar intra-mensuel</div></div></div>
        <div class="chart-canvas-wrap" style="height:420px;"><canvas id="chart-radar-intra"></canvas></div>
      </div>
    </div>
  `;
  stack.appendChild(wrapper);
}

/* =========================================================
 * UI META / STATS
 * ========================================================= */
function buildTopStats(data) {
  const statsRow = document.getElementById("stats-row");
  const global = data.global || {};

  const items = [
    { label: "Points exploités", value: fmt(global.nb_points, 0) },
    { label: "Moyenne globale", value: fmt(global.moyenne_globale) },
    { label: "Score hebdo", value: fmt(global.score_hebdo, 1) },
    { label: "Score saisonnalité", value: fmt(global.score_saisonnalite, 1) },
    { label: "Risque", value: global.niveau_risque || "—", color: global.couleur_risque || UI.accent },
  ];

  statsRow.innerHTML = items
    .map(
      (it) => `
      <div class="stat-item">
        <div class="stat-value" style="${it.color ? `color:${it.color};` : ""}">${it.value}</div>
        <div class="stat-label">${it.label}</div>
      </div>
    `
    )
    .join("");
}

function buildMeta(data) {
  const meta = document.getElementById("analysis-meta");
  const global = data.global || {};
  const debug = data.debug || {};

  meta.innerHTML = `
    <div class="chart-container" style="margin-bottom:0;">
      <div class="chart-header">
        <div>
          <div class="chart-title">${data.flux} — ${data.section}</div>
          <div class="chart-subtitle">${data.annee ? `Année ${data.annee}` : "Toutes les années"}</div>
        </div>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:10px 16px;color:${UI.textSoft};font-size:12px;line-height:1.45;">
        <div><strong style="color:${UI.text};">Orientation :</strong> ${global.orientation_flux < 0 ? "Flux négatif" : "Flux positif"}</div>
        <div><strong style="color:${UI.text};">Conservés :</strong> ${fmt(debug.kept, 0)}</div>
        <div><strong style="color:${UI.text};">Week-ends exclus :</strong> ${fmt(debug.weekend_exclus, 0)}</div>
        <div><strong style="color:${UI.text};">Sous seuil exclus :</strong> ${fmt(debug.seuil_exclus, 0)}</div>
      </div>
    </div>
  `;
}

/* =========================================================
 * OPTIONS COMMUNES CHART.JS
 * ========================================================= */
function baseCartesianOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "nearest", intersect: false },
    plugins: {
      legend: {
        labels: {
          color: UI.textSoft,
          boxWidth: 14,
          boxHeight: 8,
        },
      },
      tooltip: {
        backgroundColor: "#131720",
        borderColor: "#1A1F2E",
        borderWidth: 1,
        titleColor: "#E6EBF5",
        bodyColor: "#D1D5DB",
        padding: 12,
      },
    },
    scales: {
      x: {
        grid: { color: "#1A1F2E" },
        ticks: { color: "#FFFFFF", font: { weight: "500" } },
        title: { display: true, text: "Date", color: "#FFFFFF", font: { size: 11, weight: "500" } },
      },
      y: {
        grid: { color: "#1A1F2E" },
        ticks: {
          color: "#FFFFFF",
          callback: (v) => fmt(v, 0),
          font: { weight: "500" },
        },
        title: { display: true, text: "Montant (k€)", color: "#FFFFFF", font: { size: 11, weight: "500" } },
      },
    },
  };
}

function registerChart(chart) {
  chartRegistry.push(chart);
  return chart;
}

/* =========================================================
 * GRAPHE 1 — STRUCTURE HEBDOMADAIRE
 * ========================================================= */
function buildHebdoChart(data) {
  const canvas = createGraphCard({
    title: "Niveau 1 — Structure hebdomadaire",
    subtitle: "Moyenne par jour ouvré + stabilité issue du K-means",
    note:
      "Lecture : les barres représentent la moyenne par jour. La couleur traduit la stabilité du jour selon la domination du cluster principal. La ligne jaune pointillée représente la moyenne globale ouvrée.",
    canvasId: "chart-hebdo",
    height: 360,
  });

  const labels = data.hebdo.labels || [];
  const stats = data.hebdo.stats || [];
  const metrics = data.hebdo.cluster_metrics || {};
  const globalMean = getOr(data.global, "moyenne_globale", 0);

  const barColors = labels.map((_, i) => getOr(metrics, String(i), {}).color || UI.blue);
  const values = stats.map((s) => s.mean || 0);

  const chart = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          type: "bar",
          label: "Moyenne par jour",
          data: values,
          backgroundColor: barColors,
          borderColor: barColors,
          borderWidth: 1,
        },
        {
          type: "line",
          label: "Courbe de tendance",
          data: values,
          borderColor: UI.white,
          backgroundColor: UI.white,
          pointRadius: 4,
          pointHoverRadius: 5,
          borderWidth: 2,
          tension: 0.3,
        },
        {
          type: "line",
          label: "Moyenne globale ouvrée",
          data: labels.map(() => globalMean),
          borderColor: UI.yellow,
          borderDash: [6, 4],
          pointRadius: 0,
          borderWidth: 2,
          tension: 0,
        },
      ],
    },
    options: {
      ...baseCartesianOptions(),
      plugins: {
        ...baseCartesianOptions().plugins,
        tooltip: {
          ...baseCartesianOptions().plugins.tooltip,
          callbacks: {
            afterBody: (items) => {
              const idx = items?.[0]?.dataIndex;
              const s = stats[idx] || {};
              const m = getOr(metrics, String(idx), {});
              return [
                `CV : ${fmt(s.cv, 1)} %`,
                `Cluster dominant : ${fmt((m.share || 0) * 100, 0)} %`,
              ];
            },
          },
        },
      },
    },
  });

  registerChart(chart);
}

/* =========================================================
 * GRAPHE 2 — K-MEANS PAR JOUR
 * ========================================================= */
function buildKmeansChart(data) {
  const canvas = createGraphCard({
    title: "Segmentation K-means — différentes catégories d'un même jour",
    subtitle: "Nuages de points, centres de clusters et amplitude min/max",
    note:
      "Lecture : chaque jour ouvré peut contenir plusieurs régimes. Les points sont les observations historiques, les losanges les centres des clusters, et les segments verticaux l'étendue min/max de chaque cluster.",
    canvasId: "chart-kmeans",
    height: 420,
  });

  const labels = data.hebdo.labels || [];
  const clustersByDay = data.hebdo.clusters || {};
  const globalMean = getOr(data.global, "moyenne_globale", 0);

  const datasets = [];

  labels.forEach((label, dayIndex) => {
    const km = getOr(clustersByDay, String(dayIndex), null);
    if (!km || !km.k) return;

    const seenByCluster = {};
    const scatterPoints = [];

    const pairs = (km.values || []).map((v, idx) => ({
      value: v,
      assign: km.assignments?.[idx] ?? 0,
    }));

    pairs.sort((a, b) => (a.assign - b.assign) || (a.value - b.value));

    for (const pair of pairs) {
      seenByCluster[pair.assign] = (seenByCluster[pair.assign] || 0) + 1;
      const jitter = ((seenByCluster[pair.assign] % 7) - 3) * 0.03;
      scatterPoints.push({
        x: dayIndex + jitter,
        y: pair.value,
      });
    }

    datasets.push({
      type: "scatter",
      label: `${label} — observations`,
      data: scatterPoints,
      pointRadius: 4,
      pointHoverRadius: 5,
      backgroundColor: "rgba(255,255,255,0.25)",
      borderColor: "rgba(255,255,255,0.0)",
      showLine: false,
    });

    (km.clusters || []).forEach((cl, rank) => {
      const offsets = km.k === 1 ? [0] : km.k === 2 ? [-0.14, 0.14] : [-0.22, 0, 0.22];
      const x = dayIndex + offsets[Math.min(rank, offsets.length - 1)];

      datasets.push({
        type: "line",
        label: `${label} — ${cl.name} amplitude`,
        data: [
          { x, y: cl.min },
          { x, y: cl.max },
        ],
        borderColor: cl.color,
        backgroundColor: cl.color,
        borderWidth: 3,
        pointRadius: 0,
        tension: 0,
      });

      datasets.push({
        type: "scatter",
        label: `${label} — ${cl.name} centre`,
        data: [{ x, y: cl.center }],
        pointRadius: 8,
        pointHoverRadius: 9,
        pointStyle: "rectRot",
        backgroundColor: cl.color,
        borderColor: "#FFFFFF",
        borderWidth: 1.5,
      });
    });
  });

  datasets.push({
    type: "line",
    label: "Moyenne globale ouvrée",
    data: [
      { x: -0.5, y: globalMean },
      { x: labels.length - 0.5, y: globalMean },
    ],
    borderColor: UI.yellow,
    borderDash: [6, 4],
    borderWidth: 2,
    pointRadius: 0,
    tension: 0,
  });

  const chart = new Chart(canvas.getContext("2d"), {
    data: { datasets },
    options: {
      ...baseCartesianOptions(),
      parsing: false,
      scales: {
        x: {
          type: "linear",
          min: -0.5,
          max: labels.length - 0.5,
          grid: { color: "#1A1F2E" },
          ticks: {
            color: "#8896B0",
            stepSize: 1,
            callback: (value) => labels[value] ?? "",
          },
        },
        y: {
          grid: { color: "#1A1F2E" },
          ticks: {
            color: "#8896B0",
            callback: (v) => fmt(v, 0),
          },
        },
      },
      plugins: {
        ...baseCartesianOptions().plugins,
        legend: {
          labels: {
            color: UI.textSoft,
            filter: (item) =>
              item.text.includes("centre") ||
              item.text.includes("Moyenne globale") ||
              item.text.endsWith("observations"),
          },
        },
      },
    },
  });

  registerChart(chart);
}

/* =========================================================
 * GRAPHE 3 — GLISSEMENT MENSUEL
 * ========================================================= */
function buildSlidingChart(data) {
  const canvas = createGraphCard({
    title: "Détection des phénomènes de glissement",
    subtitle: "Jour exact du mois + moyenne glissante sur 3 jours",
    note:
      "Lecture : les barres bleues représentent la moyenne observée jour par jour dans le mois. La courbe blanche représente une fenêtre glissante sur 3 jours.",
    canvasId: "chart-sliding",
    height: 360,
  });

  const jours = data.mensuel.jours || [];
  const statsJour = data.mensuel.stats_jour || [];
  const validIdx = data.mensuel.valid_idx || [];
  const rollCenters = data.mensuel.rolling_centers || [];
  const rollStats = data.mensuel.rolling_stats || [];
  const validRollIdx = data.mensuel.valid_roll_idx || [];
  const globalMean = getOr(data.global, "moyenne_globale", 0);

  const xExact = validIdx.map((i) => jours[i]);
  const yExact = validIdx.map((i) => statsJour[i]?.mean ?? null);

  const xRoll = validRollIdx.map((i) => rollCenters[i]);
  const yRoll = validRollIdx.map((i) => rollStats[i]?.mean ?? null);

  const chart = new Chart(canvas.getContext("2d"), {
    data: {
      labels: jours.map(String),
      datasets: [
        {
          type: "bar",
          label: "Jour exact",
          data: jours.map((j, idx) => (validIdx.includes(idx) ? statsJour[idx]?.mean ?? null : null)),
          backgroundColor: "rgba(52,152,219,0.72)",
          borderColor: "#3498DB",
          borderWidth: 1,
        },
        {
          type: "line",
          label: "Fenêtre glissante 3 jours",
          data: jours.map((j, idx) => {
            const rollIndex = rollCenters.indexOf(j);
            return rollIndex >= 0 && validRollIdx.includes(rollIndex)
              ? rollStats[rollIndex]?.mean ?? null
              : null;
          }),
          borderColor: UI.white,
          backgroundColor: UI.white,
          pointRadius: 3,
          pointHoverRadius: 5,
          borderWidth: 2,
          tension: 0.3,
        },
        {
          type: "line",
          label: "Moyenne globale ouvrée",
          data: jours.map(() => globalMean),
          borderColor: UI.yellow,
          borderDash: [6, 4],
          pointRadius: 0,
          borderWidth: 2,
          tension: 0,
        },
      ],
    },
    options: {
      ...baseCartesianOptions(),
      scales: {
        x: {
          grid: { color: "#1A1F2E" },
          ticks: { color: "#8896B0", maxRotation: 70, minRotation: 70 },
        },
        y: {
          grid: { color: "#1A1F2E" },
          ticks: { color: "#8896B0", callback: (v) => fmt(v, 0) },
        },
      },
    },
  });

  registerChart(chart);
}

/* =========================================================
 * GRAPHE 4 — ANALYSE ANNUELLE
 * ========================================================= */
function buildAnnualChart(data) {
  const canvas = createGraphCard({
    title: "Analyse annuelle — tendance par mois",
    subtitle: data.annee ? `Profil mensuel — année ${data.annee}` : "Saisonnalité mensuelle moyenne",
    note:
      "Lecture : ce graphe montre le comportement moyen par mois. Si une année est sélectionnée, il décrit le profil mensuel de cette année.",
    canvasId: "chart-annual",
    height: 360,
  });

  const labels = data.annuel.labels || [];
  const stats = data.annuel.stats_mois || [];
  const globalMean = getOr(data.global, "moyenne_globale", 0);

  const chart = new Chart(canvas.getContext("2d"), {
    data: {
      labels,
      datasets: [
        {
          type: "bar",
          label: "Moyenne par mois",
          data: stats.map((s) => (s.n > 0 ? s.mean : null)),
          backgroundColor: "rgba(40,180,99,0.85)",
          borderColor: "#28B463",
          borderWidth: 1,
        },
        {
          type: "line",
          label: "Tendance annuelle",
          data: stats.map((s) => (s.n > 0 ? s.mean : null)),
          borderColor: UI.white,
          backgroundColor: UI.white,
          pointRadius: 4,
          pointHoverRadius: 5,
          borderWidth: 2,
          tension: 0.3,
        },
        {
          type: "line",
          label: "Moyenne globale ouvrée",
          data: labels.map(() => globalMean),
          borderColor: UI.yellow,
          borderDash: [6, 4],
          pointRadius: 0,
          borderWidth: 2,
          tension: 0,
        },
      ],
    },
    options: baseCartesianOptions(),
  });

  registerChart(chart);
}

/* =========================================================
 * GRAPHE 5 — RÉEL VS PRÉVISIONS
 * ========================================================= */
function buildForecastChart(data) {
  const canvas = createGraphCard({
    title: "Série temporelle — réalisé vs prévisions",
    subtitle: "Comparaison directe sur la période brute",
    note:
      "Lecture : ce graphe reprend le besoin initial réel vs prévisions, en plus des analyses avancées.",
    canvasId: "chart-forecast",
    height: 380,
  });

  const labels = (data.dates_serie || []).map((d) => {
    const dt = new Date(d);
    return Number.isNaN(dt.getTime())
      ? d
      : dt.toLocaleDateString("fr-FR", { day: "2-digit", month: "short" });
  });

  const colors = [
    "#A78BFA",
    "#32D583",
    "#FB923C",
    "#F472B6",
    "#38BDF8",
    "#FBBF24",
  ];

  const datasets = [
    {
      label: "Réalisé",
      data: data.reel_serie || [],
      borderColor: "#4D87F5",
      backgroundColor: "rgba(77,135,245,.10)",
      borderWidth: 2.5,
      pointRadius: 2,
      pointHoverRadius: 5,
      fill: true,
      tension: 0.35,
    },
  ];

  (data.previsions || []).forEach((p, i) => {
    datasets.push({
      label: p.label,
      data: p.values || [],
      borderColor: colors[i % colors.length],
      backgroundColor: "transparent",
      borderWidth: 1.5,
      borderDash: [5, 4],
      pointRadius: 1.5,
      pointHoverRadius: 4,
      fill: false,
      tension: 0.35,
    });
  });

  const chart = new Chart(canvas.getContext("2d"), {
    type: "line",
    data: { labels, datasets },
    options: {
      ...baseCartesianOptions(),
      interaction: { mode: "index", intersect: false },
    },
  });

  registerChart(chart);
}

/* =========================================================
 * RADARS
 * ========================================================= */
function radarOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: UI.textSoft,
        },
      },
      tooltip: {
        backgroundColor: "#131720",
        borderColor: "#1A1F2E",
        borderWidth: 1,
        titleColor: "#E6EBF5",
        bodyColor: "#D1D5DB",
        padding: 12,
      },
    },
    scales: {
      r: {
        angleLines: { color: UI.border },
        grid: { color: UI.border },
        pointLabels: {
          color: UI.white,
          font: { size: 11, weight: "bold" },
        },
        ticks: {
          color: UI.muted,
          backdropColor: "transparent",
        },
      },
    },
  };
}

function buildRadar(canvasId, labels, values, label, color) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const safeLabels = labels?.length ? labels : ["Aucune donnée"];
  const safeValues = values?.length ? values : [100];

  const chart = new Chart(canvas.getContext("2d"), {
    type: "radar",
    data: {
      labels: safeLabels,
      datasets: [
        {
          label,
          data: safeValues,
          borderColor: color,
          backgroundColor: `${color}33`,
          pointBackgroundColor: color,
          pointBorderColor: UI.white,
          pointRadius: 4,
          borderWidth: 2,
        },
        {
          label: "Base 100",
          data: safeLabels.map(() => 100),
          borderColor: "#9CA3AF",
          borderDash: [6, 4],
          pointRadius: 0,
          borderWidth: 1.5,
        },
      ],
    },
    options: radarOptions(),
  });

  registerChart(chart);
}

function buildRadars(data) {
  createTripleRadarRow();

  const mensuel = data.radars?.mensuel || {};
  const mensuelLabels = mensuel.labels || [];
  const mensuelValues = (mensuel.periodes || []).map((p) => mensuel.indices?.[p] ?? null);

  const hebdo = data.radars?.hebdo || {};
  const intra = data.radars?.intra_mensuel || {};

  buildRadar("chart-radar-month", mensuelLabels, mensuelValues, "Indice mensuel", "#4C7CF3");
  buildRadar("chart-radar-week", hebdo.labels || [], hebdo.values || [], "Indice hebdomadaire", "#F5A623");
  buildRadar("chart-radar-intra", intra.labels || [], intra.values || [], "Indice intra-mensuel", "#00C8B4");
}

/* =========================================================
 * TABLEAU DÉTAILLÉ
 * ========================================================= */
function createCell(tag, text) {
  const td = document.createElement(tag);
  td.textContent = text;
  td.style.padding = "10px 12px";
  td.style.borderBottom = `1px solid ${UI.borderSoft}`;
  td.style.textAlign = "center";
  td.style.whiteSpace = "nowrap";
  return td;
}

function addRow(tbody, values, bg = "") {
  const tr = document.createElement("tr");
  if (bg) tr.style.background = bg;
  values.forEach((v) => tr.appendChild(createCell("td", v)));
  tbody.appendChild(tr);
}

function buildTable(data) {
  const table = document.getElementById("tendance-table");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  const colonnes = [
    "Bloc",
    "Libellé",
    "Nb points",
    "Moyenne",
    "Médiane",
    "Min",
    "Max",
    "Écart-type",
    "CV %",
    "IC 95%",
    "Lecture",
  ];

  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trHead = document.createElement("tr");
  colonnes.forEach((col) => {
    const th = createCell("th", col);
    th.style.background = "#1D2634";
    th.style.color = UI.text;
    th.style.fontWeight = "700";
    th.style.position = "sticky";
    th.style.top = "0";
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  const jours = data.hebdo.labels || [];
  const hebdoStats = data.hebdo.stats || [];
  const metrics = data.hebdo.cluster_metrics || {};

  jours.forEach((jour, i) => {
    const s = hebdoStats[i];
    if (!s || !s.n) return;
    const kmInfo = getOr(metrics, String(i), {});
    const bg =
      kmInfo.tag === "stable" ? "#143A2E" :
      kmInfo.tag === "variable" ? "#5C4A1F" :
      "#5C1F1F";

    addRow(tbody, [
      "Hebdo",
      jour,
      fmt(s.n, 0),
      fmt(s.mean),
      fmt(s.median),
      fmt(s.min),
      fmt(s.max),
      fmt(s.stdev),
      fmt(s.cv, 1),
      fmtRange(s.ic_low, s.ic_high, 0),
      `${kmInfo.label || "—"} — cluster dominant ${fmt((kmInfo.share || 0) * 100, 1)}%`,
    ], bg);
  });

  const clustersByDay = data.hebdo.clusters || {};
  jours.forEach((jour, i) => {
    const km = getOr(clustersByDay, String(i), null);
    if (!km || !km.k) return;

    (km.clusters || []).forEach((cl) => {
      addRow(tbody, [
        "K-means",
        `${jour} - ${cl.name}`,
        fmt(cl.n, 0),
        fmt(cl.mean),
        fmt(cl.median),
        fmt(cl.min),
        fmt(cl.max),
        fmt(cl.stdev),
        fmt(cl.cv, 1),
        fmtRange(cl.ic_low, cl.ic_high, 0),
        `Centre = ${fmt(cl.center, 0)}`,
      ], "#253B56");
    });
  });

  const joursMois = data.mensuel.jours || [];
  const statsJour = data.mensuel.stats_jour || [];
  const validMonth = data.mensuel.valid_idx || [];

  validMonth.forEach((i) => {
    const s = statsJour[i];
    if (!s || !s.n) return;

    const bg =
      s.tag === "stable" ? "#143A2E" :
      s.tag === "variable" ? "#5C4A1F" :
      "#5C1F1F";

    addRow(tbody, [
      "Mensuel",
      String(joursMois[i]),
      fmt(s.n, 0),
      fmt(s.mean),
      fmt(s.median),
      fmt(s.min),
      fmt(s.max),
      fmt(s.stdev),
      fmt(s.cv, 1),
      fmtRange(s.ic_low, s.ic_high, 0),
      s.label,
    ], bg);
  });

  const positions = data.mensuel.positions || {};
  ["Début de mois", "Milieu de mois", "Fin de mois"].forEach((label) => {
    const s = positions[label];
    if (!s) return;

    addRow(tbody, [
      "Synthèse",
      label,
      fmt(s.n, 0),
      fmt(s.mean),
      fmt(s.median),
      fmt(s.min),
      fmt(s.max),
      fmt(s.stdev),
      fmt(s.cv, 1),
      fmtRange(s.ic_low, s.ic_high, 0),
      s.label,
    ], "#1B365D");
  });

  const mensuelRadar = data.radars?.mensuel || {};
  const periodes = mensuelRadar.periodes || [];
  const moisLong = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
  ];

  periodes.forEach((p) => {
    const indice = mensuelRadar.indices?.[p];
    const lecture =
      indice >= 105 ? `Fort (+${fmt(indice - 100, 1)})` :
      indice <= 95 ? `Faible (${fmt(indice - 100, 1)})` :
      "Proche de la moyenne";

    addRow(tbody, [
      "Radar Mensuel",
      moisLong[p - 1] || String(p),
      "—",
      "—",
      "—",
      "—",
      "—",
      "—",
      "—",
      "—",
      `Indice ${fmt(indice, 1)} — ${lecture}`,
    ], "#1A2E4A");
  });

  const hebdoRadar = data.radars?.hebdo || {};
  (hebdoRadar.labels || []).forEach((label, i) => {
    const s = hebdoStats[jours.indexOf(label)];
    const indice = hebdoRadar.values?.[i];

    const lecture =
      indice >= 105 ? `Fort (+${fmt(indice - 100, 1)})` :
      indice <= 95 ? `Faible (${fmt(indice - 100, 1)})` :
      "Proche de la moyenne";

    addRow(tbody, [
      "Radar Hebdo",
      label,
      fmt(s?.n, 0),
      fmt(s?.mean),
      fmt(s?.median),
      fmt(s?.min),
      fmt(s?.max),
      fmt(s?.stdev),
      fmt(s?.cv, 1),
      s ? fmtRange(s.ic_low, s.ic_high, 0) : "—",
      `Indice ${fmt(indice, 1)} — ${lecture}`,
    ], "#1A2E4A");
  });

  const intra = data.radars?.intra_mensuel || {};
  (intra.blocs || []).forEach((bloc, i) => {
    const s = bloc.stats;
    const indice = intra.values?.[i];
    const label = bloc.start === bloc.end ? `J${bloc.start}` : `J${bloc.start}–J${bloc.end}`;

    const lecture =
      indice >= 105 ? `Fort (+${fmt(indice - 100, 1)})` :
      indice <= 95 ? `Faible (${fmt(indice - 100, 1)})` :
      "Proche de la moyenne";

    addRow(tbody, [
      "Radar Intra-M",
      label,
      fmt(s?.n, 0),
      fmt(s?.mean),
      fmt(s?.median),
      fmt(s?.min),
      fmt(s?.max),
      fmt(s?.stdev),
      fmt(s?.cv, 1),
      s ? fmtRange(s.ic_low, s.ic_high, 0) : "—",
      `Indice ${fmt(indice, 1)} — ${lecture}`,
    ], "#1A2E4A");
  });
}

/* =========================================================
 * PIPELINE DE RENDU
 * ========================================================= */
function buildFullAnalysis(data) {
  destroyCharts();
  ensureResultLayout();
  buildTopStats(data);
  buildMeta(data);
  buildHebdoChart(data);
  buildKmeansChart(data);
  buildSlidingChart(data);
  buildAnnualChart(data);
  buildRadars(data);
  buildForecastChart(data);
  buildTable(data);
}

/* =========================================================
 * FETCH
 * ========================================================= */
async function analyser() {
  const section = document.getElementById("f-section").value;
  const flux = document.getElementById("f-flux").value;
  const annee = document.getElementById("f-annee").value;

  if (!section || !flux) return;

  showState("state-loading");

  const params = new URLSearchParams({ section, flux });
  if (annee) params.set("annee", annee);

  try {
    const res = await fetch(`/api/tendance?${params}`);
    const body = await res.json().catch(() => ({}));

    if (!res.ok) {
      throw new Error(body.error || `HTTP ${res.status}`);
    }

    if (body.error) {
      throw new Error(body.error);
    }

    buildFullAnalysis(body);
    showState("state-result");

    /* Activer export Excel */
    const btnExcel = document.getElementById("btn-export-excel");
    if (btnExcel) {
      btnExcel.disabled = false;
      btnExcel._tendanceBody = body;
    }
  } catch (err) {
    destroyCharts();
    const errorMsg = document.getElementById("error-msg");
    if (errorMsg) errorMsg.textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
}

/* =========================================================
 * INIT
 * ========================================================= */
(async () => {
  try {
    await loadCatalogue();
    showState("state-empty");
  } catch (err) {
    const errorMsg = document.getElementById("error-msg");
    if (errorMsg) errorMsg.textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
})();