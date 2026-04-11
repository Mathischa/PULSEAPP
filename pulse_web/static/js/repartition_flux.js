/* repartition_flux.js — Répartition des écarts par flux */
"use strict";

/* =========================================================
 * CONFIG CHART.JS
 * ========================================================= */
Chart.defaults.color = "#8896B0";
Chart.defaults.borderColor = "#1A1F2E";
Chart.defaults.font.family = "Inter, system-ui, sans-serif";
Chart.defaults.font.size = 12;

const UI = {
  accent:  "#4C7CF3",
  surface2:"#1A2230",
  text:    "#F3F4F6",
  muted:   "#9CA3AF",
  green:   "#27AE60",
  red:     "#C0392B",
  blue:    "#3498DB",
};

let chartVolume = null;
let chartFreq   = null;
let chartValo   = null;

/* =========================================================
 * ÉTAT
 * ========================================================= */
function showState(id) {
  ["state-loading", "state-error", "state-result"].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.hidden = (s !== id);
  });
}

function destroyCharts() {
  [chartVolume, chartFreq, chartValo].forEach(c => { try { c && c.destroy(); } catch(_){} });
  chartVolume = chartFreq = chartValo = null;
}

/* =========================================================
 * FORMAT
 * ========================================================= */
const fmt = (v, d = 0) =>
  typeof v === "number" && isFinite(v)
    ? v.toLocaleString("fr-FR", { minimumFractionDigits: d, maximumFractionDigits: d })
    : "—";

/* =========================================================
 * COULEURS DÉGRADÉ
 * ========================================================= */
function blueGradient(n) {
  // du bleu-gris clair au bleu accent, indexé par rang (0 = max)
  return Array.from({ length: n }, (_, i) => {
    const t = n === 1 ? 1 : 1 - i / (n - 1);   // 1 pour la barre la plus haute
    const r = Math.round(28  + t * (76  - 28));
    const g = Math.round(54  + t * (124 - 54));
    const b = Math.round(120 + t * (243 - 120));
    return `rgba(${r},${g},${b},0.92)`;
  });
}

function divergingColors(values) {
  return values.map(v =>
    v >= 0
      ? `rgba(39,174,96,0.85)`   // vert positif
      : `rgba(192,57,43,0.85)`   // rouge négatif
  );
}

/* =========================================================
 * CHART COMMUN (barh horizontal)
 * ========================================================= */
function makeBarH({ canvasId, labels, values, colors, xlabel, xMin, xMax, refLine }) {
  const ctx = document.getElementById(canvasId).getContext("2d");
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors,
        borderColor: "transparent",
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 500 },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` ${fmt(ctx.raw, 1)}${xlabel ? "  " + xlabel : ""}`,
          }
        },
      },
      scales: {
        x: {
          min: xMin,
          max: xMax,
          grid: { color: "rgba(255,255,255,0.07)" },
          ticks: { color: UI.muted },
          title: { display: !!xlabel, text: xlabel, color: UI.muted, font: { size: 11 } },
        },
        y: {
          grid: { display: false },
          ticks: {
            color: UI.text,
            font: { size: 12 },
            // Tronquer les labels trop longs
            callback(v) {
              const lbl = this.getLabelForValue(v);
              return lbl.length > 36 ? lbl.slice(0, 34) + "…" : lbl;
            }
          }
        }
      },
      ...(refLine != null ? {
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => ` ${fmt(ctx.raw, 0)} k€`,
            }
          },
          annotation: undefined,   // Chart.js annotation plugin non chargé, on dessine manuellement
        }
      } : {})
    }
  });
}

/* =========================================================
 * RESIZE DYNAMIQUE
 * ========================================================= */
function adaptChartHeight(wrapId, count) {
  const el = document.getElementById(wrapId);
  if (!el) return;
  const h = Math.min(Math.max(300, 36 * count + 100), 1400);
  el.style.height = h + "px";
}

/* =========================================================
 * KPIs
 * ========================================================= */
function renderKPIs(flux) {
  const row = document.getElementById("kpi-row");
  row.innerHTML = "";

  const totalPrev   = flux.reduce((s, f) => s + f.nb_previsions, 0);
  const totalEcarts = flux.reduce((s, f) => s + f.nb_ecarts, 0);
  const pctGlobal   = totalPrev > 0 ? totalEcarts / totalPrev * 100 : 0;
  const totalValo   = flux.reduce((s, f) => s + f.valeur_ecarts, 0);
  const nbFlux      = flux.length;

  const kpis = [
    { label: "Flux analysés",      value: fmt(nbFlux) },
    { label: "Prévisions totales", value: fmt(totalPrev) },
    { label: "Écarts ≥ 40 %",      value: fmt(totalEcarts) },
    { label: "% Écarts global",    value: fmt(pctGlobal, 1) + " %" },
    { label: "Valorisation (k€)",  value: fmt(Math.round(totalValo)) },
  ];

  kpis.forEach(({ label, value }) => {
    const card = document.createElement("div");
    card.className = "kpi-card";
    card.innerHTML = `<div class="kpi-value">${value}</div><div class="kpi-label">${label}</div>`;
    row.appendChild(card);
  });
}

/* =========================================================
 * TABLEAU
 * ========================================================= */
function renderTable(flux) {
  const tbody = document.getElementById("tbody-flux");
  tbody.innerHTML = "";

  let totalPrev = 0, totalEcarts = 0, totalValo = 0;

  flux.forEach(f => {
    totalPrev   += f.nb_previsions;
    totalEcarts += f.nb_ecarts;
    totalValo   += f.valeur_ecarts;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${f.nom}</td>
      <td class="num">${fmt(f.nb_previsions)}</td>
      <td class="num">${fmt(f.nb_ecarts)}</td>
      <td class="num">${fmt(f.pct_ecarts, 1)} %</td>
      <td class="num">${fmt(Math.round(f.valeur_ecarts))}</td>
    `;
    tbody.appendChild(tr);
  });

  // Ligne TOTAL
  const pctTotal = totalPrev > 0 ? totalEcarts / totalPrev * 100 : 0;
  const trTotal = document.createElement("tr");
  trTotal.style.fontWeight = "bold";
  trTotal.style.borderTop = "1px solid rgba(255,255,255,.15)";
  trTotal.innerHTML = `
    <td>TOTAL</td>
    <td class="num">${fmt(totalPrev)}</td>
    <td class="num">${fmt(totalEcarts)}</td>
    <td class="num">${fmt(pctTotal, 1)} %</td>
    <td class="num">${fmt(Math.round(totalValo))}</td>
  `;
  tbody.appendChild(trTotal);
}

/* =========================================================
 * RENDER PRINCIPAL
 * ========================================================= */
function renderAll(data) {
  const flux = data.flux;

  if (!flux || flux.length === 0) {
    showState("state-error");
    document.getElementById("error-msg").textContent =
      "Aucun écart important détecté pour ce filtre.";
    return;
  }

  destroyCharts();

  const labels  = flux.map(f => f.nom);
  const volumes = flux.map(f => f.nb_ecarts);
  const freqs   = flux.map(f => f.pct_ecarts);
  const valos   = flux.map(f => Math.round(f.valeur_ecarts));

  const colorsBlue = blueGradient(labels.length);
  const colorsValo = divergingColors(valos);

  const maxVol   = Math.max(...volumes, 1);
  const maxFreq  = Math.max(...freqs, 1);
  const maxValo  = Math.max(...valos.map(Math.abs), 1);
  const padValo  = maxValo * 0.15;

  // Adapter la hauteur des wraps
  adaptChartHeight("wrap-chart1", labels.length);
  adaptChartHeight("wrap-chart2", labels.length);
  adaptChartHeight("wrap-chart3", labels.length);

  // Chart 1 — Volume
  chartVolume = makeBarH({
    canvasId: "chart-volume",
    labels,
    values: volumes,
    colors: colorsBlue,
    xlabel: "Nombre d'écarts ≥ 40 %",
    xMin: 0,
    xMax: Math.ceil(maxVol * 1.15),
  });

  // Chart 2 — Fréquence
  chartFreq = makeBarH({
    canvasId: "chart-freq",
    labels,
    values: freqs,
    colors: colorsBlue,
    xlabel: "% d'écarts / prévisions",
    xMin: 0,
    xMax: Math.ceil(maxFreq * 1.15),
  });

  // Chart 3 — Valorisation divergente
  chartValo = makeBarH({
    canvasId: "chart-valo",
    labels,
    values: valos,
    colors: colorsValo,
    xlabel: "Valorisation signée (k€)",
    xMin: -(maxValo + padValo),
    xMax:   maxValo + padValo,
    refLine: 0,
  });

  // KPIs + tableau
  renderKPIs(flux);
  renderTable(flux);

  showState("state-result");
}

/* =========================================================
 * FETCH DONNÉES
 * ========================================================= */
async function analyser() {
  showState("state-loading");

  const section = document.getElementById("f-section").value;
  const annee   = document.getElementById("f-annee").value;
  const profil  = document.getElementById("f-profil").value;

  const params = new URLSearchParams();
  if (section) params.set("section", section);
  if (annee)   params.set("annee", annee);
  if (profil)  params.set("profil", profil);

  try {
    const res  = await fetch(`/api/repartition_flux?${params}`);
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
    renderAll(body);
  } catch (err) {
    destroyCharts();
    document.getElementById("error-msg").textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
}

/* =========================================================
 * CHARGEMENT DES PROFILS (dynamique)
 * ========================================================= */
async function loadProfils() {
  const section = document.getElementById("f-section").value;
  const annee   = document.getElementById("f-annee").value;

  const selProfil = document.getElementById("f-profil");
  const currentProfil = selProfil.value;

  selProfil.innerHTML = `<option value="">Tous profils</option>`;

  const params = new URLSearchParams();
  if (section) params.set("section", section);
  if (annee)   params.set("annee", annee);

  try {
    const res  = await fetch(`/api/repartition_flux/profils?${params}`);
    const body = await res.json();
    (body.profils || []).forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      if (p === currentProfil) opt.selected = true;
      selProfil.appendChild(opt);
    });
  } catch (_) {
    // silencieux
  }
}

/* =========================================================
 * INIT
 * ========================================================= */
(async () => {
  try {
    // Catalogue → sections
    const catRes = await fetch("/api/catalogue");
    if (catRes.ok) {
      const cat = await catRes.json();
      const selSection = document.getElementById("f-section");
      Object.keys(cat).sort().forEach(s => {
        const opt = document.createElement("option");
        opt.value = s;
        opt.textContent = s;
        selSection.appendChild(opt);
      });
    }

    // Accueil → années
    const metaRes = await fetch("/api/accueil");
    if (metaRes.ok) {
      const meta = await metaRes.json();
      const annees = (meta.annees || []).slice().sort((a, b) => a - b);
      const selAnnee = document.getElementById("f-annee");
      selAnnee.innerHTML = `<option value="">Toutes</option>`;
      annees.forEach(y => {
        const opt = document.createElement("option");
        opt.value = String(y);
        opt.textContent = String(y);
        selAnnee.appendChild(opt);
      });
      if (annees.length) {
        const currentYear = new Date().getFullYear();
        const best = annees.filter(y => y < currentYear).pop() ?? annees[annees.length - 1];
        selAnnee.value = String(best);
      }
    }

    // Événements filtres
    document.getElementById("f-section").addEventListener("change", async () => {
      await loadProfils();
      analyser();
    });
    document.getElementById("f-annee").addEventListener("change", async () => {
      await loadProfils();
      analyser();
    });
    document.getElementById("f-profil").addEventListener("change", analyser);
    document.getElementById("btn-analyser").addEventListener("click", analyser);

    // Chargement initial des profils + analyse
    await loadProfils();
    analyser();

  } catch (err) {
    document.getElementById("error-msg").textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
})();
