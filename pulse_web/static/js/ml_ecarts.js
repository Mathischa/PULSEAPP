/* ml_ecarts.js */
"use strict";

Chart.defaults.color       = "#64748B";
Chart.defaults.borderColor = "#1A2230";
Chart.defaults.font.family = "Inter, system-ui, sans-serif";
Chart.defaults.font.size   = 12;

/* ── ÉTAT ─────────────────────────────────────────────────── */
let _algo      = "kmeans";
let _mode      = "cluster";
let _catalogue = {};
let _chartScatter = null;
let _chartBarCl   = null;
let _chartExplain = null;

/* ── UTILS ────────────────────────────────────────────────── */
const fmt  = v => typeof v === "number" && isFinite(v)
  ? v.toLocaleString("fr-FR", { maximumFractionDigits: 0 })
  : "—";
const fmt1 = v => typeof v === "number" && isFinite(v)
  ? v.toLocaleString("fr-FR", { maximumFractionDigits: 1 })
  : "—";
const fmt2 = v => typeof v === "number" && isFinite(v)
  ? v.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  : "—";
function escHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function destroyChart(c) { if (c) { try { c.destroy(); } catch(_) {} } return null; }

/* ── ÉTATS ────────────────────────────────────────────────── */
function showState(id) {
  ["state-idle","state-loading","state-error","state-cluster","state-explain"].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.hidden = (s !== id);
  });
}

/* ── ALGO / MODE TOGGLES ──────────────────────────────────── */
function setAlgo(a) {
  _algo = a;
  document.getElementById("btn-kmeans").classList.toggle("active", a === "kmeans");
  document.getElementById("btn-dbscan").classList.toggle("active", a === "dbscan");
}
function setMode(m) {
  _mode = m;
  document.getElementById("btn-cluster").classList.toggle("active", m === "cluster");
  document.getElementById("btn-explain").classList.toggle("active", m === "explain");
}
window.setAlgo = setAlgo;
window.setMode = setMode;

/* ── SCATTER ──────────────────────────────────────────────── */
function renderScatter(points, clusters) {
  _chartScatter = destroyChart(_chartScatter);

  // Grouper par cluster
  const byCluster = {};
  for (const p of points) {
    if (!byCluster[p.cl_name]) byCluster[p.cl_name] = { color: p.color, data: [] };
    byCluster[p.cl_name].data.push({ x: p.x, y: p.y, ...p });
  }

  // Dataset anomalies (cerclés en blanc)
  const anomalies = points.filter(p => p.outlier);

  const datasets = Object.entries(byCluster).map(([name, info]) => ({
    label:           name,
    data:            info.data,
    backgroundColor: info.color + "CC",
    borderColor:     info.color,
    borderWidth:     0,
    pointRadius:     4,
    pointHoverRadius:7,
  }));

  if (anomalies.length) {
    datasets.push({
      label:           "Anomalies",
      data:            anomalies,
      backgroundColor: "rgba(0,0,0,0)",
      borderColor:     "#FFFFFF",
      borderWidth:     1.5,
      pointRadius:     7,
      pointHoverRadius:9,
      pointStyle:      "circle",
    });
  }

  // Centroides (X noir)
  const withCentroid = clusters.filter(c => c.cx !== undefined);
  if (withCentroid.length) {
    datasets.push({
      label:           "Centroides",
      data:            withCentroid.map(c => ({ x: c.cx, y: c.cy })),
      backgroundColor: "#000",
      borderColor:     "#fff",
      borderWidth:     1.5,
      pointRadius:     10,
      pointHoverRadius:12,
      pointStyle:      "crossRot",
    });
  }

  const ctx = document.getElementById("chart-scatter").getContext("2d");
  _chartScatter = new Chart(ctx, {
    type: "scatter",
    data: { datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "point", intersect: true },
      plugins: {
        legend: {
          position: "bottom",
          labels: { color: "#CBD5E1", boxWidth: 10, padding: 16, font: { size: 11 }, usePointStyle: true },
        },
        tooltip: {
          backgroundColor: "#1A2230",
          borderColor: "#2B3647", borderWidth: 1,
          titleColor: "#F3F4F6", bodyColor: "#94A3B8", padding: 12,
          callbacks: {
            title: items => items[0]?.raw?.cl_name ?? "",
            label: item => {
              const r = item.raw;
              if (!r?.flux) return ` x=${fmt2(item.parsed.x)}  y=${fmt(item.parsed.y)}`;
              return [
                ` Flux : ${r.flux}`,
                ` Date : ${r.date}`,
                ` Écart : ${fmt2(r.x)} %`,
                ` Valo  : ${fmt(r.y)} k€`,
                r.outlier ? " ⚠ Anomalie" : "",
              ].filter(Boolean);
            },
          },
        },
      },
      scales: {
        x: {
          title: { display: true, text: "Écart (%)", color: "#FFFFFF", font: { size: 12, weight: "500" } },
          grid: { color: "rgba(255,255,255,.04)" },
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          border: { color: "rgba(255,255,255,.08)" },
        },
        y: {
          title: { display: true, text: "Valorisation signée (k€)", color: "#FFFFFF", font: { size: 12, weight: "500" } },
          grid: { color: "rgba(255,255,255,.04)" },
          ticks: { color: "#FFFFFF", callback: v => fmt(v), font: { weight: "500" } },
          border: { color: "rgba(255,255,255,.08)" },
        },
      },
    },
  });
}

/* ── BAR CLUSTERS ─────────────────────────────────────────── */
function renderBarClusters(clusters) {
  _chartBarCl = destroyChart(_chartBarCl);
  const ctx = document.getElementById("chart-bar-cl").getContext("2d");
  _chartBarCl = new Chart(ctx, {
    type: "bar",
    data: {
      labels:   clusters.map(c => c.label),
      datasets: [{
        label:           "% des points",
        data:            clusters.map(c => c.pct),
        backgroundColor: clusters.map(c => c.color + "CC"),
        borderColor:     clusters.map(c => c.color),
        borderWidth:     1,
        borderRadius:    6,
        borderSkipped:   false,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "#1A2230", borderColor: "#2B3647", borderWidth: 1,
          titleColor: "#F3F4F6", bodyColor: "#94A3B8", padding: 12,
          callbacks: {
            label: ctx => ` ${fmt1(ctx.parsed.y)} % — ${clusters[ctx.dataIndex].count} points`,
          },
        },
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Clusters", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          border: { color: "rgba(255,255,255,.08)" },
        },
        y: {
          grid: { color: "rgba(255,255,255,.04)" },
          ticks: { color: "#FFFFFF", callback: v => v + "%", font: { weight: "500" } },
          title: { display: true, text: "Pourcentage (%)", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          border: { color: "rgba(255,255,255,.08)" },
          max: Math.ceil(Math.max(...clusters.map(c => c.pct)) * 1.25 / 10) * 10,
        },
      },
    },
  });
}

/* ── LÉGENDE CLUSTERS ─────────────────────────────────────── */
function renderLegend(clusters) {
  const el = document.getElementById("scatter-legend");
  el.innerHTML = clusters.map(c =>
    `<span class="ml-legend-chip">
      <span class="ml-legend-dot" style="background:${c.color}"></span>
      ${escHtml(c.label)} — ${c.count} pts (${c.pct}%)
    </span>`
  ).join("");
}

/* ── KPI ROW ──────────────────────────────────────────────── */
function renderKPIs(gs) {
  const sign    = gs.total_impact >= 0 ? "SURPLUS" : "DÉFICIT";
  const aKlass  = gs.anom_pct > 15 ? "ml-kpi--danger" : gs.anom_pct > 8 ? "ml-kpi--warn" : "ml-kpi--ok";
  document.getElementById("kpi-row").innerHTML = `
    <div class="ml-kpi">
      <div class="ml-kpi-label">Points analysés</div>
      <div class="ml-kpi-value">${gs.n_total.toLocaleString("fr-FR")}</div>
      <div class="ml-kpi-sub">Paires (réel, prévision)</div>
    </div>
    <div class="ml-kpi ${aKlass}">
      <div class="ml-kpi-label">Anomalies</div>
      <div class="ml-kpi-value">${gs.n_anomalies} <small style="font-size:13px;">(${gs.anom_pct}%)</small></div>
      <div class="ml-kpi-sub">IsolationForest</div>
    </div>
    <div class="ml-kpi">
      <div class="ml-kpi-label">Impact global</div>
      <div class="ml-kpi-value" style="font-size:16px;">${fmt(gs.total_impact)} k€</div>
      <div class="ml-kpi-sub">${sign}</div>
    </div>
    <div class="ml-kpi">
      <div class="ml-kpi-label">Flux à risque</div>
      <div class="ml-kpi-value" style="font-size:13px;line-height:1.3;">${escHtml(gs.top_flux)}</div>
      <div class="ml-kpi-sub">${fmt(gs.top_flux_val)} k€ cumulé</div>
    </div>
  `;
}

/* ── TABLE SUMMARY ────────────────────────────────────────── */
function renderSummaryTable(summary, gs) {
  const tbody = document.getElementById("summary-tbody");
  tbody.innerHTML = summary.map(r => `
    <tr>
      <td><span class="cl-dot" style="background:${r.color}"></span>${escHtml(r.label)}</td>
      <td>${r.count}</td>
      <td>${r.pct}%</td>
      <td>${fmt2(r.mean_pct)}</td>
      <td>${fmt1(r.mean_valo)}</td>
      <td>${fmt(r.sum_valo)}</td>
      <td>${r.anomalies} (${r.anom_pct}%)</td>
    </tr>
  `).join("") + `
    <tr class="total-row">
      <td>Total</td>
      <td>${gs.n_total}</td>
      <td>100%</td>
      <td>—</td><td>—</td>
      <td>${fmt(gs.total_impact)}</td>
      <td>${gs.n_anomalies} (${gs.anom_pct}%)</td>
    </tr>
  `;
}

/* ── DIAGNOSTIC ───────────────────────────────────────────── */
function renderDiagnostic(gs) {
  const el = document.getElementById("ml-diagnostic");
  let klass, icon, msg;
  if (gs.anom_pct > 15) {
    klass = "ml-diagnostic--danger";
    icon  = "🔴";
    msg   = `Taux d'anomalies élevé (${gs.anom_pct}%). Investigation immédiate recommandée.`;
  } else if (gs.anom_pct > 8) {
    klass = "ml-diagnostic--warn";
    icon  = "🟡";
    msg   = `Taux d'anomalies modéré (${gs.anom_pct}%). À surveiller.`;
  } else {
    klass = "ml-diagnostic--ok";
    icon  = "🟢";
    msg   = `Distribution normale des écarts. Taux d'anomalies acceptable (${gs.anom_pct}%).`;
  }
  const sign = gs.total_impact >= 0 ? "surplus" : "déficit";
  el.innerHTML = `
    <div class="ml-diagnostic ${klass}">
      <div class="ml-diagnostic-title">${icon} Diagnostic automatique</div>
      ${escHtml(msg)}<br>
      <strong>Section la plus risquée :</strong> ${escHtml(gs.top_section)} (${fmt(gs.top_section_val)} k€)<br>
      <strong>Flux le plus risqué :</strong> ${escHtml(gs.top_flux)} (${fmt(gs.top_flux_val)} k€)<br>
      <strong>Impact global :</strong> ${fmt(gs.total_impact)} k€ ${sign}
    </div>
  `;
}

/* ── EXPLAIN CHART ────────────────────────────────────────── */
function renderExplain(data) {
  _chartExplain = destroyChart(_chartExplain);

  const sorted = [...data.features].sort((a, b) => a.score - b.score);

  const ctx = document.getElementById("chart-explain").getContext("2d");
  _chartExplain = new Chart(ctx, {
    type: "bar",
    data: {
      labels: sorted.map(f => f.label),
      datasets: [{
        label:           "Score F",
        data:            sorted.map(f => f.score),
        backgroundColor: sorted.map(f => f.score === Math.max(...data.features.map(x => x.score))
          ? "#4C7CF3CC" : "#4C7CF355"),
        borderColor:     "#4C7CF3",
        borderWidth:     1,
        borderRadius:    6,
        borderSkipped:   false,
        indexAxis:       "y",
      }],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "#1A2230", borderColor: "#2B3647", borderWidth: 1,
          titleColor: "#F3F4F6", bodyColor: "#94A3B8", padding: 12,
          callbacks: { label: ctx => ` Score F : ${fmt2(ctx.parsed.x)}` },
        },
      },
      scales: {
        x: {
          grid: { color: "rgba(255,255,255,.04)" },
          ticks: { color: "#FFFFFF", callback: v => fmt2(v), font: { weight: "500" } },
          title: { display: true, text: "Score F", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          border: { color: "rgba(255,255,255,.08)" },
        },
        y: {
          grid: { display: false },
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Flux", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          border: { color: "rgba(255,255,255,.08)" },
        },
      },
    },
  });

  // Interprétation
  const interp = document.getElementById("explain-interp");
  const hasSignal = data.top_score > 2;
  interp.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;">
      <div class="ml-kpi">
        <div class="ml-kpi-label">Seuil gros écarts</div>
        <div class="ml-kpi-value" style="font-size:16px;">${fmt(data.seuil)} k€</div>
        <div class="ml-kpi-sub">|valo| ≥ ce seuil → gros écart</div>
      </div>
      <div class="ml-kpi">
        <div class="ml-kpi-label">Points classifiés "gros écart"</div>
        <div class="ml-kpi-value">${data.n_gros} <small style="font-size:13px;">(${fmt1(data.n_gros/data.n_total*100)}%)</small></div>
        <div class="ml-kpi-sub">Cible : quantile 75%</div>
      </div>
    </div>
    <div class="ml-diagnostic ${hasSignal ? "ml-diagnostic--ok" : "ml-diagnostic--warn"}">
      <div class="ml-diagnostic-title">${hasSignal ? "🔝 Variable principale identifiée" : "⚠ Signal faible"}</div>
      ${hasSignal
        ? `La variable <strong>${escHtml(data.top_feature)}</strong> est la plus discriminante pour expliquer les gros écarts
           (Score F = ${fmt2(data.top_score)}).<br>
           <em>Recommandation :</em> concentrer l'analyse sur cette dimension en priorité.`
        : `Aucune variable très discriminante détectée (Score max = ${fmt2(data.top_score)}).
           Les gros écarts semblent aléatoires ou multifactoriels.`
      }
    </div>
  `;
}

/* ── LANCER L'ANALYSE ─────────────────────────────────────── */
async function lancerAnalyse() {
  const section = document.getElementById("ml-section").value;
  const flux    = document.getElementById("ml-flux").value;
  const annee   = document.getElementById("ml-annee").value;
  if (!section) return;

  showState("state-loading");
  document.getElementById("ml-title").textContent =
    _mode === "cluster" ? "Clustering des écarts" : "Analyse explicative";
  document.getElementById("ml-sub").textContent =
    `${section}${flux ? " · " + flux : ""}${annee ? " · " + annee : ""}`;

  const endpoint = _mode === "cluster" ? "/api/ml_ecarts/analyse" : "/api/ml_ecarts/explication";
  const body     = { section, flux, annee, algo: _algo };

  try {
    const res  = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);

    if (_mode === "cluster") {
      document.getElementById("ml-algo-badge").textContent = data.algo_info || "—";
      renderKPIs(data.global_stats);
      renderScatter(data.points, data.clusters);
      renderBarClusters(data.clusters);
      renderLegend(data.clusters);
      renderSummaryTable(data.summary, data.global_stats);
      renderDiagnostic(data.global_stats);
      showState("state-cluster");
    } else {
      document.getElementById("ml-algo-badge").textContent = "SelectKBest / f_classif";
      renderExplain(data);
      showState("state-explain");
    }
    const btnExcel = document.getElementById("btn-export-excel");
    if (btnExcel) {
      btnExcel.disabled = false;
      btnExcel._mlData = data;
    }
  } catch (err) {
    document.getElementById("ml-error-msg").textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
}

/* ── INIT ─────────────────────────────────────────────────── */
(async () => {
  document.getElementById("btn-lancer").addEventListener("click", lancerAnalyse);

  document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
    window.pulseChartPDF(null, "ML-Analyse-ecarts-PULSE");
  });

  document.getElementById("btn-export-excel")?.addEventListener("click", () => {
    const btnExcel = document.getElementById("btn-export-excel");
    const chart = _chartScatter || _chartBarCl || _chartExplain;
    if (chart) {
      window.pulseExcelChart(chart, "ml_ecarts_analyse");
    } else {
      window.toast?.("Lancez d'abord une analyse ML.", "error");
    }
  });

  document.getElementById("btn-reset-filters")?.addEventListener("click", () => {
    const selSection = document.getElementById("ml-section");
    const selFlux    = document.getElementById("ml-flux");
    if (selSection) selSection.selectedIndex = 0;
    if (selFlux)    { selFlux.innerHTML = '<option value="">Tous les flux</option>'; selFlux.disabled = true; }
    document.getElementById("btn-lancer").disabled = true;
    window.toast?.("Filtres réinitialisés", "info");
  });

  // Charger catalogue
  try {
    const res = await fetch("/api/catalogue");
    if (!res.ok) throw new Error();
    _catalogue = await res.json();

    const selSection = document.getElementById("ml-section");
    Object.keys(_catalogue).sort().forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      selSection.appendChild(o);
    });

    selSection.addEventListener("change", () => {
      const s = selSection.value;
      const selFlux  = document.getElementById("ml-flux");
      const selAnnee = document.getElementById("ml-annee");

      selFlux.innerHTML = `<option value="">Tous les flux</option>`;
      selAnnee.innerHTML = `<option value="">Toutes</option>`;

      if (!s) {
        selFlux.disabled  = true;
        selAnnee.disabled = true;
        document.getElementById("btn-lancer").disabled = true;
        return;
      }

      (_catalogue[s] || []).forEach(f => {
        const o = document.createElement("option");
        o.value = f; o.textContent = f;
        selFlux.appendChild(o);
      });
      selFlux.disabled  = false;
      selAnnee.disabled = false;
      document.getElementById("btn-lancer").disabled = false;

      // Charger les années depuis l'API accueil
      _loadAnnees(s);
    });

  } catch (_) {
    document.getElementById("ml-error-msg").textContent = "Impossible de charger le catalogue.";
    showState("state-error");
  }
})();

async function _loadAnnees(section) {
  const selAnnee = document.getElementById("ml-annee");
  selAnnee.innerHTML = `<option value="">Toutes</option>`;
  try {
    // On lit les années depuis une requête d'analyse légère (sans flux)
    const res  = await fetch(`/api/visualisation?section=${encodeURIComponent(section)}&flux=${encodeURIComponent((_catalogue[section] || [])[0] || "")}`);
    const data = await res.json().catch(() => ({}));
    (data.annees || []).forEach(y => {
      const o = document.createElement("option");
      o.value = String(y); o.textContent = String(y);
      selAnnee.appendChild(o);
    });
    // Présélectionner l'avant-dernière année (dernière complète)
    if ((data.annees || []).length) {
      const now  = new Date().getFullYear();
      const best = [...data.annees].filter(y => y < now).pop() ?? data.annees[data.annees.length - 1];
      selAnnee.value = String(best);
    }
  } catch (_) {}
}
