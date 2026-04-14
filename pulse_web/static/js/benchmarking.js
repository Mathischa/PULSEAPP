// benchmarking.js — Matrice Réel/Prévision + Radar Performance

"use strict";

document.addEventListener("DOMContentLoaded", async () => {
  const chartsContainer = document.getElementById("charts-container");
  const loadingState = document.getElementById("loading");
  const perfTbody = document.getElementById("perf-tbody");

  let scatterChart = null;
  let radarChart = null;

  // Couleurs par filiale
  const COLORS = {
    "ACE A": "#3B82F6",
    "ACE B": "#10B981",
    "ACE C": "#F59E0B",
    "ACE D": "#EF4444",
  };

  // =========================================================
  // FETCH DATA
  // =========================================================
  async function fetchData(filters = {}) {
    try {
      const params = new URLSearchParams(filters);
      const res = await fetch(`/api/benchmarking?${params}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.json();
    } catch (e) {
      console.error("Erreur fetch benchmarking:", e);
      alert("Erreur lors du chargement des données.");
      return null;
    }
  }

  // =========================================================
  // SCATTER PLOT: Réel vs Prévision
  // =========================================================
  function initScatterChart(data) {
    const ctx = document.getElementById("scatter-chart").getContext("2d");
    
    if (scatterChart) scatterChart.destroy();

    // Préparer les datasets par filiale
    const filialeData = {};
    const allPoints = [];

    data.forEach((point) => {
      const filiale = point.filiale || "Autre";
      if (!filialeData[filiale]) {
        filialeData[filiale] = [];
      }
      filialeData[filiale].push(point);
      allPoints.push(point);
    });

    // Calculer régression linéaire
    function linearRegression(points) {
      if (points.length < 2) return null;
      const n = points.length;
      const sumX = points.reduce((s, p) => s + p.x, 0);
      const sumY = points.reduce((s, p) => s + p.y, 0);
      const sumXY = points.reduce((s, p) => s + p.x * p.y, 0);
      const sumX2 = points.reduce((s, p) => s + p.x * p.x, 0);

      const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
      const intercept = (sumY - slope * sumX) / n;

      // Points pour la ligne
      const xMin = Math.min(...points.map((p) => p.x));
      const xMax = Math.max(...points.map((p) => p.x));
      return [
        { x: xMin, y: slope * xMin + intercept },
        { x: xMax, y: slope * xMax + intercept },
      ];
    }

    const datasets = [];
    let datasetIdx = 0;

    Object.keys(filialeData).forEach((filiale) => {
      const points = filialeData[filiale];
      const color = COLORS[filiale] || "#666";

      // Points
      datasets.push({
        label: filiale,
        data: points.map((p) => ({
          x: p.prevision,
          y: p.reel,
          r: Math.sqrt(Math.abs(p.ecart)) + 3, // taille = écart
        })),
        backgroundColor: color + "55",
        borderColor: color,
        borderWidth: 1.5,
        type: "bubble",
      });

      // Ligne régression
      const regressionLine = linearRegression(points);
      if (regressionLine) {
        datasets.push({
          label: `Régression ${filiale}`,
          data: regressionLine,
          type: "line",
          borderColor: color,
          borderWidth: 2,
          borderDash: [5, 5],
          fill: false,
          pointRadius: 0,
          showLine: true,
          tension: 0,
        });
      }

      datasetIdx++;
    });

    // Ligne parfaite (Réel = Prévision)
    const maxVal = Math.max(...allPoints.map((p) => Math.max(p.x, p.y)));
    datasets.push({
      label: "Parfait (Réel = Prévision)",
      data: [
        { x: 0, y: 0 },
        { x: maxVal, y: maxVal },
      ],
      type: "line",
      borderColor: "#999",
      borderWidth: 1,
      borderDash: [3, 3],
      fill: false,
      pointRadius: 0,
      showLine: true,
      tension: 0,
    });

    scatterChart = new Chart(ctx, {
      type: "scatter",
      data: { datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            labels: {
              color: "#F1F5F9",
              font: { size: 11 },
            },
          },
          tooltip: {
            backgroundColor: "rgba(0,0,0,0.8)",
            titleColor: "#fff",
            bodyColor: "#fff",
            callbacks: {
              label: function (context) {
                const point = context.raw;
                const ecart =
                  ((point.y - point.x) / point.x) * 100 || 0;
                return `Prévision: ${point.x.toFixed(0)} k€ | Réel: ${point.y.toFixed(
                  0
                )} k€ | Écart: ${ecart.toFixed(1)}%`;
              },
            },
          },
        },
        scales: {
          x: {
            type: "linear",
            position: "bottom",
            title: {
              display: true,
              text: "Prévision (k€)",
              color: "#CBD5E1",
            },
            ticks: { color: "#CBD5E1" },
            grid: { color: "rgba(100,100,100,0.1)" },
          },
          y: {
            title: {
              display: true,
              text: "Réel (k€)",
              color: "#CBD5E1",
            },
            ticks: { color: "#CBD5E1" },
            grid: { color: "rgba(100,100,100,0.1)" },
          },
        },
      },
    });
  }

  // =========================================================
  // RADAR: Performance Filiales
  // =========================================================
  function initRadarChart(data) {
    const ctx = document.getElementById("radar-chart").getContext("2d");

    if (radarChart) radarChart.destroy();

    // Calculer KPIs par filiale
    const filialeMetrics = {};

    data.forEach((point) => {
      const filiale = point.filiale || "Autre";
      if (!filialeMetrics[filiale]) {
        filialeMetrics[filiale] = {
          ecarts: [],
          favorables: 0,
          total: 0,
          volatility: [],
        };
      }

      const m = filialeMetrics[filiale];
      m.ecarts.push(Math.abs(point.ecart_pct));
      m.volatility.push(point.reel);
      if (point.favorable) m.favorables++;
      m.total++;
    });

    // Calculer les scores
    const datasets = [];
    const labels = ["Précision", "Volatilité", "Favorabilité", "Stabilité"];
    const filiales = Object.keys(filialeMetrics);

    filiales.forEach((filiale) => {
      const m = filialeMetrics[filiale];

      // Précision: 100 - écart moyen %
      const precision = Math.max(
        0,
        100 - (m.ecarts.reduce((a, b) => a + b, 0) / m.ecarts.length || 0)
      );

      // Volatilité: Inverse du CV (Coefficient de Variation)
      const mean = m.volatility.reduce((a, b) => a + b, 0) / m.volatility.length;
      const variance =
        m.volatility.reduce((s, v) => s + Math.pow(v - mean, 2), 0) /
        m.volatility.length;
      const cv = Math.sqrt(variance) / (mean || 1);
      const volatilite = Math.max(0, 100 - cv * 50); // Inverser: moins de volatilité = mieux

      // Favorabilité: % écarts favorables
      const favorabilite = (m.favorables / m.total) * 100;

      // Stabilité: 100 - écart-type des écarts
      const ecartMean = m.ecarts.reduce((a, b) => a + b, 0) / m.ecarts.length;
      const ecartVariance =
        m.ecarts.reduce((s, e) => s + Math.pow(e - ecartMean, 2), 0) /
        m.ecarts.length;
      const stabilite = Math.max(0, 100 - Math.sqrt(ecartVariance));

      const scores = [precision, volatilite, favorabilite, stabilite];

      datasets.push({
        label: filiale,
        data: scores,
        borderColor: COLORS[filiale] || "#666",
        backgroundColor: (COLORS[filiale] || "#666") + "22",
        borderWidth: 2,
        fill: true,
        tension: 0.4,
        pointRadius: 4,
        pointHoverRadius: 6,
      });

      // Ajouter ligne perf tableau
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><span style="color:${COLORS[filiale] || "#666"}">●</span> ${filiale}</td>
        <td>${precision.toFixed(1)}%</td>
        <td>${volatilite.toFixed(1)}%</td>
        <td>${favorabilite.toFixed(1)}%</td>
        <td>${stabilite.toFixed(1)}%</td>
      `;
      perfTbody.appendChild(tr);
    });

    radarChart = new Chart(ctx, {
      type: "radar",
      data: {
        labels,
        datasets,
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            labels: {
              color: "#F1F5F9",
              font: { size: 12 },
              usePointStyle: true,
            },
          },
          tooltip: {
            backgroundColor: "rgba(0,0,0,0.8)",
            titleColor: "#fff",
            bodyColor: "#fff",
            callbacks: {
              label: function (context) {
                return context.dataset.label + ": " + context.parsed.r.toFixed(1) + "%";
              },
            },
          },
        },
        scales: {
          r: {
            beginAtZero: true,
            max: 100,
            ticks: {
              color: "#CBD5E1",
              stepSize: 20,
            },
            grid: { color: "rgba(100,100,100,0.2)" },
            pointLabels: {
              color: "#CBD5E1",
              font: { size: 12, weight: "bold" },
            },
          },
        },
      },
    });
  }

  // =========================================================
  // LOAD & RENDER
  // =========================================================
  async function loadAndRender(filters = {}) {
    loadingState.hidden = false;
    chartsContainer.hidden = true;
    perfTbody.innerHTML = "";

    const data = await fetchData(filters);
    if (!data || data.length === 0) {
      loadingState.innerHTML = "Aucune donnée disponible.";
      return;
    }

    initScatterChart(data);
    initRadarChart(data);

    loadingState.hidden = true;
    chartsContainer.hidden = false;
  }

  // =========================================================
  // FILTRES
  // =========================================================
  const filterAnnee = document.getElementById("f-annee");
  const filterFiliale = document.getElementById("f-filiale");
  const filterFluxType = document.getElementById("f-flux-type");

  async function applyFilters() {
    const filters = {
      annee: filterAnnee.value || "",
      filiale: filterFiliale.value || "",
      flux_type: filterFluxType.value || "",
    };
    await loadAndRender(filters);
  }

  filterAnnee.addEventListener("change", applyFilters);
  filterFiliale.addEventListener("change", applyFilters);
  filterFluxType.addEventListener("change", applyFilters);

  // Initial load
  await loadAndRender();

  // Export PDF
  document.getElementById("btn-export-pdf").addEventListener("click", () => {
    window.pulsePDF("Benchmarking-Comparatif");
  });
});
