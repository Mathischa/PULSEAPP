// ======================================================
// PREVISION REPARTITION PAGE
// ======================================================

// STATE
let chartRateInstance = null;
let chartValueInstance = null;

// ======================================================
// HELPERS
// ======================================================
function showLoading(container) {
  container.innerHTML = `
    <div class="chart-placeholder">
      <div class="loading-state">
        <div class="loading-spinner"></div>
        Chargement des données...
      </div>
    </div>
  `;
}

function showEmpty(container) {
  container.innerHTML = `
    <div class="chart-placeholder">
      <div style="color: #7c8798; font-size: 14px;">
        ⚠️ Aucune donnée disponible pour ce filtre.
      </div>
    </div>
  `;
}

function formatNumber(n) {
  return Math.round(n).toLocaleString("fr-FR");
}

// ======================================================
// FETCH CONFIG
// ======================================================
async function loadConfig() {
  try {
    const res = await fetch('/api/prevision_repartition/config');
    const data = await res.json();

    // Populate filters
    const filialeSelect = document.getElementById('filterFiliale');
    const anneeSelect = document.getElementById('filterAnnee');
    const fluxSelect = document.getElementById('filterFlux');

    filialeSelect.innerHTML = data.filiales.map(f =>
      `<option ${f === 'Toutes filiales' ? 'selected' : ''}>${f}</option>`
    ).join('');

    anneeSelect.innerHTML = data.annees.map(a =>
      `<option ${a === data.annees[0] ? 'selected' : ''}>${a}</option>`
    ).join('');

    fluxSelect.innerHTML = data.flux.map(fl =>
      `<option ${fl === 'Tous les flux' ? 'selected' : ''}>${fl}</option>`
    ).join('');

    // Add listeners
    filialeSelect.addEventListener('change', updateCharts);
    anneeSelect.addEventListener('change', updateCharts);
    fluxSelect.addEventListener('change', updateCharts);

    document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
      window.pulsePDF("Répartition des écarts par profil — PULSE");
    });

    document.getElementById("btn-export-excel")?.addEventListener("click", () => {
      const chart = chartRateInstance || chartValueInstance;
      if (chart) {
        window.pulseExcelChart(chart, "prevision_repartition");
      } else {
        alert("Aucun graphique disponible.");
      }
    });

    // Initial load
    await updateCharts();
  } catch (error) {
    console.error('Error loading config:', error);
  }
}

// ======================================================
// FETCH DATA + RENDER
// ======================================================
async function updateCharts() {
  const filiale = document.getElementById('filterFiliale').value;
  const annee = document.getElementById('filterAnnee').value;
  const flux = document.getElementById('filterFlux').value;

  // Show loading
  showLoading(document.getElementById('chart1Container'));
  showLoading(document.getElementById('chart2Container'));
  document.getElementById('tableBody').innerHTML = `
    <tr><td colspan="5" class="no-data">Chargement...</td></tr>
  `;

  try {
    const params = new URLSearchParams({
      filiale: filiale,
      annee: annee,
      flux: flux
    });

    const res = await fetch(`/api/prevision_repartition?${params}`);
    const data = await res.json();

    if (data.empty) {
      showEmpty(document.getElementById('chart1Container'));
      showEmpty(document.getElementById('chart2Container'));
      document.getElementById('tableBody').innerHTML = `
        <tr><td colspan="5" class="no-data">Aucune donnée pour ce filtre</td></tr>
      `;
      return;
    }

    renderCharts(data, filiale, annee);
    renderTable(data);
    const btnExcel = document.getElementById("btn-export-excel");
    if (btnExcel) btnExcel.disabled = false;
  } catch (error) {
    console.error('Error updating charts:', error);
    document.getElementById('tableBody').innerHTML = `
      <tr><td colspan="5" class="no-data">Erreur lors du chargement</td></tr>
    `;
  }
}

// ======================================================
// RENDER CHARTS
// ======================================================
function renderCharts(data, filiale, annee) {
  const titleFiliale = filiale;
  const titleSuffix = annee && annee !== 'Toutes années' ? ` — ${annee}` : '';

  // ===== CHART 1: TAUX =====
  const container1 = document.getElementById('chart1Container');
  const canvas1 = document.createElement('canvas');
  container1.innerHTML = '';
  container1.appendChild(canvas1);

  const ctx1 = canvas1.getContext('2d');
  const colors1 = data.taux.map(pct => {
    const ratio = Math.min(pct / Math.max(...data.taux, 1), 1);
    return `rgba(76, 124, 243, ${0.4 + ratio * 0.6})`;
  });

  if (chartRateInstance) {
    chartRateInstance.destroy();
  }

  chartRateInstance = new Chart(ctx1, {
    type: 'bar',
    data: {
      labels: data.profils,
      datasets: [{
        label: 'Taux (%)',
        data: data.taux,
        backgroundColor: colors1,
        borderRadius: 6,
        borderSkipped: false
      }]
    },
    options: {
      indexAxis: 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: {
          display: true,
          text: `${titleFiliale} — Taux d'écarts (écarts/prévisions)${titleSuffix}`,
          color: '#f3f4f6',
          font: { size: 14, weight: 'bold' }
        },
        tooltip: {
          backgroundColor: 'rgba(0, 0, 0, 0.8)',
          padding: 12,
          titleColor: '#f3f4f6',
          bodyColor: '#d1d5db',
          borderColor: '#2b3647',
          borderWidth: 1,
          corners: 4,
          callbacks: {
            label: (ctx) => `${ctx.parsed.y.toFixed(2)}%`
          }
        }
      },
      scales: {
        y: {
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Taux d'\u00e9carts (%)", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          grid: { color: "rgba(139, 148, 168, 0.1)" },
          beginAtZero: true
        },
        x: {
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Profils", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          grid: { display: false }
        }
      }
    }
  });

  // ===== CHART 2: VALORISATION =====
  const container2 = document.getElementById('chart2Container');
  const canvas2 = document.createElement('canvas');
  container2.innerHTML = '';
  container2.appendChild(canvas2);

  const ctx2 = canvas2.getContext('2d');
  const maxV = Math.max(...data.valorisation, 1);
  const colors2 = data.valorisation.map(val => {
    const ratio = Math.min(val / maxV, 1);
    return `rgba(31, 157, 99, ${0.4 + ratio * 0.6})`;
  });

  if (chartValueInstance) {
    chartValueInstance.destroy();
  }

  chartValueInstance = new Chart(ctx2, {
    type: 'bar',
    data: {
      labels: data.profils,
      datasets: [{
        label: 'Valorisation (k€)',
        data: data.valorisation,
        backgroundColor: colors2,
        borderRadius: 6,
        borderSkipped: false
      }]
    },
    options: {
      indexAxis: 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: {
          display: true,
          text: `${titleFiliale} — Valorisation des écarts${titleSuffix}`,
          color: '#f3f4f6',
          font: { size: 14, weight: 'bold' }
        },
        tooltip: {
          backgroundColor: 'rgba(0, 0, 0, 0.8)',
          padding: 12,
          titleColor: '#f3f4f6',
          bodyColor: '#d1d5db',
          borderColor: '#2b3647',
          borderWidth: 1,
          corners: 4,
          callbacks: {
            label: (ctx) => `${formatNumber(ctx.parsed.y)} k€`
          }
        }
      },
      scales: {
        y: {
          ticks: {
            color: "#FFFFFF",
            callback: (val) => formatNumber(val),
            font: { weight: "500" }
          },
          title: { display: true, text: "Valorisation (k€)", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          grid: { color: "rgba(139, 148, 168, 0.1)" },
          beginAtZero: true
        },
        x: {
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Profils", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          grid: { display: false }
        }
      }
    }
  });
}

// ======================================================
// RENDER TABLE
// ======================================================
function renderTable(data) {
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';

  if (!data.table || data.table.length === 0) {
    tbody.innerHTML = `
      <tr><td colspan="5" class="no-data">Aucune donnée</td></tr>
    `;
    return;
  }

  data.table.forEach((row, idx) => {
    const isTotalRow = row[0] === 'TOTAL';
    const tr = document.createElement('tr');
    tr.className = isTotalRow ? 'total-row' : '';

    const [profil, nbPrev, nbEcarts, taux, valo] = row;

    tr.innerHTML = `
      <td>${profil}</td>
      <td>${nbPrev}</td>
      <td>${nbEcarts}</td>
      <td>${taux.toFixed(2)}%</td>
      <td>${formatNumber(valo)}</td>
    `;

    tbody.appendChild(tr);
  });
}

// ======================================================
// INIT
// ======================================================
document.addEventListener('DOMContentLoaded', loadConfig);
