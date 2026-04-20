// ======================================================
// REPARTITION PAGE (Répartition par Filiale)
// ======================================================

// STATE
let chartBarInstance = null;
let chartDonutInstance = null;

// ======================================================
// HELPERS
// ======================================================
function formatNumber(n) {
  if (n === null || n === undefined) return '—';
  return n.toLocaleString('fr-FR', { maximumFractionDigits: 0 });
}

function formatPercent(n) {
  if (n === null || n === undefined) return '—';
  return n.toLocaleString('fr-FR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

// ======================================================
// STATE MANAGEMENT
// ======================================================
function showLoading() {
  document.getElementById('state-loading').removeAttribute('hidden');
  document.getElementById('state-result').setAttribute('hidden', '');
}

function showResult() {
  document.getElementById('state-loading').setAttribute('hidden', '');
  document.getElementById('state-result').removeAttribute('hidden');
}

// ======================================================
// LOAD AND RENDER
// ======================================================
async function updateData() {
  showLoading();

  try {
    const annee = document.getElementById('f-annee').value || '';
    const params = new URLSearchParams();
    if (annee) {
      params.append('annee', annee);
    }

    const res = await fetch(`/api/repartition?${params}`);
    if (!res.ok) throw new Error(`API error: ${res.status}`);

    const data = await res.json();

    renderKPIs(data);
    renderCharts(data);
    renderTable(data);
    showResult();

    /* Activer export Excel */
    const btnExcel = document.getElementById("btn-export-excel");
    if (btnExcel) {
      btnExcel.disabled = false;
      btnExcel._repartData = data;
    }
  } catch (error) {
    console.error('Error loading data:', error);
    document.getElementById('state-result').innerHTML = `
      <div style="padding: 20px; color: #ef4444; text-align: center;">
        ⚠️ Erreur lors du chargement des données
      </div>
    `;
    showResult();
  }
}

// ======================================================
// RENDER KPIs
// ======================================================
function renderKPIs(data) {
  const kpiRow = document.getElementById('kpi-row');
  kpiRow.innerHTML = '';

  // KPI 1: Total écarts
  const kpi1 = document.createElement('div');
  kpi1.className = 'kpi-card';
  kpi1.innerHTML = `
    <div class="kpi-value">${formatNumber(data.total_ecarts)}</div>
    <div class="kpi-label">Total écarts</div>
  `;
  kpiRow.appendChild(kpi1);

  // KPI 2: Nombre de filiales
  const kpi2 = document.createElement('div');
  kpi2.className = 'kpi-card';
  kpi2.innerHTML = `
    <div class="kpi-value">${formatNumber(data.nb_filiales)}</div>
    <div class="kpi-label">Filiales</div>
  `;
  kpiRow.appendChild(kpi2);

  // KPI 3: Écarts favorables
  const totalFav = data.par_filiale.reduce((sum, f) => sum + f.nb_favorables, 0);
  const pctFav = data.total_ecarts > 0 ? (totalFav / data.total_ecarts * 100) : 0;
  const kpi3 = document.createElement('div');
  kpi3.className = 'kpi-card';
  kpi3.innerHTML = `
    <div class="kpi-value" style="color: #1f9d63;">${formatNumber(totalFav)}</div>
    <div class="kpi-label">${formatPercent(pctFav)}% Favorables</div>
  `;
  kpiRow.appendChild(kpi3);

  // KPI 4: Écarts défavorables
  const totalDef = data.par_filiale.reduce((sum, f) => sum + f.nb_defavorables, 0);
  const pctDef = data.total_ecarts > 0 ? (totalDef / data.total_ecarts * 100) : 0;
  const kpi4 = document.createElement('div');
  kpi4.className = 'kpi-card';
  kpi4.innerHTML = `
    <div class="kpi-value" style="color: #ef4444;">${formatNumber(totalDef)}</div>
    <div class="kpi-label">${formatPercent(pctDef)}% Défavorables</div>
  `;
  kpiRow.appendChild(kpi4);
}

// ======================================================
// RENDER CHARTS
// ======================================================
function renderCharts(data) {
  renderBarChart(data);
  renderDonutChart(data);
}

function renderBarChart(data) {
  // Prepare data per filiale with fav/def split
  const labels = data.par_filiale.map(f => f.filiale);
  const favData = data.par_filiale.map(f => f.nb_favorables);
  const defData = data.par_filiale.map(f => f.nb_defavorables);

  const ctx = document.getElementById('chart-bar').getContext('2d');

  if (chartBarInstance) {
    chartBarInstance.destroy();
  }

  chartBarInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Favorables',
          data: favData,
          backgroundColor: 'rgba(31, 157, 99, 0.8)',
          borderRadius: 6,
          borderSkipped: false
        },
        {
          label: 'Défavorables',
          data: defData,
          backgroundColor: 'rgba(239, 68, 68, 0.8)',
          borderRadius: 6,
          borderSkipped: false
        }
      ]
    },
    options: {
      indexAxis: 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          labels: {
            color: '#d1d5db',
            boxWidth: 12,
            padding: 15,
            font: { size: 12 }
          }
        },
        tooltip: {
          backgroundColor: 'rgba(0, 0, 0, 0.8)',
          padding: 12,
          titleColor: '#f3f4f6',
          bodyColor: '#d1d5db',
          borderColor: '#2b3647',
          borderWidth: 1
        }
      },
      scales: {
        x: {
          stacked: true,
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Filiales", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          grid: { display: false }
        },
        y: {
          stacked: true,
          ticks: { color: "#FFFFFF", font: { weight: "500" } },
          title: { display: true, text: "Nombre d'\u00e9carts", color: "#FFFFFF", font: { size: 11, weight: "500" } },
          grid: { color: 'rgba(139, 148, 168, 0.1)' },
          beginAtZero: true
        }
      }
    }
  });
}

function renderDonutChart(data) {
  const totalFav = data.par_filiale.reduce((sum, f) => sum + f.nb_favorables, 0);
  const totalDef = data.par_filiale.reduce((sum, f) => sum + f.nb_defavorables, 0);

  const ctx = document.getElementById('chart-donut').getContext('2d');

  if (chartDonutInstance) {
    chartDonutInstance.destroy();
  }

  chartDonutInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Favorables', 'Défavorables'],
      datasets: [{
        data: [totalFav, totalDef],
        backgroundColor: [
          'rgba(31, 157, 99, 0.8)',
          'rgba(239, 68, 68, 0.8)'
        ],
        borderColor: '#0B0F17',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          labels: {
            color: '#d1d5db',
            boxWidth: 12,
            padding: 15,
            font: { size: 12 }
          }
        },
        tooltip: {
          backgroundColor: 'rgba(0, 0, 0, 0.8)',
          padding: 12,
          titleColor: '#f3f4f6',
          bodyColor: '#d1d5db',
          borderColor: '#2b3647',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return `${ctx.label}: ${formatNumber(ctx.parsed)} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

// ======================================================
// RENDER TABLE
// ======================================================
function renderTable(data) {
  const tbody = document.getElementById('tbody-repartition');
  tbody.innerHTML = '';

  if (!data.par_filiale || data.par_filiale.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" style="text-align: center; color: #7c8798; padding: 20px;">
          ⚠️ Aucune donnée disponible
        </td>
      </tr>
    `;
    return;
  }

  data.par_filiale.forEach((row) => {
    const tr = document.createElement('tr');

    tr.innerHTML = `
      <td>${row.filiale}</td>
      <td class="num">${formatNumber(row.nb_ecarts)}</td>
      <td class="num" style="color: #1f9d63;">${formatNumber(row.nb_favorables)}</td>
      <td class="num" style="color: #ef4444;">${formatNumber(row.nb_defavorables)}</td>
      <td class="num">${formatPercent(row.pct_favorables)}%</td>
      <td class="num">${formatPercent(row.ecart_moy_pct)}%</td>
      <td class="num">${formatPercent(row.ecart_max_pct)}%</td>
    `;

    tbody.appendChild(tr);
  });
}

// ======================================================
// LOAD YEARS
// ======================================================
async function loadYears() {
  try {
    const res = await fetch('/api/accueil');
    const data = await res.json();

    if (data.annees && Array.isArray(data.annees)) {
      const select = document.getElementById('f-annee');
      const currentValue = select.value;
      
      // Add year options after "Toutes"
      data.annees.forEach(year => {
        const opt = document.createElement('option');
        opt.value = year;
        opt.textContent = year;
        select.appendChild(opt);
      });
      
      // Restore or set default
      if (!currentValue && data.annees.length > 0) {
        select.value = data.annees[0];
      }
    }
  } catch (error) {
    console.error('Error loading years:', error);
  }
}

// ======================================================
// INIT + EVENT LISTENERS
// ======================================================
document.addEventListener('DOMContentLoaded', () => {
  // Load years dropdown
  loadYears();

  // Listeners
  document.getElementById('f-annee').addEventListener('change', updateData);
  document.getElementById('btn-analyser').addEventListener('click', updateData);

  // Export PDF
  document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
    window.pulseChartPDF(null, "Repartition-filiale-PULSE");
  });

  // Export Excel
  document.getElementById("btn-export-excel")?.addEventListener("click", () => {
    const btnExcel = document.getElementById("btn-export-excel");
    if (chartBarInstance) {
      window.pulseExcelChart(chartBarInstance, "repartition_filiale");
    } else if (btnExcel?._repartData) {
      const d = btnExcel._repartData;
      const headers = ["Section", "Favorables", "Défavorables", "Total"];
      const rows = (d.sections || []).map((s, i) => [
        s,
        d.favorables?.[i] ?? 0,
        d.defavorables?.[i] ?? 0,
        (d.favorables?.[i] ?? 0) + (d.defavorables?.[i] ?? 0),
      ]);
      window.pulseExcelData(headers, rows, "repartition_filiale");
    } else {
      window.toast?.("Aucune donnée à exporter.", "error");
    }
  });

  // Reset filtres
  document.getElementById("btn-reset-filters")?.addEventListener("click", () => {
    document.getElementById("f-annee").selectedIndex = 0;
    updateData();
    window.toast?.("Filtres réinitialisés", "info");
  });

  // Initial load
  updateData();
});
