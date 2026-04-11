/* visualisation.js */
"use strict";

/* ── PALETTE ──────────────────────────────────────────────── */
const PALETTE = [
  "#4C7CF3","#F59E0B","#10B981","#EF4444","#A78BFA",
  "#06B6D4","#EC4899","#84CC16","#F97316","#6366F1",
];

Chart.defaults.color       = "#64748B";
Chart.defaults.borderColor = "#1A2230";
Chart.defaults.font.family = "Inter, system-ui, sans-serif";
Chart.defaults.font.size   = 12;

/* ── ÉTAT GLOBAL ──────────────────────────────────────────── */
let _chart       = null;
let _data        = null;
let _chartType   = "line";
let _granularite = "mois";   // "jour" | "mois"
let _catalogue   = {};

/* ── CANVAS — toujours recréé ─────────────────────────────── */
function freshCanvas() {
  if (_chart) { try { _chart.destroy(); } catch (_) {} _chart = null; }
  const wrap = document.getElementById("chart-wrap");
  wrap.innerHTML = "";
  const cv = document.createElement("canvas");
  wrap.appendChild(cv);
  return cv.getContext("2d");
}

/* ── ZOOM ─────────────────────────────────────────────────── */
function _showResetZoom(show) {
  const btn = document.getElementById("btn-reset-zoom");
  if (!btn) return;
  btn.hidden = !show;
  btn.style.display = show ? "flex" : "none";
}
function resetZoom() {
  if (_chart) { _chart.resetZoom(); _showResetZoom(false); }
}
window.resetZoom = resetZoom;

/* ── ÉTATS ────────────────────────────────────────────────── */
function showState(id) {
  ["state-loading","state-error","state-empty","state-result"].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.hidden = (s !== id);
  });
}

/* ── TYPE DE GRAPHIQUE ────────────────────────────────────── */
function setChartType(type) {
  _chartType = type;
  document.getElementById("btn-line").classList.toggle("active", type === "line");
  document.getElementById("btn-bar" ).classList.toggle("active", type === "bar");
  document.getElementById("chart-type-badge").textContent =
    type === "line" ? "Linéaire" : "Cumulé mensuel";
  if (_data) render(_data);
}
window.setChartType = setChartType;

/* ── GRANULARITÉ ──────────────────────────────────────────── */
function setGranularite(g) {
  _granularite = g;
  document.getElementById("btn-jour").classList.toggle("active", g === "jour");
  document.getElementById("btn-mois").classList.toggle("active", g === "mois");
  if (_data) render(_data);
}
window.setGranularite = setGranularite;

/* ── AGRÉGATION MENSUELLE ─────────────────────────────────── */
function aggregateByMonth(dates, reel, profils) {
  const months = [], monthIdx = {};
  dates.forEach(iso => {
    const m = iso.slice(0, 7);
    if (!(m in monthIdx)) { monthIdx[m] = months.length; months.push(m); }
  });

  const reelSum  = new Array(months.length).fill(null);
  const reelCnt  = new Array(months.length).fill(0);
  dates.forEach((iso, i) => {
    const mi = monthIdx[iso.slice(0, 7)];
    if (reel[i] != null) { reelSum[mi] = (reelSum[mi] ?? 0) + reel[i]; reelCnt[mi]++; }
  });

  const aggProfils = profils.map(p => {
    const vals = new Array(months.length).fill(null);
    const cnt  = new Array(months.length).fill(0);
    dates.forEach((iso, i) => {
      const mi = monthIdx[iso.slice(0, 7)];
      if (p.valeurs[i] != null) { vals[mi] = (vals[mi] ?? 0) + p.valeurs[i]; cnt[mi]++; }
    });
    return { nom: p.nom, valeurs: vals };
  });

  // Dates représentatives : 1er du mois
  const aggDates = months.map(m => m + "-01");
  return [aggDates, reelSum, aggProfils];
}

/* ── FORMAT ───────────────────────────────────────────────── */
const fmt = v =>
  typeof v === "number" && isFinite(v)
    ? v.toLocaleString("fr-FR", { maximumFractionDigits: 0 })
    : "—";

function fmtLabel(iso) {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("fr-FR", { month: "short", year: "2-digit" });
}

/* ── FILTRE PÉRIODE ───────────────────────────────────────── */
function applyPeriod(dates, ...series) {
  const mS = parseInt(document.getElementById("f-mois-debut").value, 10);
  const mE = parseInt(document.getElementById("f-mois-fin"  ).value, 10);
  const mask = dates.map(iso => {
    const m = new Date(iso + "T00:00:00").getMonth() + 1;
    return mS <= mE ? (m >= mS && m <= mE) : (m >= mS || m <= mE);
  });
  return [
    dates.filter((_, i) => mask[i]),
    ...series.map(s => s.filter((_, i) => mask[i])),
  ];
}

/* ── UTILS ────────────────────────────────────────────────── */
function escHtml(s) {
  return String(s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function getSelectedProfils() {
  const s = new Set();
  document.querySelectorAll(".profil-chk:checked").forEach(c => s.add(c.value));
  return s;
}

/* ── PROFILS : construction de la liste ───────────────────── */
function buildProfilsList(profils) {
  const list    = document.getElementById("profils-list");
  const counter = document.getElementById("profils-count");
  list.innerHTML = "";

  const actifs = profils.filter(p => p.valeurs.some(v => v !== null));

  if (!actifs.length) {
    list.innerHTML = `<div class="profils-empty">Aucun profil pour cette sélection.</div>`;
    counter.textContent = "0 profil";
    return;
  }

  counter.textContent = `${actifs.length} profil${actifs.length > 1 ? "s" : ""}`;

  actifs.forEach((p, idx) => {
    const color = PALETTE[idx % PALETTE.length];
    const row   = document.createElement("label");
    row.className = "profil-item";
    row.innerHTML = `
      <input type="checkbox" class="profil-chk" value="${escHtml(p.nom)}">
      <span class="profil-color-bar" style="background:${color};"></span>
      <span class="profil-name" title="${escHtml(p.nom)}">${escHtml(p.nom)}</span>
    `;
    row.querySelector("input").addEventListener("change", () => {
      if (_data) render(_data);
    });
    list.appendChild(row);
  });
}

/* ── PROFILS : tout cocher / décocher ─────────────────────── */
function setAllProfils(checked) {
  document.querySelectorAll(".profil-chk").forEach(c => { c.checked = checked; });
  if (_data) render(_data);
}

/* ── BADGES TOPBAR ────────────────────────────────────────── */
function updateBadges() {
  const section = document.getElementById("f-section").value;
  const annee   = document.getElementById("f-annee"  ).value;
  const badges  = document.getElementById("chart-badges");
  badges.innerHTML = `<span class="chart-badge chart-badge--type" id="chart-type-badge">
    ${_chartType === "line" ? "Linéaire" : "Cumulé mensuel"}
  </span>`;
  if (section) badges.insertAdjacentHTML("afterbegin",
    `<span class="chart-badge chart-badge--section">${escHtml(section)}</span>`);
  if (annee) badges.insertAdjacentHTML("beforeend",
    `<span class="chart-badge chart-badge--annee">${escHtml(annee)}</span>`);
}

/* ── RENDER LINÉAIRE ──────────────────────────────────────── */
function renderLine(dates, reel, profils) {
  const showReel   = document.getElementById("chk-reel").checked;
  const selProfils = getSelectedProfils();
  const isMensuel  = _granularite === "mois";
  const labels     = dates.map(fmtLabel);
  const datasets   = [];

  if (showReel) {
    datasets.push({
      label: "Réel",
      data:  reel,
      borderColor:          "#FFFFFF",
      backgroundColor:      "rgba(255,255,255,0.08)",
      borderWidth:          isMensuel ? 2.5 : 1.8,
      pointRadius:          isMensuel ? 4   : 0,
      pointHoverRadius:     isMensuel ? 7   : 5,
      pointBackgroundColor: "#FFFFFF",
      tension: isMensuel ? 0.3 : 0.2,
      fill: false,
    });
  }

  const actifs = profils.filter(p => p.valeurs.some(v => v !== null));
  actifs.forEach((p, idx) => {
    if (!selProfils.has(p.nom)) return;
    const color = PALETTE[idx % PALETTE.length];
    datasets.push({
      label: p.nom,
      data:  p.valeurs,
      borderColor:          color,
      backgroundColor:      color + "18",
      pointBackgroundColor: color,
      borderWidth:          isMensuel ? 2   : 1.5,
      pointRadius:          isMensuel ? 3.5 : 0,
      pointHoverRadius:     isMensuel ? 6   : 5,
      tension: isMensuel ? 0.3 : 0.2,
      fill: false,
      spanGaps: true,
    });
  });

  if (!datasets.length) {
    showState("state-empty");
    document.getElementById("state-empty").querySelector("strong").textContent =
      "Aucune série sélectionnée";
    return;
  }

  const ctx = freshCanvas();
  showState("state-result");
  _showResetZoom(false);

  _chart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#CBD5E1",
            boxWidth: 12,
            boxHeight: 12,
            padding: 18,
            font: { size: 11, weight: "500" },
            usePointStyle: true,
            pointStyle: "circle",
          },
        },
        tooltip: {
          backgroundColor: "#1A2230",
          borderColor: "#2B3647",
          borderWidth: 1,
          titleColor: "#F3F4F6",
          bodyColor: "#94A3B8",
          padding: 12,
          callbacks: {
            label: ctx => ` ${ctx.dataset.label} : ${fmt(ctx.parsed.y)} k€`,
          },
        },
        zoom: {
          zoom: {
            wheel: { enabled: true, speed: 0.08 },
            pinch: { enabled: true },
            mode: "x",
            onZoom: () => _showResetZoom(true),
          },
          pan: {
            enabled: true,
            mode: "x",
            onPan: () => _showResetZoom(true),
          },
        },
      },
      scales: {
        x: {
          grid: { color: "rgba(255,255,255,0.04)", drawTicks: false },
          ticks: { color: "#475569", maxRotation: 40, padding: 6 },
          border: { color: "rgba(255,255,255,.08)" },
        },
        y: {
          grid: { color: "rgba(255,255,255,0.04)", drawTicks: false },
          ticks: { color: "#475569", padding: 10, callback: v => fmt(v) },
          border: { color: "rgba(255,255,255,.08)" },
        },
      },
    },
  });
}

/* ── RENDER CUMULÉ ────────────────────────────────────────── */
function renderBar(dates, reel, profils) {
  const showReel   = document.getElementById("chk-reel").checked;
  const selProfils = getSelectedProfils();

  const months = [], monthIdx = {};
  dates.forEach(iso => {
    const m = iso.slice(0, 7);
    if (!(m in monthIdx)) { monthIdx[m] = months.length; months.push(m); }
  });

  const reelSum = new Array(months.length).fill(0);
  dates.forEach((iso, i) => {
    const mi = monthIdx[iso.slice(0, 7)];
    if (reel[i] != null) reelSum[mi] += reel[i];
  });

  const labels   = months.map(m => {
    const [y, mo] = m.split("-");
    return new Date(+y, +mo - 1, 1).toLocaleDateString("fr-FR", { month: "short", year: "2-digit" });
  });

  const datasets = [];
  if (showReel) {
    datasets.push({
      label: "Réel",
      data: reelSum,
      backgroundColor: "rgba(255,255,255,.65)",
      borderRadius: 5,
      borderSkipped: false,
    });
  }

  const actifs = profils.filter(p => p.valeurs.some(v => v !== null));
  actifs.forEach((p, idx) => {
    if (!selProfils.has(p.nom)) return;
    const color = PALETTE[idx % PALETTE.length];

    // 1. Trouver le premier mois où ce profil a au moins une valeur
    const hasPrev = new Array(months.length).fill(false);
    dates.forEach((iso, i) => {
      if (p.valeurs[i] != null) hasPrev[monthIdx[iso.slice(0, 7)]] = true;
    });
    const firstMonth = hasPrev.findIndex(h => h);
    if (firstMonth === -1) return;

    // 2. Pour chaque mois >= firstMonth :
    //    - jours avec valeur profil → somme profil
    //    - jours sans valeur profil (démarrage en cours de mois) → somme réel
    //    Mois antérieurs au démarrage → null (barre absente)
    const combined = new Array(months.length).fill(null);
    dates.forEach((iso, i) => {
      const mi = monthIdx[iso.slice(0, 7)];
      if (mi < firstMonth) return;
      const val = p.valeurs[i] != null ? p.valeurs[i]
                : (reel[i] != null    ? reel[i] : 0);
      combined[mi] = (combined[mi] ?? 0) + val;
    });

    datasets.push({
      label: p.nom,
      data: combined,
      backgroundColor: color + "CC",
      borderRadius: 5,
      borderSkipped: false,
    });
  });

  if (!datasets.length) {
    showState("state-empty");
    return;
  }

  const ctx = freshCanvas();
  showState("state-result");
  _showResetZoom(false);

  _chart = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#CBD5E1", boxWidth: 12, boxHeight: 12,
            padding: 18, font: { size: 11, weight: "500" },
            usePointStyle: true, pointStyle: "rect",
          },
        },
        tooltip: {
          backgroundColor: "#1A2230", borderColor: "#2B3647", borderWidth: 1,
          titleColor: "#F3F4F6", bodyColor: "#94A3B8", padding: 12,
          callbacks: { label: ctx => ` ${ctx.dataset.label} : ${fmt(ctx.parsed.y)} k€` },
        },
        zoom: {
          zoom: {
            wheel: { enabled: true, speed: 0.08 },
            pinch: { enabled: true },
            mode: "x",
            onZoom: () => _showResetZoom(true),
          },
          pan: {
            enabled: true,
            mode: "x",
            onPan: () => _showResetZoom(true),
          },
        },
      },
      scales: {
        x: {
          grid: { color: "rgba(255,255,255,0.04)", drawTicks: false },
          ticks: { color: "#475569", maxRotation: 40, padding: 6 },
          border: { color: "rgba(255,255,255,.08)" },
        },
        y: {
          grid: { color: "rgba(255,255,255,0.04)", drawTicks: false },
          ticks: { color: "#475569", padding: 10, callback: v => fmt(v) },
          border: { color: "rgba(255,255,255,.08)" },
        },
      },
    },
  });
}

/* ── RENDER PRINCIPAL ─────────────────────────────────────── */
function render(data) {
  if (!data) return;

  const flux  = document.getElementById("f-flux" ).value;
  const annee = document.getElementById("f-annee").value;

  document.getElementById("chart-title"   ).textContent = flux  || "Graphique";
  document.getElementById("chart-subtitle").textContent = annee ? `Année ${annee}` : "Toutes années";
  updateBadges();

  const profilVals = data.profils.map(p => p.valeurs);
  const [fDates, fReel, ...fPV] = applyPeriod(data.dates, data.reel, ...profilVals);
  let fProfils = data.profils.map((p, i) => ({ nom: p.nom, valeurs: fPV[i] }));

  // Agrégation mensuelle (vue linéaire journalier non agrégeée, cumulé toujours par mois)
  let rDates = fDates, rReel = fReel, rProfils = fProfils;
  if (_granularite === "mois" && _chartType === "line") {
    [rDates, rReel, rProfils] = aggregateByMonth(fDates, fReel, fProfils);
  }

  if (_chartType === "line") renderLine(rDates, rReel, rProfils);
  else                       renderBar (fDates, fReel, fProfils);
}

/* ── FETCH ────────────────────────────────────────────────── */
async function fetchAndRender() {
  const section = document.getElementById("f-section").value;
  const flux    = document.getElementById("f-flux"   ).value;
  const annee   = document.getElementById("f-annee"  ).value;
  if (!section || !flux) return;

  showState("state-loading");
  const params = new URLSearchParams({ section, flux });
  if (annee) params.set("annee", annee);

  try {
    const res  = await fetch(`/api/visualisation?${params}`);
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
    _data = body;
    buildProfilsList(body.profils);   // profils décochés par défaut
    render(body);
  } catch (err) {
    if (_chart) { try { _chart.destroy(); } catch (_) {} _chart = null; }
    document.getElementById("error-msg").textContent = err.message || "Erreur inconnue";
    showState("state-error");
  }
}

/* ── CHANGEMENT DE FLUX ───────────────────────────────────── */
async function onFluxChange() {
  const section = document.getElementById("f-section").value;
  const flux    = document.getElementById("f-flux"   ).value;
  if (!section || !flux) return;

  const selAnnee = document.getElementById("f-annee");
  selAnnee.innerHTML = `<option value="">Toutes</option>`;
  selAnnee.disabled  = true;
  showState("state-loading");

  try {
    const res  = await fetch(`/api/visualisation?section=${encodeURIComponent(section)}&flux=${encodeURIComponent(flux)}`);
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);

    (body.annees || []).forEach(y => {
      const o = document.createElement("option");
      o.value = String(y); o.textContent = String(y);
      selAnnee.appendChild(o);
    });

    if (body.annees && body.annees.length) {
      const now  = new Date().getFullYear();
      const best = [...body.annees].filter(y => y < now).pop() ?? body.annees[body.annees.length - 1];
      selAnnee.value = String(best);
    }
    selAnnee.disabled = false;
    await fetchAndRender();
  } catch (err) {
    document.getElementById("error-msg").textContent = err.message;
    showState("state-error");
  }
}

/* ── INIT ─────────────────────────────────────────────────── */
(async () => {
  document.getElementById("f-mois-fin").value = "12";
  // Granularité par défaut : mensuel
  document.getElementById("btn-mois").classList.add("active");
  document.getElementById("btn-jour").classList.remove("active");

  document.getElementById("f-mois-debut").addEventListener("change", () => { if (_data) render(_data); });
  document.getElementById("f-mois-fin"  ).addEventListener("change", () => { if (_data) render(_data); });
  document.getElementById("chk-reel"    ).addEventListener("change", () => { if (_data) render(_data); });
  document.getElementById("f-annee"     ).addEventListener("change", fetchAndRender);
  document.getElementById("btn-afficher").addEventListener("click",  fetchAndRender);
  document.getElementById("btn-sel-all" ).addEventListener("click",  () => setAllProfils(true));
  document.getElementById("btn-sel-none").addEventListener("click",  () => setAllProfils(false));

  try {
    const res = await fetch("/api/catalogue");
    if (!res.ok) throw new Error("Catalogue indisponible");
    _catalogue = await res.json();

    const selSection = document.getElementById("f-section");
    Object.keys(_catalogue).sort().forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      selSection.appendChild(o);
    });

    selSection.addEventListener("change", () => {
      const section = selSection.value;
      const selFlux = document.getElementById("f-flux");
      const selAnn  = document.getElementById("f-annee");

      selFlux.innerHTML = "";
      selFlux.disabled  = !section;
      selAnn.innerHTML  = `<option value="">Toutes</option>`;
      selAnn.disabled   = true;
      _data = null;

      if (!section) {
        selFlux.innerHTML = `<option value="">— sélectionner une section —</option>`;
        document.getElementById("btn-afficher").disabled = true;
        showState("state-empty");
        return;
      }

      (_catalogue[section] || []).forEach(f => {
        const o = document.createElement("option");
        o.value = f; o.textContent = f;
        selFlux.appendChild(o);
      });
      document.getElementById("btn-afficher").disabled = false;
      onFluxChange();
    });

    document.getElementById("f-flux").addEventListener("change", onFluxChange);

  } catch (err) {
    document.getElementById("error-msg").textContent = err.message;
    showState("state-error");
  }
})();
