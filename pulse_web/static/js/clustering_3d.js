// surface_3d.js — PULSE · Surface 3D Performance (Enhanced)
// Filiale (Y) × Mois (X) → Écart performance (Z)
// Surface lissée cosinus + isolignes + projection sol + plan zéro + markers sommets + rotation auto

"use strict";

const BG         = "rgb(6,10,20)";
const BG_PLANE   = "rgba(14,24,50,0.7)";
const GRID_COLOR  = "rgba(255,255,255,0.09)";
const ZERO_COLOR  = "rgba(255,255,255,0.20)";
const TICK_COLOR  = "rgba(255,255,255,0.82)";
const LABEL_COLOR = "rgba(255,255,255,0.95)";

let currentData  = null;
let autoRotateId = null;
let rotAngle     = Math.PI * 0.55;   // angle initial caméra
let _rotFrame    = 0;

// ── Colorscales ──────────────────────────────────────────────────────────────

const CS_PERF = [
  [0.00, "rgb( 40,  0,  5)"],
  [0.08, "rgb(120, 10, 10)"],
  [0.18, "rgb(185, 22, 22)"],
  [0.30, "rgb(222, 62, 25)"],
  [0.43, "rgb(245,148,  8)"],
  [0.50, "rgb( 18, 28, 55)"],  // zéro — bleu nuit profond
  [0.57, "rgb(  8,108, 76)"],
  [0.70, "rgb( 10,172,108)"],
  [0.84, "rgb(  5,215,135)"],
  [1.00, "rgb(  0, 88, 56)"],
];

const CS_PRECISION = [
  [0.00, "rgb(100, 15, 15)"],
  [0.28, "rgb(210, 75, 30)"],
  [0.55, "rgb(240,175, 25)"],
  [0.80, "rgb( 40,200,140)"],
  [1.00, "rgb(  5,105, 65)"],
];

const CS_REEL = [
  [0.00, "rgb(  8, 15, 45)"],
  [0.20, "rgb( 15, 50,110)"],
  [0.45, "rgb( 25, 95,180)"],
  [0.70, "rgb( 60,155,230)"],
  [0.90, "rgb(120,200,255)"],
  [1.00, "rgb(186,230,253)"],
];

// ── Utils ─────────────────────────────────────────────────────────────────────

function fmtPct(v) {
  if (v == null) return "—";
  return (v >= 0 ? "+" : "") + v.toFixed(1) + "%";
}
function fmtKe(v) {
  if (v == null) return "—";
  return v.toLocaleString("fr-FR", { maximumFractionDigits: 0 }) + " k€";
}

// ── Auto-rotation cinématique ─────────────────────────────────────────────────

function _autoRotateStep() {
  _rotFrame++;
  if (_rotFrame % 2 === 0) {        // ~30 fps pour Plotly
    rotAngle += 0.006;              // ≈ 0.34°/frame → tour complet en ~17s
    const r = 1.88;
    Plotly.relayout("s3d-plotly", {
      "scene.camera.eye": {
        x: r * Math.cos(rotAngle),
        y: r * Math.sin(rotAngle) * -0.82,
        z: 0.95,
      }
    });
  }
  autoRotateId = requestAnimationFrame(_autoRotateStep);
}

function _stopRotate() {
  if (!autoRotateId) return;
  cancelAnimationFrame(autoRotateId);
  autoRotateId = null;
  const btn = document.getElementById("btn-rotate");
  if (btn) { btn.textContent = "▶ Rotation auto"; btn.dataset.active = "false"; }
}

function toggleRotate() {
  const btn = document.getElementById("btn-rotate");
  if (autoRotateId) {
    _stopRotate();
  } else {
    if (btn) { btn.textContent = "⏸ Arrêter"; btn.dataset.active = "true"; }
    _rotFrame = 0;
    _autoRotateStep();
  }
}

// ── Plein écran ───────────────────────────────────────────────────────────────

function toggleFullscreen() {
  const el  = document.querySelector(".s3d-canvas-area");
  const btn = document.getElementById("btn-fs");
  if (!document.fullscreenElement) {
    el.requestFullscreen().catch(console.error);
    if (btn) btn.textContent = "⊠ Quitter plein écran";
  } else {
    document.exitFullscreen();
    if (btn) btn.textContent = "⊞ Plein écran";
  }
}

document.addEventListener("fullscreenchange", () => {
  if (!document.fullscreenElement) {
    const btn = document.getElementById("btn-fs");
    if (btn) btn.textContent = "⊞ Plein écran";
  }
});

// ── Build traces + layout ────────────────────────────────────────────────────

function buildSurface(data, metrique) {
  const filiales  = data.filiales  || [];
  const x_vals    = data.x_vals    || (data.mois ? data.mois.map((_, i) => i) : []);
  const tick_vals = data.tick_vals || (data.mois ? data.mois.map((_, i) => i) : []);
  const tick_text = data.tick_text || data.mois || [];
  const z         = data.z         || [];

  // ── Colorscale & bornes Z
  let colorscale, colorbarTitle, zmin, zmax;
  if (metrique === "reel") {
    colorscale    = CS_REEL;
    colorbarTitle = "Volume réel (k€)";
  } else if (metrique === "precision") {
    colorscale    = CS_PRECISION;
    colorbarTitle = "Précision (%)";
    zmin = 0; zmax = 100;
  } else {
    colorscale    = CS_PERF;
    colorbarTitle = "Performance (%)";
    const flat = z.flatMap(row => row.filter(v => v != null));
    if (flat.length) {
      const absMax = flat.reduce((m, v) => Math.max(m, Math.abs(v)), 0);
      if (absMax > 0) { zmin = -absMax; zmax = absMax; }
    }
  }

  // ── Hover text
  const hovertext = filiales.map((filiale, fi) =>
    x_vals.map((xv, xi) => {
      const val       = (z[fi] || [])[xi];
      const moisIdx   = Math.min(Math.round(xv), tick_text.length - 1);
      const moisLabel = tick_text[moisIdx] || "";
      const valStr    = val == null ? "—"
        : metrique === "reel"      ? fmtKe(val)
        : metrique === "precision" ? val.toFixed(1) + " %"
        : fmtPct(val);
      const icon = val == null ? "" : (metrique === "perf" ? (val >= 0 ? "▲ " : "▼ ") : "");
      return `<b>${filiale}</b><br>${moisLabel}<br>${icon}${valStr}`;
    })
  );

  // ── TRACE 1 — Surface principale lissée avec isolignes
  const mainSurface = {
    type:     "surface",
    x:         x_vals,
    y:         filiales,
    z:         z,
    text:      hovertext,
    hovertemplate: "%{text}<extra></extra>",
    colorscale,
    reversescale: false,
    ...(zmin !== undefined ? { cmin: zmin, cmax: zmax } : {}),
    lighting: {
      ambient:   0.62,
      diffuse:   0.88,
      specular:  3.5,
      roughness: 0.06,   // quasi miroir → reflets spectaculaires
      fresnel:   1.8,
    },
    lightposition: { x: 200, y: 400, z: 600 },
    colorbar: {
      title:        colorbarTitle,
      titlefont:    { size: 11, color: "rgba(255,255,255,0.7)" },
      titleside:    "right",
      tickfont:     { size: 10, color: "rgba(255,255,255,0.65)" },
      bgcolor:      "rgba(0,0,0,0)",
      borderwidth:  0,
      outlinewidth: 0,
      thickness:    12,
      len:          0.60,
    },
    opacity: 0.97,
    // Isolignes sur la surface + projection en ombre sur le plancher
    contours: {
      z: {
        show:        true,
        usecolormap: true,
        width:       1.5,
        project:     { z: true },   // projette les contours sur le sol
      }
    },
  };

  const traces = [mainSurface];

  // ── TRACE 2 — Plan de référence translucide (zéro / 80% précision)
  if ((metrique === "perf" || metrique === "precision") && filiales.length && x_vals.length) {
    const zRef   = metrique === "precision" ? 80 : 0;
    const zPlane = filiales.map(() => x_vals.map(() => zRef));
    traces.push({
      type:       "surface",
      x:           x_vals,
      y:           filiales,
      z:           zPlane,
      showscale:   false,
      opacity:     0.14,
      colorscale:  [[0, "rgb(148,163,200)"], [1, "rgb(148,163,200)"]],
      hoverinfo:  "skip",
      lighting:   { ambient: 1.0, diffuse: 0.0, specular: 0.0, roughness: 1.0 },
      name:       metrique === "precision" ? "Seuil 80%" : "Équilibre 0%",
    });
  }

  // ── TRACE 3 — Markers sommets / creux par filiale (perf uniquement)
  if (metrique === "perf" && z.length) {
    const sx = [], sy = [], sz_m = [], st = [], sc = [];
    filiales.forEach((f, fi) => {
      const row = z[fi] || [];
      let maxV = -Infinity, maxI = 0, minV = Infinity, minI = 0;
      row.forEach((v, xi) => {
        if (v != null) {
          if (v > maxV) { maxV = v; maxI = xi; }
          if (v < minV) { minV = v; minI = xi; }
        }
      });
      const mI = (xv) => tick_text[Math.min(Math.round(xv || 0), tick_text.length - 1)] || "";
      if (isFinite(maxV) && maxV > 3) {
        sx.push(x_vals[maxI]); sy.push(f); sz_m.push(maxV);
        st.push(`<b>▲ ${f}</b><br>${mI(x_vals[maxI])} — meilleur<br>${fmtPct(maxV)}`);
        sc.push("#10B981");
      }
      if (isFinite(minV) && minV < -3) {
        sx.push(x_vals[minI]); sy.push(f); sz_m.push(minV);
        st.push(`<b>▼ ${f}</b><br>${mI(x_vals[minI])} — difficile<br>${fmtPct(minV)}`);
        sc.push("#EF4444");
      }
    });
    if (sx.length) {
      traces.push({
        type: "scatter3d",
        x: sx, y: sy, z: sz_m,
        mode: "markers",
        marker: {
          size:   6,
          color:  sc,
          symbol: "circle",
          line:   { color: "rgba(255,255,255,0.95)", width: 1.5 },
        },
        text:          st,
        hovertemplate: "%{text}<extra></extra>",
        showlegend:    false,
        name:          "Sommets",
      });
    }
  }

  // ── Axe de base
  const axisBase = {
    showbackground:  true,
    backgroundcolor: BG_PLANE,
    gridcolor:       GRID_COLOR,
    gridwidth:       1,
    zerolinecolor:   ZERO_COLOR,
    zerolinewidth:   2,
    showspikes:      true,
    spikecolor:      "rgba(255,255,255,0.30)",
    spikesides:      false,
    tickfont:        { size: 10, color: TICK_COLOR, family: "Inter, sans-serif" },
    titlefont:       { size: 12, color: LABEL_COLOR, family: "Inter, sans-serif" },
  };

  const layout = {
    paper_bgcolor: BG,
    plot_bgcolor:  BG,
    font:   { family: "Inter, sans-serif", color: TICK_COLOR, size: 11 },
    margin: { l: 0, r: 50, t: 10, b: 0 },

    scene: {
      bgcolor: BG,

      xaxis: {
        ...axisBase,
        title:    "Mois",
        tickmode: "array",
        tickvals:  tick_vals,
        ticktext:  tick_text,
        tickangle: 0,
      },
      yaxis: {
        ...axisBase,
        title:    "Filiale",
        tickfont: { size: 8.5, color: TICK_COLOR, family: "Inter, sans-serif" },
      },
      zaxis: {
        ...axisBase,
        title:    data.metrique_label || "Valeur",
        zeroline: true,
      },

      camera: {
        eye:    { x: 1.62, y: -1.42, z: 0.98 },
        center: { x: 0,    y:  0,    z: -0.05 },
        up:     { x: 0,    y:  0,    z:  1    },
      },
      aspectmode:  "manual",
      aspectratio: { x: 2.0, y: 1.0, z: 0.62 },
    },

    hoverlabel: {
      bgcolor:     "rgba(5,8,18,0.97)",
      bordercolor: "rgba(255,255,255,0.18)",
      font:        { family: "Inter, sans-serif", size: 12, color: "#E2E8F0" },
    },
  };

  const config = {
    responsive:             true,
    displaylogo:            false,
    modeBarButtonsToRemove: ["sendDataToCloud","select2d","lasso2d","autoScale2d"],
    toImageButtonOptions:   { format: "png", width: 1920, height: 1080, scale: 2, filename: "PULSE_Surface3D" },
  };

  return { traces, layout, config };
}

// ── Render ────────────────────────────────────────────────────────────────────

function renderChart(data) {
  const metrique = document.getElementById("f-metrique")?.value || "perf";
  const { traces, layout, config } = buildSurface(data, metrique);
  Plotly.react("s3d-plotly", traces, layout, config);
}

// ── KPIs ──────────────────────────────────────────────────────────────────────

function updateKPIs(data) {
  const { kpis, total } = data;
  const perf = kpis.perf_globale;
  document.getElementById("kpi-perf").textContent = (perf >= 0 ? "+" : "") + perf + "%";
  document.getElementById("kpi-perf").style.color  = perf >= 0 ? "#10B981" : "#EF4444";
  document.getElementById("kpi-total").textContent  = total.toLocaleString("fr-FR");
  document.getElementById("kpi-n-fil").textContent  = kpis.n_filiales;
  document.getElementById("kpi-best-fil").textContent  = kpis.best_filiale  || "—";
  document.getElementById("kpi-worst-fil").textContent = kpis.worst_filiale || "—";
  document.getElementById("kpi-best-mois").textContent  = kpis.best_mois  || "—";
  document.getElementById("kpi-worst-mois").textContent = kpis.worst_mois || "—";
}

// ── Chargement données ────────────────────────────────────────────────────────

async function loadData() {
  _stopRotate();

  const loadingEl = document.getElementById("s3d-loading");
  const emptyEl   = document.getElementById("s3d-empty");
  const chartEl   = document.getElementById("s3d-plotly");

  loadingEl.style.display = "flex";
  emptyEl.style.display   = "none";
  chartEl.style.opacity   = "0";

  const annee    = document.getElementById("f-annee")?.value    || "";
  const filiale  = document.getElementById("f-filiale")?.value  || "";
  const fluxType = document.getElementById("f-flux-type")?.value|| "";
  const metrique = document.getElementById("f-metrique")?.value || "perf";

  const params = new URLSearchParams({ metrique });
  if (annee)    params.set("annee",     annee);
  if (filiale)  params.set("filiale",   filiale);
  if (fluxType) params.set("flux_type", fluxType);

  try {
    const res  = await fetch(`/api/clustering_3d?${params}`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    if (!data.filiales?.length) {
      emptyEl.style.display   = "flex";
      loadingEl.style.display = "none";
      return;
    }

    currentData = data;

    // Peupler filtres dynamiques
    const fAnnee   = document.getElementById("f-annee");
    const fFiliale = document.getElementById("f-filiale");

    if (fAnnee.options.length <= 1 && data.annees?.length) {
      data.annees.forEach(a => {
        const o = document.createElement("option");
        o.value = o.textContent = a;
        fAnnee.appendChild(o);
      });
      fAnnee.value = data.annees[data.annees.length - 1];
      loadingEl.style.display = "none";
      return loadData();
    }

    if (fFiliale.options.length <= 1 && data.filiales?.length) {
      data.filiales.forEach(f => {
        const o = document.createElement("option");
        o.value = o.textContent = f;
        fFiliale.appendChild(o);
      });
    }

    updateKPIs(data);
    renderChart(data);
    requestAnimationFrame(() => { chartEl.style.opacity = "1"; });

  } catch (e) {
    console.error("[Surface3D]", e);
    document.querySelector(".s3d-empty-text").textContent = `Erreur : ${e.message}`;
    emptyEl.style.display = "flex";
  } finally {
    loadingEl.style.display = "none";
  }
}

// ── Init ──────────────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
  let _debounceTimer = null;
  const debouncedLoad = () => { clearTimeout(_debounceTimer); _debounceTimer = setTimeout(loadData, 350); };

  ["f-annee", "f-filiale", "f-flux-type"].forEach(id => {
    document.getElementById(id)?.addEventListener("change", debouncedLoad);
  });
  document.getElementById("f-metrique")?.addEventListener("change", () => {
    if (currentData) renderChart(currentData);
    debouncedLoad();
  });

  // Stopper la rotation sur interaction manuelle caméra
  document.getElementById("s3d-plotly")?.addEventListener("mousedown", () => {
    if (autoRotateId) _stopRotate();
  });
  document.getElementById("s3d-plotly")?.addEventListener("touchstart", () => {
    if (autoRotateId) _stopRotate();
  });

  loadData();
});
