/* accueil.js — Tableau de bord PULSE — Hub d'aide à la décision */
"use strict";

/* =========================================================
   HELPERS
   ========================================================= */
function qs(id) { return document.getElementById(id); }

function setText(id, value, fallback = "—") {
  const el = qs(id);
  if (el) el.textContent = value ?? fallback;
}

function setHTML(id, html) {
  const el = qs(id);
  if (el) el.innerHTML = html;
}

function formatNumber(value) {
  if (typeof value !== "number" || Number.isNaN(value)) return "—";
  return value.toLocaleString("fr-FR");
}

function formatKo(value) {
  if (typeof value !== "number" || Number.isNaN(value)) return "—";
  return `${value.toLocaleString("fr-FR")} Ko`;
}

function plural(count, singular, pluralForm = null) {
  return count > 1 ? (pluralForm || `${singular}s`) : singular;
}

/* =========================================================
   ANIMATION — Count-up fluide
   ========================================================= */
function animateCount(el, target, suffix = "", duration = 900) {
  if (!el) return;
  if (typeof target !== "number" || Number.isNaN(target)) {
    el.textContent = `${target ?? "—"}${suffix}`;
    return;
  }
  const start = performance.now();
  const step = (now) => {
    const t    = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - t, 3);
    el.textContent = Math.round(ease * target).toLocaleString("fr-FR") + suffix;
    if (t < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}

function staggerReveal(selector, delayUnit = 80) {
  document.querySelectorAll(selector).forEach((el, idx) => {
    el.style.animation = `fadeInUp 0.6s cubic-bezier(0.34, 1.56, 0.64, 1) forwards`;
    el.style.animationDelay = `${idx * delayUnit}ms`;
  });
}

/* =========================================================
   SVG SNIPPETS
   ========================================================= */
const SVG_FILE = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75"
  stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" style="width:16px;height:16px;display:block;">
  <path d="M13 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V9z"/>
  <polyline points="13 2 13 9 20 9"/>
</svg>`;

const SVG_FOLDER_EMPTY = `<svg class="empty-state-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"
  stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
  <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/>
</svg>`;

const SVG_WARNING = `<svg class="error-icon-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor"
  stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
  <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
  <line x1="12" y1="9" x2="12" y2="13"/>
  <line x1="12" y1="17" x2="12.01" y2="17"/>
</svg>`;

/* SVGs pour les recommandations */
const RECO_ICONS = {
  trend: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>`,
  compare: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="7" r="1.5"/><circle cx="17" cy="17" r="1.5"/><circle cx="7" cy="17" r="1.5"/><circle cx="17" cy="7" r="1.5"/><line x1="7" y1="7" x2="17" y2="17"/></svg>`,
  alert: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
};

const PRIORITY_META = {
  warning: { label: "À surveiller", cls: "reco-card--warning" },
  info:    { label: "Recommandé",   cls: "reco-card--info"    },
  normal:  { label: "Disponible",   cls: "reco-card--normal"  },
};

/* =========================================================
   BUILDERS
   ========================================================= */

/* ── Signal Bar ── */
function buildSignalBar(signals = []) {
  const el = qs("signal-bar");
  if (!el) return;

  if (!signals.length) {
    el.innerHTML = `
      <div class="signal-chip signal--ok">
        <span class="signal-dot">🟢</span>
        <span class="signal-label">Système opérationnel</span>
        <span class="signal-detail">Aucun signal à surveiller</span>
      </div>`;
    return;
  }

  el.innerHTML = signals.map(sig => {
    const cls = sig.type === "critical" ? "signal--critical"
              : sig.type === "warning"  ? "signal--warning"
              : "signal--ok";
    const dot = sig.type === "critical" ? "🔴"
              : sig.type === "warning"  ? "🟡"
              : "🟢";
    return `
      <div class="signal-chip ${cls}">
        <span class="signal-dot">${dot}</span>
        <span class="signal-label">${sig.label}</span>
        <span class="signal-detail">${sig.detail}</span>
      </div>`;
  }).join("");
}

/* ── Recommandations ── */
function buildRecommendations(recommendations = []) {
  const el = qs("reco-grid");
  if (!el || !recommendations.length) return;

  el.innerHTML = recommendations.map((reco, idx) => {
    const meta  = PRIORITY_META[reco.priority] || PRIORITY_META.normal;
    const icon  = RECO_ICONS[reco.icon]        || RECO_ICONS.alert;
    return `
      <a href="${reco.url}" class="reco-card ${meta.cls}" style="animation-delay:${idx * 80}ms">
        <div class="reco-card__icon">${icon}</div>
        <div class="reco-card__body">
          <div class="reco-card__badge">${meta.label}</div>
          <div class="reco-card__title">${reco.title}</div>
          <div class="reco-card__detail">${reco.detail}</div>
        </div>
        <div class="reco-card__arrow">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg>
        </div>
      </a>`;
  }).join("");
}

/* ── Années ── */
function buildAnnees(annees = []) {
  const el = qs("annees-badges");
  if (!el) return;
  if (!Array.isArray(annees) || annees.length === 0) {
    el.innerHTML = `<span class="badge badge--gray">Aucune année</span>`;
    return;
  }
  el.innerHTML = annees.map(a => `<span class="badge badge--blue">${a}</span>`).join("");
  setText("annees-count", String(annees.length));
}

/* ── Sections ── */
function buildSections(sections = []) {
  const el = qs("sections-chips");
  if (!el) return;
  if (!Array.isArray(sections) || sections.length === 0) {
    el.innerHTML = `<span class="badge badge--gray">Aucune section</span>`;
    return;
  }
  el.innerHTML = sections.map(s => `<div class="chip">${s}</div>`).join("");
  setText("sections-count", String(sections.length));
}

/* ── Fichiers ── */
function buildFichiers(fichiers = []) {
  const tbody = qs("fichiers-tbody");
  if (!tbody) return;
  if (!Array.isArray(fichiers) || fichiers.length === 0) {
    tbody.innerHTML = `
      <tr><td colspan="4">
        <div class="empty-state" style="min-height:120px;">
          ${SVG_FOLDER_EMPTY}
          <div>Aucun fichier récent disponible</div>
        </div>
      </td></tr>`;
    return;
  }
  tbody.innerHTML = fichiers.map((f, idx) => `
    <tr class="table-row" style="animation-delay:${idx * 50}ms;">
      <td class="td-file">
        <div class="td-file-cell">
          <div class="td-file-icon">${SVG_FILE}</div>
          <div class="td-file-text">
            <span class="file-name">${f.nom ?? "—"}</span>
            <span class="file-sub">Fichier mensuel PULSE</span>
          </div>
        </div>
      </td>
      <td class="td-size">${formatKo(f.taille_ko)}</td>
      <td class="td-date">${f.modifie ?? "—"}</td>
      <td class="td-status">
        <span class="file-status file-status--active">
          <span class="status-indicator"></span>Actif
        </span>
      </td>
    </tr>`).join("");

  document.querySelectorAll(".table-row").forEach(row => {
    row.style.animation = "fadeInUp 0.4s ease forwards";
  });
}

/* =========================================================
   UPDATE KPIs avec contexte interprétatif
   ========================================================= */
function updateKpis(data) {
  const annees = Array.isArray(data.annees) ? data.annees : [];

  animateCount(qs("kpi-fichiers"), data.nb_fichiers);
  animateCount(qs("kpi-sections"), data.nb_sections);
  animateCount(qs("kpi-annees"),   annees.length);
  setText("kpi-maj", data.derniere_maj);

  // Contextes interprétatifs
  if (annees.length >= 2) {
    setText("kpi-fichiers-ctx", `Données sur ${annees.length} ans (${Math.min(...annees)}–${Math.max(...annees)})`);
  } else {
    setText("kpi-fichiers-ctx", data.nb_fichiers >= 12 ? "Données complètes sur l'année" : "Données partielles");
  }

  const freshMap = { recent: "Données fraîches", normal: "Synchronisation normale", stale: "Synchronisation ancienne", unknown: "" };
  const freshBadgeMap = { recent: "badge--green", normal: "badge--green", stale: "badge--orange", unknown: "badge--gray" };
  const freshLabelMap = { recent: "Récent", normal: "Sync", stale: "Ancien", unknown: "Sync" };

  setText("kpi-maj-ctx", freshMap[data.freshness] || "");
  const majBadge = qs("kpi-maj-badge");
  if (majBadge) {
    majBadge.textContent = freshLabelMap[data.freshness] || "Sync";
    majBadge.className = `badge badge--sm ${freshBadgeMap[data.freshness] || "badge--gray"}`;
  }

  setText("kpi-sections-ctx", data.nb_sections > 0 ? `${data.nb_sections} filiale(s) sous surveillance` : "Aucune section active");

  if (annees.length >= 2) {
    const hasGaps = data.missing_years && data.missing_years.length > 0;
    setText("kpi-annees-ctx", hasGaps ? `⚠ ${data.missing_years.length} année(s) manquante(s)` : `Série continue ${Math.min(...annees)}–${Math.max(...annees)}`);
  } else {
    setText("kpi-annees-ctx", annees.length === 1 ? `Année ${annees[0]} uniquement` : "—");
  }
}

function updateHero(data) {
  const annees = Array.isArray(data.annees) ? data.annees : [];
  animateCount(qs("hero-total-files"),   data.nb_fichiers);
  animateCount(qs("hero-stat-annees"),   annees.length);
  animateCount(qs("hero-stat-sections"), data.nb_sections);

  const perimEl = qs("hero-perimeter");
  if (perimEl) {
    perimEl.textContent = annees.length >= 2
      ? `${annees.length} ans · ${data.nb_sections} sections`
      : `${data.nb_sections} section(s)`;
  }

  const totalLabel = qs("total-fichiers-label");
  if (totalLabel) {
    totalLabel.textContent = `${formatNumber(data.nb_fichiers)} ${plural(data.nb_fichiers, "fichier")} détecté${data.nb_fichiers > 1 ? "s" : ""}`;
  }
}


/* =========================================================
   LIFECYCLE
   ========================================================= */
function showDashboard() {
  const loadingEl   = qs("loading");
  const dashboardEl = qs("dashboard");
  if (loadingEl)   loadingEl.hidden   = true;
  if (dashboardEl) dashboardEl.hidden = false;

  setTimeout(() => {
    staggerReveal(".kpi-card",      100);
    staggerReveal(".hub-card",       80);
    staggerReveal(".reco-card",      80);
    staggerReveal(".card--premium", 120);
  }, 50);
}

function showError(message) {
  const loadingEl = qs("loading");
  if (!loadingEl) return;
  loadingEl.innerHTML = `
    <div class="error-state">
      ${SVG_WARNING}
      <span>${message}</span>
    </div>`;
}

/* =========================================================
   SPARKLINES — mini bar chart SVG inline
   ========================================================= */
function buildSparkline(data, color = "#4C7CF3") {
  if (!Array.isArray(data) || data.length === 0) return "";
  const W = 80, H = 24, gap = 2;
  const max = Math.max(...data.map(d => d.count), 1);
  const barW = (W - gap * (data.length - 1)) / data.length;
  const bars = data.map((d, i) => {
    const h = Math.max(3, Math.round((d.count / max) * H));
    const x = i * (barW + gap);
    const y = H - h;
    return `<rect x="${x.toFixed(1)}" y="${y}" width="${barW.toFixed(1)}" height="${h}" rx="1.5"
      fill="${color}" opacity="${d.count === max ? '1' : '0.5'}" />`;
  }).join("");
  return `<svg class="kpi-sparkline" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg"
    aria-hidden="true">${bars}</svg>`;
}

function injectSparkline(cardSelector, svgHtml) {
  const card = document.querySelector(cardSelector);
  if (!card || !svgHtml) return;
  const wrap = document.createElement("div");
  wrap.className = "kpi-sparkline-wrap";
  wrap.innerHTML = svgHtml;
  card.querySelector(".kpi-card__inner")?.appendChild(wrap);
}

function buildSparklines(fichiers_par_annee) {
  if (!Array.isArray(fichiers_par_annee) || fichiers_par_annee.length < 2) return;
  const svg = buildSparkline(fichiers_par_annee, "#4C7CF3");
  injectSparkline('.kpi-card[data-index="0"]', svg);
}

function normalizeData(data) {
  return {
    nb_fichiers:          Number(data?.nb_fichiers      ?? 0),
    nb_sections:          Number(data?.nb_sections      ?? 0),
    nb_entrees_cache:     Number(data?.nb_entrees_cache ?? 0),
    derniere_maj:         data?.derniere_maj             ?? "—",
    annees:               Array.isArray(data?.annees)              ? data.annees              : [],
    sections:             Array.isArray(data?.sections)            ? data.sections            : [],
    fichiers_recents:     Array.isArray(data?.fichiers_recents)    ? data.fichiers_recents    : [],
    days_ago:             data?.days_ago                 ?? null,
    freshness:            data?.freshness                ?? "unknown",
    year_span:            data?.year_span                ?? 0,
    missing_years:        Array.isArray(data?.missing_years)       ? data.missing_years       : [],
    signals:              Array.isArray(data?.signals)             ? data.signals             : [],
    recommendations:      Array.isArray(data?.recommendations)     ? data.recommendations     : [],
    fichiers_par_annee:   Array.isArray(data?.fichiers_par_annee)  ? data.fichiers_par_annee  : [],
  };
}

async function loadAccueil() {
  const res = await fetch("/api/accueil", { method: "GET", headers: { Accept: "application/json" } });
  if (!res.ok) throw new Error(`Erreur serveur : HTTP ${res.status}`);
  return normalizeData(await res.json());
}

function renderAccueil(data) {
  updateKpis(data);
  updateHero(data);
  buildSignalBar(data.signals);
  buildRecommendations(data.recommendations);
  buildAnnees(data.annees);
  buildSections(data.sections);
  buildFichiers(data.fichiers_recents);
  showDashboard();
  buildSparklines(data.fichiers_par_annee);
}

/* =========================================================
   HERO CHART — Graph tracing canvas animation
   ========================================================= */
function initHeroChart() {
  const canvas = document.getElementById("hero-chart-canvas");
  if (!canvas) return;

  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const W = 180, H = 115;
  canvas.width  = W * dpr;
  canvas.height = H * dpr;
  canvas.style.width  = W + "px";
  canvas.style.height = H + "px";

  const ctx = canvas.getContext("2d");
  ctx.scale(dpr, dpr);

  /* ── Generate realistic-looking chart paths ── */
  function makePath(n, baseY, trend, freqs) {
    const pts = [];
    for (let i = 0; i <= n; i++) {
      const t = i / n;
      let y = baseY - trend * t;
      freqs.forEach(([f, a, ph]) => { y += Math.sin(t * f * Math.PI * 2 + ph) * a; });
      pts.push({ x: t * W, y: Math.max(H * 0.06, Math.min(H * 0.94, y)) });
    }
    return pts;
  }

  /* Main (cyan) — upward trend */
  const P1 = makePath(72, H * 0.60, H * 0.22, [
    [1.3, H * 0.13, 0.7],
    [3.6, H * 0.07, 1.9],
    [8.1, H * 0.03, 3.4],
  ]);

  /* Secondary (violet) — slightly lagged */
  const P2 = makePath(72, H * 0.68, H * 0.14, [
    [1.0, H * 0.09, 2.2],
    [2.9, H * 0.06, 0.5],
    [6.5, H * 0.03, 2.8],
  ]);

  /* Interpolate a point along a path at normalized progress t */
  function atT(pts, t) {
    const raw = t * (pts.length - 1);
    const i   = Math.min(Math.floor(raw), pts.length - 2);
    const f   = raw - i;
    return {
      x: pts[i].x + (pts[i + 1].x - pts[i].x) * f,
      y: pts[i].y + (pts[i + 1].y - pts[i].y) * f,
    };
  }

  /* Catmull-Rom smoothed stroke up to progress t (0..1) */
  function drawLine(pts, t, r, g, b, lw) {
    const end = Math.min(Math.ceil(t * (pts.length - 1)) + 1, pts.length);
    if (end < 2) return;

    /* Horizontal fade-in gradient */
    const x1 = pts[end - 1].x;
    const grad = ctx.createLinearGradient(0, 0, x1, 0);
    grad.addColorStop(0,                       `rgba(${r},${g},${b},0.00)`);
    grad.addColorStop(Math.max(0, t - 0.40),   `rgba(${r},${g},${b},0.20)`);
    grad.addColorStop(1,                       `rgba(${r},${g},${b},0.85)`);

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);
    for (let i = 1; i < end; i++) {
      const pp = pts[Math.max(0, i - 2)];
      const p0 = pts[i - 1];
      const p1 = pts[i];
      const p2 = pts[Math.min(pts.length - 1, i + 1)];
      /* Catmull-Rom → cubic Bezier control points */
      const cp1x = p0.x + (p1.x - pp.x) / 6;
      const cp1y = p0.y + (p1.y - pp.y) / 6;
      const cp2x = p1.x - (p2.x - p0.x) / 6;
      const cp2y = p1.y - (p2.y - p0.y) / 6;
      ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p1.x, p1.y);
    }
    ctx.strokeStyle = grad;
    ctx.lineWidth   = lw;
    ctx.lineCap     = "round";
    ctx.lineJoin    = "round";
    ctx.stroke();
    ctx.restore();
  }

  /* Glowing cyan tip */
  function drawTip(x, y) {
    /* Multi-layer halo */
    [[30, 0.04], [18, 0.09], [10, 0.18], [5.5, 0.35]].forEach(([radius, alpha]) => {
      const grd = ctx.createRadialGradient(x, y, 0, x, y, radius);
      grd.addColorStop(0, `rgba(56,189,248,${alpha})`);
      grd.addColorStop(1, "rgba(56,189,248,0)");
      ctx.beginPath();
      ctx.arc(x, y, radius, 0, Math.PI * 2);
      ctx.fillStyle = grd;
      ctx.fill();
    });
    /* Core dot — solid cyan */
    ctx.beginPath();
    ctx.arc(x, y, 2.6, 0, Math.PI * 2);
    ctx.fillStyle = "#38bdf8";
    ctx.fill();
    /* White-hot centre */
    ctx.beginPath();
    ctx.arc(x, y, 1.1, 0, Math.PI * 2);
    ctx.fillStyle = "#ffffff";
    ctx.fill();
  }

  /* ── Animation state machine ── */
  const DRAW_MS = 3400, HOLD_MS = 1800, FADE_MS = 700;
  let phase = "draw", progress = 0, lastTs = null, phaseStart = null;

  function frame(ts) {
    if (!lastTs) { lastTs = ts; phaseStart = ts; }
    const dt  = ts - lastTs;
    lastTs = ts;

    ctx.clearRect(0, 0, W, H);

    if (phase === "draw") {
      progress = Math.min(1, progress + dt / DRAW_MS);
      ctx.globalAlpha = 1;

      drawLine(P2, Math.max(0, progress - 0.09), 139, 92, 246, 1.1);
      drawLine(P1, progress, 56, 189, 248, 2.0);

      const tip = atT(P1, progress);
      drawTip(tip.x, tip.y);

      if (progress >= 1) { phase = "hold"; phaseStart = ts; }

    } else if (phase === "hold") {
      ctx.globalAlpha = 1;
      drawLine(P2, 1, 139, 92, 246, 1.1);
      drawLine(P1, 1, 56, 189, 248, 2.0);
      const tip = atT(P1, 1);
      drawTip(tip.x, tip.y);

      if (ts - phaseStart >= HOLD_MS) { phase = "fade"; phaseStart = ts; }

    } else {
      /* fade out then reset */
      const a = Math.max(0, 1 - (ts - phaseStart) / FADE_MS);
      ctx.globalAlpha = a;
      drawLine(P2, 1, 139, 92, 246, 1.1);
      drawLine(P1, 1, 56, 189, 248, 2.0);
      ctx.globalAlpha = 1;

      if (a <= 0) { phase = "draw"; progress = 0; lastTs = null; phaseStart = null; }
    }

    requestAnimationFrame(frame);
  }

  requestAnimationFrame(frame);
}

/* =========================================================
   HERO HISTOGRAM — Animated bar chart canvas
   ========================================================= */
function initHeroHist() {
  const canvas = document.getElementById("hero-hist-canvas");
  if (!canvas) return;

  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const W = 148, H = 82;
  canvas.width  = W * dpr;
  canvas.height = H * dpr;
  canvas.style.width  = W + "px";
  canvas.style.height = H + "px";

  const ctx = canvas.getContext("2d");
  ctx.scale(dpr, dpr);

  /* Bar heights (0..1) — realistic distribution, not uniform */
  const BARS = [0.42, 0.68, 0.35, 0.82, 0.55, 0.91, 0.47, 0.75, 0.60, 0.88];
  const N     = BARS.length;
  const GAP   = 3.5;
  const BAR_W = (W - GAP * (N + 1)) / N;
  const PAD_T = 10; /* top padding so glow doesn't clip */

  const STAGGER_MS = 60;   /* delay between each bar */
  const GROW_MS    = 480;  /* time for one bar to reach its height */
  const HOLD_MS    = 2000;
  const FADE_MS    = 650;

  const progress = new Float32Array(N);

  function easeOutCubic(t) { return 1 - Math.pow(1 - t, 3); }

  function drawBars(alpha) {
    ctx.globalAlpha = alpha;

    for (let i = 0; i < N; i++) {
      const p    = easeOutCubic(progress[i]);
      const x    = GAP + i * (BAR_W + GAP);
      const maxH = BARS[i] * (H - PAD_T);
      const bh   = p * maxH;
      const y    = H - bh;

      if (bh < 0.5) continue;

      /* ── Body gradient: dark base → bright cyan cap ── */
      const g = ctx.createLinearGradient(x, H, x, y);
      g.addColorStop(0.0, "rgba(14,165,233,0.18)");
      g.addColorStop(0.5, "rgba(56,189,248,0.55)");
      g.addColorStop(1.0, "rgba(56,189,248,0.90)");

      /* Rounded top corners */
      const r = Math.min(3.5, bh / 2);
      ctx.beginPath();
      ctx.moveTo(x,           H);
      ctx.lineTo(x,           y + r);
      ctx.quadraticCurveTo(x,          y, x + r,          y);
      ctx.lineTo(x + BAR_W - r,        y);
      ctx.quadraticCurveTo(x + BAR_W,  y, x + BAR_W, y + r);
      ctx.lineTo(x + BAR_W,  H);
      ctx.closePath();
      ctx.fillStyle = g;
      ctx.fill();

      /* ── Top-edge highlight (thin bright line) ── */
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + BAR_W - r, y);
      ctx.strokeStyle = `rgba(186,230,253,${p * 0.55})`;
      ctx.lineWidth = 1;
      ctx.stroke();
      ctx.restore();

      /* ── Glow halo at cap — appears as bar nears full height ── */
      if (p > 0.7) {
        const glowA = ((p - 0.7) / 0.3) * 0.22;
        const cx    = x + BAR_W / 2;
        const glow  = ctx.createRadialGradient(cx, y, 0, cx, y, BAR_W * 1.5);
        glow.addColorStop(0, `rgba(56,189,248,${glowA})`);
        glow.addColorStop(1, "rgba(56,189,248,0)");
        ctx.beginPath();
        ctx.arc(cx, y, BAR_W * 1.5, 0, Math.PI * 2);
        ctx.fillStyle = glow;
        ctx.fill();
      }
    }

    ctx.globalAlpha = 1;
  }

  /* ── State machine: draw → hold → fade → reset ── */
  let phase = "draw", lastTs = null, phaseStart = null;

  function frame(ts) {
    if (!lastTs) { lastTs = ts; phaseStart = ts; }
    lastTs = ts;

    ctx.clearRect(0, 0, W, H);

    if (phase === "draw") {
      const elapsed = ts - phaseStart;
      let done = true;
      for (let i = 0; i < N; i++) {
        const t = (elapsed - i * STAGGER_MS) / GROW_MS;
        progress[i] = Math.min(1, Math.max(0, t));
        if (progress[i] < 1) done = false;
      }
      drawBars(1);
      if (done) { phase = "hold"; phaseStart = ts; }

    } else if (phase === "hold") {
      drawBars(1);
      if (ts - phaseStart >= HOLD_MS) { phase = "fade"; phaseStart = ts; }

    } else {
      const a = Math.max(0, 1 - (ts - phaseStart) / FADE_MS);
      drawBars(a);
      if (a <= 0) {
        phase = "draw";
        progress.fill(0);
        lastTs = null;
        phaseStart = null;
      }
    }

    requestAnimationFrame(frame);
  }

  requestAnimationFrame(frame);
}

/* =========================================================
   INIT
   ========================================================= */
initHeroChart();
initHeroHist();

(async () => {
  try {
    const data = await loadAccueil();
    renderAccueil(data);
  } catch (err) {
    showError(err?.message || "Impossible de charger les données du tableau de bord.");
  }
})();
