/* accueil.js — Tableau de bord PULSE PREMIUM */
"use strict";

/* =========================================================
   HELPERS
   ========================================================= */
function qs(id) {
  return document.getElementById(id);
}

function setText(id, value, fallback = "—") {
  const el = qs(id);
  if (!el) return;
  el.textContent = value ?? fallback;
}

function setHTML(id, html) {
  const el = qs(id);
  if (!el) return;
  el.innerHTML = html;
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
    const t = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - t, 3); // ease-out cubic
    el.textContent = Math.round(ease * target).toLocaleString("fr-FR") + suffix;
    if (t < 1) requestAnimationFrame(step);
  };

  requestAnimationFrame(step);
}

/* =========================================================
   ANIMATION — Stagger reveal au load
   ========================================================= */
function staggerReveal(selector, delayUnit = 80) {
  const elements = document.querySelectorAll(selector);
  elements.forEach((el, idx) => {
    el.style.animation = `fadeInUp 0.6s cubic-bezier(0.34, 1.56, 0.64, 1) forwards`;
    el.style.animationDelay = `${idx * delayUnit}ms`;
  });
}

/* =========================================================
   SVG SNIPPETS (inline, no external dependency)
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

/* =========================================================
   BUILDERS
   ========================================================= */
function buildAnnees(annees = []) {
  const el = qs("annees-badges");
  if (!el) return;

  if (!Array.isArray(annees) || annees.length === 0) {
    el.innerHTML = `<span class="badge badge--gray">Aucune année</span>`;
    return;
  }

  el.innerHTML = annees
    .map((annee) => `<span class="badge badge--blue">${annee}</span>`)
    .join("");

  setText("annees-count", String(annees.length));
  animateCount(qs("stat-annees"), annees.length);
}

function buildSections(sections = []) {
  const el = qs("sections-chips");
  if (!el) return;

  if (!Array.isArray(sections) || sections.length === 0) {
    el.innerHTML = `<span class="badge badge--gray">Aucune section</span>`;
    return;
  }

  el.innerHTML = sections
    .map((section) => `<div class="chip">${section}</div>`)
    .join("");

  setText("sections-count", String(sections.length));
  animateCount(qs("stat-filiales"), sections.length);
}

function buildFichiers(fichiers = []) {
  const tbody = qs("fichiers-tbody");
  if (!tbody) return;

  if (!Array.isArray(fichiers) || fichiers.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="4">
          <div class="empty-state" style="min-height:120px;">
            ${SVG_FOLDER_EMPTY}
            <div>Aucun fichier récent disponible</div>
          </div>
        </td>
      </tr>
    `;
    animateCount(qs("stat-recents"), 0);
    return;
  }

  tbody.innerHTML = fichiers
    .map((f, idx) => {
      const nom     = f.nom     ?? "—";
      const taille  = formatKo(f.taille_ko);
      const modifie = f.modifie ?? "—";

      return `
        <tr class="table-row" style="animation-delay:${idx * 50}ms;">
          <td class="td-file">
            <div class="td-file-cell">
              <div class="td-file-icon">${SVG_FILE}</div>
              <div class="td-file-text">
                <span class="file-name">${nom}</span>
                <span class="file-sub">Fichier mensuel PULSE</span>
              </div>
            </div>
          </td>
          <td class="td-size">${taille}</td>
          <td class="td-date">${modifie}</td>
          <td class="td-status">
            <span class="file-status file-status--active">
              <span class="status-indicator"></span>
              Actif
            </span>
          </td>
        </tr>
      `;
    })
    .join("");

  document.querySelectorAll(".table-row").forEach((row) => {
    row.style.animation = "fadeInUp 0.4s ease forwards";
  });

  animateCount(qs("stat-recents"), fichiers.length);
}

function buildResume(data) {
  const nbFichiers  = data.nb_fichiers       ?? 0;
  const nbSections  = data.nb_sections       ?? 0;
  const nbCache     = data.nb_entrees_cache  ?? 0;
  const nbAnnees    = Array.isArray(data.annees) ? data.annees.length : 0;
  const derniereMaj = data.derniere_maj      ?? "—";

  setHTML(
    "resume-operationnel",
    `<strong>${formatNumber(nbFichiers)}</strong> ${plural(nbFichiers, "fichier")}
    sont actuellement chargés en mémoire, répartis sur
    <strong>${formatNumber(nbAnnees)}</strong> ${plural(nbAnnees, "année")}
    et <strong>${formatNumber(nbSections)}</strong> ${plural(nbSections, "section")}.
    La dernière mise à jour détectée est
    <strong>${derniereMaj}</strong>, avec
    <strong>${formatNumber(nbCache)}</strong> ${plural(nbCache, "entrée", "entrées")}
    conservées en cache pour accélérer l'analyse.`
  );
}

/* =========================================================
   UPDATES
   ========================================================= */
function updateInsights(data) {
  setText("insight-last-update", data.derniere_maj);
  setText("insight-coverage",    `${formatNumber(data.nb_sections)} sections`);
  setText("insight-cache",       `${formatNumber(data.nb_entrees_cache)} séries`);

  setText("sys-cache",    `${formatNumber(data.nb_entrees_cache)} séries`);
  setText("sys-sync",     data.derniere_maj);
  setText("sys-sections", `${formatNumber(data.nb_sections)} sections`);

  setText("stat-health", "OK");
}

function updateHero(data) {
  animateCount(qs("hero-total-files"),    data.nb_fichiers);
  animateCount(qs("hero-stat-annees"),    Array.isArray(data.annees) ? data.annees.length : 0);
  animateCount(qs("hero-stat-sections"),  data.nb_sections);

  // Fix: set text directly, el is already the badge
  const totalLabel = qs("total-fichiers-label");
  if (totalLabel) {
    totalLabel.textContent = `${formatNumber(data.nb_fichiers)} ${plural(data.nb_fichiers, "fichier")} détecté${data.nb_fichiers > 1 ? "s" : ""}`;
  }
}

function updateKpis(data) {
  animateCount(qs("kpi-fichiers"), data.nb_fichiers);
  animateCount(qs("kpi-sections"), data.nb_sections);
  animateCount(qs("kpi-cache"),    data.nb_entrees_cache);
  setText("kpi-maj", data.derniere_maj);
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
    staggerReveal(".kpi-card",    100);
    staggerReveal(".stat-item",   100);
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
    </div>
  `;
}

function normalizeData(data) {
  return {
    nb_fichiers:      Number(data?.nb_fichiers      ?? 0),
    nb_sections:      Number(data?.nb_sections      ?? 0),
    nb_entrees_cache: Number(data?.nb_entrees_cache ?? 0),
    derniere_maj:     data?.derniere_maj             ?? "—",
    annees:           Array.isArray(data?.annees)           ? data.annees           : [],
    sections:         Array.isArray(data?.sections)         ? data.sections         : [],
    fichiers_recents: Array.isArray(data?.fichiers_recents) ? data.fichiers_recents : []
  };
}

async function loadAccueil() {
  const res = await fetch("/api/accueil", {
    method: "GET",
    headers: { "Accept": "application/json" }
  });

  if (!res.ok) throw new Error(`Erreur serveur : HTTP ${res.status}`);

  const json = await res.json();
  return normalizeData(json);
}

function renderAccueil(data) {
  updateKpis(data);
  updateHero(data);
  updateInsights(data);
  buildAnnees(data.annees);
  buildSections(data.sections);
  buildFichiers(data.fichiers_recents);
  buildResume(data);
  showDashboard();
}

/* =========================================================
   INIT
   ========================================================= */
(async () => {
  try {
    const data = await loadAccueil();
    renderAccueil(data);
  } catch (err) {
    showError(err?.message || "Impossible de charger les données du tableau de bord.");
  }
})();
