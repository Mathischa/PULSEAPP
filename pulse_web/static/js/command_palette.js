/* command_palette.js — PULSE Command Palette (Ctrl+K) */
"use strict";

(function () {

/* ── Commandes ─────────────────────────────────────────────── */
const COMMANDS = [
  { label: "Tableau de bord",       url: "/",                     icon: "home",    cat: "Principal"              },
  { label: "Import des profils",    url: "/import",               icon: "upload",  cat: "Principal"              },
  { label: "Écarts importants",     url: "/ecarts",               icon: "alert",   cat: "Surveillance"           },
  { label: "Tendance flux",         url: "/tendance",             icon: "trend",   cat: "Surveillance"           },
  { label: "Visualisation",         url: "/visualisation",        icon: "eye",     cat: "Analyse"                },
  { label: "Superposition années",  url: "/visualisation_flux",   icon: "layers",  cat: "Analyse"                },
  { label: "Benchmarking",          url: "/benchmarking",         icon: "compare", cat: "Analyse"                },
  { label: "Répartition filiales",  url: "/repartition",          icon: "bar",     cat: "Analyse"                },
  { label: "Répartition par flux",  url: "/repartition_flux",     icon: "list",    cat: "Analyse"                },
  { label: "Par profil",            url: "/prevision_repartition",icon: "chart",   cat: "Analyse"                },
  { label: "Heatmap anomalies",     url: "/heatmap",              icon: "grid",    cat: "Intelligence artificielle" },
  { label: "Heatmap des écarts",    url: "/heatmap_ecarts",       icon: "grid",    cat: "Intelligence artificielle" },
  { label: "Clustering ML",         url: "/ml_ecarts",            icon: "neural",  cat: "Intelligence artificielle" },
  { label: "Surface 3D",            url: "/clustering_3d",        icon: "cube",    cat: "Intelligence artificielle" },
];

const KEYWORDS = {
  "Tableau de bord":       ["dashboard","accueil","home","bord"],
  "Import des profils":    ["import","profil","fichier","excel","upload"],
  "Écarts importants":     ["ecart","alerte","deviation","important","ecarts"],
  "Tendance flux":         ["tendance","evolution","flux","time","serie"],
  "Visualisation":         ["visu","graph","courbe","chart","reel","prevision"],
  "Superposition années":  ["superposition","annee","multi","annees","flux"],
  "Benchmarking":          ["bench","compare","scatter","radar","performance"],
  "Répartition filiales":  ["repartition","filiale","bar","donut"],
  "Répartition par flux":  ["repartition","flux","liste","ecart"],
  "Par profil":            ["profil","prevision","repartition"],
  "Heatmap anomalies":     ["heatmap","anomalie","chaleur","heat","map"],
  "Heatmap des écarts":    ["heatmap","ecart","heat","map"],
  "Clustering ML":         ["ml","clustering","machine","learning","ia","ai","neural"],
  "Surface 3D":            ["3d","surface","plotly","3","dimension","spatial"],
};

const ICONS = {
  home:    `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12L12 3l9 9"/><path d="M9 21V12h6v9"/></svg>`,
  upload:  `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`,
  alert:   `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
  trend:   `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>`,
  eye:     `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="1.5"/><path d="M2 12C4 7 8 4 12 4s8 3 10 8c-2 5-6 8-10 8S4 17 2 12z"/></svg>`,
  layers:  `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>`,
  compare: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="7" r="1.5"/><circle cx="17" cy="17" r="1.5"/><circle cx="7" cy="17" r="1.5"/><circle cx="17" cy="7" r="1.5"/><line x1="7" y1="7" x2="17" y2="17" stroke-width="1.2" opacity=".6"/></svg>`,
  bar:     `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="18" y="3" width="3" height="18" rx="1"/><rect x="10.5" y="8" width="3" height="13" rx="1"/><rect x="3" y="13" width="3" height="8" rx="1"/></svg>`,
  list:    `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="8" x2="21" y2="8"/><line x1="3" y1="12" x2="17" y2="12"/><line x1="3" y1="16" x2="13" y2="16"/></svg>`,
  chart:   `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="12" x2="21" y2="12"/><rect x="3" y="2" width="4" height="8" rx="1"/><rect x="9" y="5" width="4" height="14" rx="1"/><rect x="15" y="3" width="4" height="16" rx="1"/></svg>`,
  grid:    `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/></svg>`,
  neural:  `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="4" r="2"/><circle cx="4" cy="20" r="2"/><circle cx="20" cy="20" r="2"/><line x1="12" y1="6" x2="4" y2="18"/><line x1="12" y1="6" x2="20" y2="18"/><line x1="4" y1="18" x2="20" y2="18"/></svg>`,
  cube:    `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M2 20h20M2 20V8l10-6 10 6v12"/><path d="M8 20v-6h8v6"/></svg>`,
};

/* ── State ─────────────────────────────────────────────────── */
let isOpen = false, selected = -1, filtered = [];
let backdrop, palette, inputEl, resultsEl;

/* ── DOM Init ───────────────────────────────────────────────── */
function init() {
  backdrop  = document.getElementById('cmd-backdrop');
  palette   = document.getElementById('cmd-palette');
  inputEl   = document.getElementById('cmd-input');
  resultsEl = document.getElementById('cmd-results');
  if (!backdrop || !palette || !inputEl || !resultsEl) return;

  document.addEventListener('keydown', onGlobalKey);
  backdrop.addEventListener('click', close);
  inputEl.addEventListener('input', renderResults);
  inputEl.addEventListener('keydown', onNavKey);

  // Trigger button
  const triggerBtn = document.getElementById('cmd-trigger-btn');
  if (triggerBtn) triggerBtn.addEventListener('click', open);
}

/* ── Open / Close ───────────────────────────────────────────── */
function open() {
  if (isOpen) return;
  isOpen = true;
  backdrop.hidden = false;
  palette.hidden  = false;
  requestAnimationFrame(() => {
    backdrop.classList.add('cmd-backdrop--in');
    palette.classList.add('cmd-palette--in');
  });
  inputEl.value = '';
  renderResults();
  setTimeout(() => inputEl.focus(), 50);
}

function close() {
  if (!isOpen) return;
  isOpen = false;
  backdrop.classList.remove('cmd-backdrop--in');
  palette.classList.remove('cmd-palette--in');
  setTimeout(() => { backdrop.hidden = true; palette.hidden = true; }, 220);
}

/* ── Fuzzy search ───────────────────────────────────────────── */
function scoreCmd(cmd, q) {
  if (!q) return 1;
  const parts  = [cmd.label.toLowerCase(), ...(KEYWORDS[cmd.label] || [])];
  const needle = q.toLowerCase();
  for (const part of parts) {
    if (part.startsWith(needle)) return 3;
    if (part.includes(needle))  return 2;
  }
  // char-by-char fuzzy on label
  let i = 0;
  const lbl = cmd.label.toLowerCase();
  for (const ch of needle) {
    i = lbl.indexOf(ch, i);
    if (i === -1) return 0;
    i++;
  }
  return 0.5;
}

function highlight(text, q) {
  if (!q) return text;
  const safe = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return text.replace(new RegExp(`(${safe})`, 'gi'), '<mark>$1</mark>');
}

/* ── Render ─────────────────────────────────────────────────── */
function renderResults() {
  const q = inputEl.value.trim();

  filtered = COMMANDS
    .map(cmd => ({ cmd, s: scoreCmd(cmd, q) }))
    .filter(x => x.s > 0)
    .sort((a, b) => b.s - a.s)
    .map(x => x.cmd);

  selected = filtered.length ? 0 : -1;

  if (!filtered.length) {
    resultsEl.innerHTML = `<div class="cmd-empty">Aucun résultat pour « ${q} »</div>`;
    return;
  }

  // Group by category when no search
  const grouped = !q;
  const cats = {};
  filtered.forEach((cmd, idx) => {
    if (!cats[cmd.cat]) cats[cmd.cat] = [];
    cats[cmd.cat].push({ cmd, idx });
  });

  let html = '';
  for (const [cat, items] of Object.entries(cats)) {
    if (grouped) html += `<div class="cmd-group-label">${cat}</div>`;
    items.forEach(({ cmd, idx }) => {
      html += `
        <a href="${cmd.url}" class="cmd-item${idx === 0 ? ' cmd-item--active' : ''}" data-idx="${idx}">
          <span class="cmd-item__icon">${ICONS[cmd.icon] || ''}</span>
          <span class="cmd-item__label">${highlight(cmd.label, q)}</span>
          ${q ? `<span class="cmd-item__cat">${cmd.cat}</span>` : ''}
          <span class="cmd-item__chevron"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg></span>
        </a>`;
    });
  }
  resultsEl.innerHTML = html;

  resultsEl.querySelectorAll('.cmd-item').forEach(el => {
    el.addEventListener('mouseenter', () => setSelected(+el.dataset.idx));
    el.addEventListener('click', e => { e.preventDefault(); go(el.getAttribute('href')); });
  });
}

function setSelected(idx) {
  if (idx < 0 || idx >= filtered.length) return;
  selected = idx;
  resultsEl.querySelectorAll('.cmd-item').forEach(el => {
    const active = +el.dataset.idx === idx;
    el.classList.toggle('cmd-item--active', active);
    if (active) el.scrollIntoView({ block: 'nearest' });
  });
}

function go(url) {
  close();
  setTimeout(() => { window.location.href = url; }, 80);
}

/* ── Keyboard ───────────────────────────────────────────────── */
function onGlobalKey(e) {
  if ((e.ctrlKey || e.metaKey) && e.key === 'k') { e.preventDefault(); isOpen ? close() : open(); }
  if (e.key === 'Escape' && isOpen) close();
}

function onNavKey(e) {
  if (e.key === 'ArrowDown') { e.preventDefault(); setSelected(Math.min(selected + 1, filtered.length - 1)); }
  else if (e.key === 'ArrowUp')  { e.preventDefault(); setSelected(Math.max(selected - 1, 0)); }
  else if (e.key === 'Enter')    { e.preventDefault(); if (filtered[selected]) go(filtered[selected].url); }
}

/* ── Expose global ──────────────────────────────────────────── */
window.pulseCmd = { open, close };

document.addEventListener('DOMContentLoaded', init);

})();
