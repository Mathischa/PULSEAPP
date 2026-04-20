/* ecarts.js — Analyse des écarts PULSE */
"use strict";

let allData  = [];
let sortCol  = "ecart_pct";
let sortDesc = true;

/* ── Formatage ────────────────────────────────────────────── */
const fmt    = (v) => typeof v === "number"
  ? v.toLocaleString("fr-FR", { minimumFractionDigits: 0, maximumFractionDigits: 2 })
  : v;

const fmtPct = (v) => {
  const s = fmt(Math.abs(v));
  return `${v > 0 ? "+" : "−"}${s} %`;
};

/* ── Peupler un <select> ──────────────────────────────────── */
function populateSelect(id, values) {
  const sel = document.getElementById(id);
  for (const v of values) {
    const opt     = document.createElement("option");
    opt.value     = opt.textContent = v;
    sel.appendChild(opt);
  }
}

/* ── Initialiser les filtres ──────────────────────────────── */
function initFilters(data) {
  const annees   = [...new Set(data.map((r) => r.annee))].sort((a, b) => a - b);
  const filiales = [...new Set(data.map((r) => r.filiale))].sort();
  const flux     = [...new Set(data.map((r) => r.flux))].sort();

  populateSelect("f-annee",   annees);
  populateSelect("f-filiale", filiales);
  populateSelect("f-flux",    flux);

  /* Présélectionner la dernière année */
  if (annees.length) {
    document.getElementById("f-annee").value = annees[annees.length - 1];
  }
}

/* ── Filtrer les données ──────────────────────────────────── */
function getFiltered() {
  const annee   = document.getElementById("f-annee").value;
  const filiale = document.getElementById("f-filiale").value;
  const flux    = document.getElementById("f-flux").value;
  const fav     = document.getElementById("f-fav").value;

  return allData.filter((r) => {
    if (annee   && String(r.annee)        !== annee)   return false;
    if (filiale && r.filiale              !== filiale)  return false;
    if (flux    && r.flux                 !== flux)     return false;
    if (fav !== "" && String(r.favorable) !== fav)     return false;
    return true;
  });
}

/* ── Trier ────────────────────────────────────────────────── */
function sortData(data) {
  return [...data].sort((a, b) => {
    const va = sortCol === "ecart_pct" ? Math.abs(a[sortCol]) : a[sortCol];
    const vb = sortCol === "ecart_pct" ? Math.abs(b[sortCol]) : b[sortCol];
    if (va < vb) return sortDesc ? 1 : -1;
    if (va > vb) return sortDesc ? -1 : 1;
    return 0;
  });
}

/* ── Rendre le tableau ────────────────────────────────────── */
function render(data) {
  const tbody = document.getElementById("ecarts-tbody");
  tbody.innerHTML = "";

  document.getElementById("table-count").textContent =
    `${data.length.toLocaleString("fr-FR")} écart${data.length !== 1 ? "s" : ""}`;

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;color:var(--text-3);padding:32px;">
      Aucun écart pour les filtres sélectionnés</td></tr>`;
    return;
  }

  const frag = document.createDocumentFragment();
  for (const r of data) {
    const tr   = document.createElement("tr");
    tr.className = r.favorable ? "row-fav" : "row-def";
    const cls  = r.favorable ? "val-pos" : "val-neg";

    tr.innerHTML = `
      <td>${r.date}</td>
      <td style="color:var(--text-2);font-size:12px;">${r.profil}</td>
      <td><span class="badge badge--gray">${r.filiale}</span></td>
      <td>${r.flux}</td>
      <td class="num">${fmt(r.reel)}</td>
      <td class="num">${fmt(r.prevision)}</td>
      <td class="num ${cls}">${fmt(r.ecart_k)}</td>
      <td class="num ${cls}" style="font-weight:700;">${fmtPct(r.ecart_pct)}</td>
    `;
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

/* ── URL state ────────────────────────────────────────────── */
function pushUrlState() {
  const p = new URLSearchParams();
  const annee   = document.getElementById("f-annee")?.value;
  const filiale = document.getElementById("f-filiale")?.value;
  const flux    = document.getElementById("f-flux")?.value;
  const fav     = document.getElementById("f-fav")?.value;
  if (annee)   p.set("annee",   annee);
  if (filiale) p.set("filiale", filiale);
  if (flux)    p.set("flux",    flux);
  if (fav)     p.set("fav",     fav);
  const qs = p.toString();
  history.replaceState(null, "", qs ? "?" + qs : location.pathname);
}

function restoreFromUrl() {
  const p = new URLSearchParams(location.search);
  if (p.get("annee"))   document.getElementById("f-annee").value   = p.get("annee");
  if (p.get("filiale")) document.getElementById("f-filiale").value = p.get("filiale");
  if (p.get("flux"))    document.getElementById("f-flux").value    = p.get("flux");
  if (p.get("fav"))     document.getElementById("f-fav").value     = p.get("fav");
}

/* ── Mise à jour complète ─────────────────────────────────── */
function update() { render(sortData(getFiltered())); pushUrlState(); }

/* ── Export PDF ───────────────────────────────────────────── */
function exportPDF() {
  window.pulsePDF("Écarts importants — PULSE");
}

/* ── Export Excel ─────────────────────────────────────────── */
function exportExcel() {
  const filtered = sortData(getFiltered());
  const headers  = ["Date","Profil","Filiale","Flux","Réel (k€)","Prévision (k€)","Écart (k€)","Écart (%)","Favorable"];
  const rows = filtered.map((r) => [
    r.date, r.profil, r.filiale, r.flux,
    r.reel, r.prevision, r.ecart_k, r.ecart_pct,
    r.favorable ? "Oui" : "Non",
  ]);
  window.pulseExcelData(headers, rows, `ecarts_pulse_${new Date().toISOString().slice(0,10)}`);
  window.toast(`${filtered.length} lignes exportées en Excel`, "success");
}

/* ── Export CSV ───────────────────────────────────────────── */
function exportCSV() {
  const filtered = sortData(getFiltered());
  const headers  = ["Date","Profil","Filiale","Flux","Réel (k€)","Prévision (k€)","Écart (k€)","Écart (%)","Favorable"];
  const rows     = filtered.map((r) => [
    r.date, r.profil, r.filiale, r.flux,
    r.reel, r.prevision, r.ecart_k, r.ecart_pct,
    r.favorable ? "Oui" : "Non",
  ]);

  const csv  = [headers, ...rows]
    .map((row) => row.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
    .join("\r\n");

  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement("a"), {
    href: url,
    download: `ecarts_pulse_${new Date().toISOString().slice(0, 10)}.csv`,
  });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  window.toast(`${filtered.length} lignes exportées`, "success");
}

/* ── Sort headers ─────────────────────────────────────────── */
document.querySelectorAll(".data-table thead th[data-col]").forEach((th) => {
  th.addEventListener("click", () => {
    const col = th.dataset.col;
    sortDesc  = sortCol === col ? !sortDesc : true;
    sortCol   = col;

    document.querySelectorAll(".data-table thead th").forEach((h) =>
      h.classList.remove("sort-asc", "sort-desc")
    );
    th.classList.add(sortDesc ? "sort-desc" : "sort-asc");
    update();
  });
});

/* ── Filtres ──────────────────────────────────────────────── */
["f-annee", "f-filiale", "f-flux", "f-fav"].forEach((id) =>
  document.getElementById(id).addEventListener("change", update)
);

/* ── Reset filtres ────────────────────────────────────────── */
document.getElementById("btn-reset-filters")?.addEventListener("click", () => {
  ["f-annee", "f-filiale", "f-flux", "f-fav"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.selectedIndex = 0;
  });
  update();
  window.toast?.("Filtres réinitialisés", "info");
});

/* ── Export buttons ───────────────────────────────────────── */
document.getElementById("btn-export-pdf").addEventListener("click", exportPDF);
document.getElementById("btn-export-excel").addEventListener("click", exportExcel);
document.getElementById("btn-export").addEventListener("click", exportCSV);

/* ── Chargement initial ───────────────────────────────────── */
(async () => {
  const loadingEl   = document.getElementById("loading");
  const containerEl = document.getElementById("table-container");
  const btnExport   = document.getElementById("btn-export");

  try {
    const res = await fetch("/api/ecarts");
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    allData = await res.json();

    loadingEl.hidden       = true;
    containerEl.hidden     = false;
    btnExport.disabled     = false;
    const btnExcel = document.getElementById("btn-export-excel");
    if (btnExcel) btnExcel.disabled = false;

    initFilters(allData);
    restoreFromUrl();

    /* Indicateur tri initial */
    document.querySelector('th[data-col="ecart_pct"]')?.classList.add("sort-desc");

    update();
    window.toast(`${allData.length.toLocaleString("fr-FR")} écarts chargés`, "success");

  } catch (err) {
    loadingEl.innerHTML = `<div class="error-state">⚠ ${err.message}</div>`;
  }
})();
