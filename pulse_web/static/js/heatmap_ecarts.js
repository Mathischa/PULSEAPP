/* heatmap_ecarts.js — Heatmap des écarts significatifs (≥40%) profil × mois */
"use strict";

let _data      = null;
let _selCell   = null;
let _catalogue = {};   // { section: [flux_list] }

/* ── UTILS ────────────────────────────────────────────────── */
function escHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
const fmt2 = v => typeof v === "number" && isFinite(v)
  ? v.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "—";

/* ── COULEUR HEATMAP ──────────────────────────────────────── */
function cellColor(value, maxVal) {
  if (value === 0 || maxVal === 0) return null;
  const t = Math.min(value / maxVal, 1);
  const stops = [
    [76, 124, 243],   // bleu  (t=0)
    [251, 146,  60],  // orange (t=0.5)
    [239,  68,  68],  // rouge  (t=1)
  ];
  let r, g, b;
  if (t <= 0.5) {
    const f = t / 0.5;
    r = Math.round(stops[0][0] + f * (stops[1][0] - stops[0][0]));
    g = Math.round(stops[0][1] + f * (stops[1][1] - stops[0][1]));
    b = Math.round(stops[0][2] + f * (stops[1][2] - stops[0][2]));
  } else {
    const f = (t - 0.5) / 0.5;
    r = Math.round(stops[1][0] + f * (stops[2][0] - stops[1][0]));
    g = Math.round(stops[1][1] + f * (stops[2][1] - stops[1][1]));
    b = Math.round(stops[1][2] + f * (stops[2][2] - stops[1][2]));
  }
  const alpha = 0.18 + t * 0.72;
  return `rgba(${r},${g},${b},${alpha})`;
}

/* ── ÉTATS ────────────────────────────────────────────────── */
function showState(id) {
  ["hm-idle", "hm-loading", "hm-error", "hm-result"].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.hidden = (s !== id);
  });
}

/* ── TOOLTIP ──────────────────────────────────────────────── */
const tooltip = document.getElementById("hm-tooltip");
function showTooltip(e, html) {
  tooltip.innerHTML = html;
  tooltip.style.display = "block";
  positionTooltip(e);
}
function positionTooltip(e) {
  const x = e.clientX + 14, y = e.clientY + 14;
  tooltip.style.left = Math.min(x, window.innerWidth  - tooltip.offsetWidth  - 10) + "px";
  tooltip.style.top  = Math.min(y, window.innerHeight - tooltip.offsetHeight - 10) + "px";
}
function hideTooltip() { tooltip.style.display = "none"; }

/* ── LÉGENDE ──────────────────────────────────────────────── */
function buildLegend(maxVal) {
  const scale = document.getElementById("hm-legend-scale");
  const steps = 20;
  scale.innerHTML = "";
  for (let i = 0; i <= steps; i++) {
    const v   = (i / steps) * maxVal;
    const div = document.createElement("div");
    div.style.flex       = "1";
    div.style.background = cellColor(v, maxVal) || "rgba(255,255,255,.03)";
    scale.appendChild(div);
  }
  document.getElementById("hm-legend-max").textContent = maxVal;
}

/* ── RENDER HEATMAP ───────────────────────────────────────── */
function renderHeatmap(data) {
  _data    = data;
  _selCell = null;
  const btnExcel = document.getElementById("btn-export-excel");
  if (btnExcel) btnExcel.disabled = false;

  const { profils, mois, matrix, max_val, n_ecarts } = data;

  // En-têtes (mois)
  const thead = document.getElementById("hm-thead");
  thead.innerHTML = "";
  const headerRow = document.createElement("tr");
  const thCorner  = document.createElement("th");
  thCorner.className   = "hm-profil-header";
  thCorner.textContent = "Profil \\ Mois";
  headerRow.appendChild(thCorner);
  mois.forEach(m => {
    const th = document.createElement("th");
    // Afficher MM/YYYY au lieu de YYYY-MM
    const parts = m.split("-");
    th.textContent = parts.length === 2 ? `${parts[1]}/${parts[0]}` : m;
    th.title = m;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Corps
  const tbody = document.getElementById("hm-tbody");
  tbody.innerHTML = "";
  profils.forEach((profil, rowIdx) => {
    const tr = document.createElement("tr");

    const tdProfil = document.createElement("td");
    tdProfil.className   = "hm-profil-cell";
    tdProfil.textContent = profil;
    tr.appendChild(tdProfil);

    mois.forEach((m, colIdx) => {
      const count = matrix[rowIdx][colIdx];
      const td    = document.createElement("td");

      const cell = document.createElement("div");
      cell.className   = "hm-cell" + (count === 0 ? " hm-cell--zero" : "");
      cell.textContent = count > 0 ? count : "";

      if (count > 0) {
        cell.style.background = cellColor(count, max_val);
        cell.style.color      = "#fff";

        cell.addEventListener("mousemove", e => {
          showTooltip(e,
            `<strong>${escHtml(profil)}</strong> — <strong>${escHtml(m)}</strong><br>
             Écarts ≥ 40% : <span style="color:#FCA5A5;font-weight:700;">${count}</span>`
          );
          positionTooltip(e);
        });
        cell.addEventListener("mouseleave", hideTooltip);

        cell.addEventListener("click", () => {
          document.querySelectorAll(".hm-cell--selected").forEach(el =>
            el.classList.remove("hm-cell--selected")
          );
          cell.classList.add("hm-cell--selected");
          _selCell = { profil, mois: m };
          showDetail(profil, m, count);
        });
      }

      td.appendChild(cell);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  buildLegend(max_val);
  showState("hm-result");

  // Résumé chips
  const flux = document.getElementById("hm-flux").value;
  document.getElementById("hm-summary").innerHTML = `
    <span class="hm-chip hm-chip--red">⚠ ${n_ecarts} écart(s) ≥ 40%</span>
    <span class="hm-chip hm-chip--blue">${profils.length} profils · ${mois.length} mois</span>
    <span class="hm-chip hm-chip--muted">${flux === "Tous flux" ? "Tous flux" : escHtml(flux)}</span>
  `;
}

/* ── DÉTAIL CELLULE ───────────────────────────────────────── */
function showDetail(profil, moisVal, count) {
  const key      = `${profil}|||${moisVal}`;
  const details  = (_data?.details || {})[key] || [];

  document.getElementById("hm-detail-title").textContent =
    `Détail — ${profil} / ${moisVal}`;
  document.getElementById("hm-detail-sub").textContent =
    `${count} écart(s) ≥ 40% détecté(s)`;

  const body = document.getElementById("hm-detail-body");
  if (!details.length) {
    body.innerHTML = `<div class="hm-state" style="min-height:80px;">
      <span style="color:#64748B;font-size:12px;">Aucun détail disponible.</span>
    </div>`;
    return;
  }

  const rows = details.map(d => {
    const pct = d.ecart_pct;
    const cls = pct > 0 ? "neg" : "pos";   // écart absolu toujours >= 40
    return `<tr>
      <td style="text-align:right;">${fmt2(d.reel)}</td>
      <td style="text-align:right;">${fmt2(d.prev)}</td>
      <td class="${cls}" style="text-align:right;font-weight:600;">${pct.toLocaleString("fr-FR", {minimumFractionDigits:1})} %</td>
    </tr>`;
  }).join("");

  body.innerHTML = `
    <div class="hm-detail-table-wrap">
      <table class="hm-detail-table">
        <thead>
          <tr>
            <th style="text-align:right;">Réel (k€)</th>
            <th style="text-align:right;">Prévision (k€)</th>
            <th style="text-align:right;">Écart %</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  document.getElementById("hm-detail-card")
    .scrollIntoView({ behavior: "smooth", block: "start" });
}

/* ── LANCER L'ANALYSE ─────────────────────────────────────── */
async function lancerHeatmap() {
  const section = document.getElementById("hm-section").value;
  const annee   = document.getElementById("hm-annee").value;
  const flux    = document.getElementById("hm-flux").value || "Tous flux";
  if (!section || !annee) return;

  showState("hm-loading");
  document.getElementById("hm-summary").innerHTML = "";
  document.getElementById("hm-title").textContent =
    `Heatmap écarts — ${section} — ${annee}`;

  try {
    const res  = await fetch("/api/heatmap_ecarts/analyse", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ section, annee: parseInt(annee), flux }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
    renderHeatmap(data);
  } catch (err) {
    document.getElementById("hm-error-msg").textContent = err.message || "Erreur inconnue";
    showState("hm-error");
  }
}

/* ── INIT ─────────────────────────────────────────────────── */
(async () => {
  const selSection = document.getElementById("hm-section");
  const selAnnee   = document.getElementById("hm-annee");
  const selFlux    = document.getElementById("hm-flux");
  const btnLancer  = document.getElementById("btn-hm-lancer");

  btnLancer.addEventListener("click", lancerHeatmap);

  document.getElementById("btn-reset-filters")?.addEventListener("click", () => {
    if (selSection) selSection.selectedIndex = 0;
    if (selAnnee)   { selAnnee.innerHTML = '<option value="">—</option>'; selAnnee.disabled = true; }
    if (selFlux)    { selFlux.innerHTML = '<option value="Tous flux">Tous flux</option>'; selFlux.disabled = true; }
    btnLancer.disabled = true;
    window.toast?.("Filtres réinitialisés", "info");
  });

  document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
    window.pulsePDF("Heatmap-ecarts-PULSE", ".hm-layout");
  });

  document.getElementById("btn-export-excel")?.addEventListener("click", () => {
    if (!_data) { window.toast?.("Lancez d'abord la heatmap.", "error"); return; }
    const section = selSection?.value || "Section";
    /* Exporter via le premier graphique disponible */
    const chart = chartVolume || chartFreq || chartValo;
    if (chart) {
      window.pulseExcelChart(chart, `heatmap_ecarts_${section}`);
    } else {
      window.toast?.("Aucune donnée graphique à exporter.", "error");
    }
  });

  // Charger le catalogue → { section: [flux_list] }
  try {
    const res = await fetch("/api/catalogue");
    if (!res.ok) throw new Error();
    _catalogue = await res.json();

    // Peupler les sections
    Object.keys(_catalogue).sort().forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      selSection.appendChild(o);
    });

    selSection.addEventListener("change", async () => {
      const s = selSection.value;

      // Reset année
      selAnnee.innerHTML = `<option value="">—</option>`;
      selAnnee.disabled  = !s;

      // Reset flux
      selFlux.innerHTML = `<option value="Tous flux">Tous flux</option>`;
      selFlux.disabled  = !s;

      btnLancer.disabled = true;
      if (!s) return;

      // Peupler les flux
      (_catalogue[s] || []).forEach(f => {
        const o = document.createElement("option");
        o.value = f; o.textContent = f;
        selFlux.appendChild(o);
      });

      // Charger les années via l'API visualisation (premier flux disponible)
      try {
        const firstFlux = (_catalogue[s] || [])[0] || "";
        const r = await fetch(
          `/api/visualisation?section=${encodeURIComponent(s)}&flux=${encodeURIComponent(firstFlux)}`
        );
        const d = await r.json().catch(() => ({}));
        (d.annees || []).forEach(y => {
          const o = document.createElement("option");
          o.value = String(y); o.textContent = String(y);
          selAnnee.appendChild(o);
        });
        if ((d.annees || []).length) {
          const now  = new Date().getFullYear();
          const best = [...(d.annees || [])].filter(y => y <= now).pop()
                    ?? d.annees[d.annees.length - 1];
          selAnnee.value     = String(best);
          btnLancer.disabled = false;
        }
      } catch (_) {}
    });

    selAnnee.addEventListener("change", () => {
      btnLancer.disabled = !selSection.value || !selAnnee.value;
    });

  } catch (_) {
    document.getElementById("hm-error-msg").textContent =
      "Impossible de charger le catalogue.";
    showState("hm-error");
  }
})();
