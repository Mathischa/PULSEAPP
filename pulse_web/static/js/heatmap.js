/* heatmap.js */
"use strict";

/* ── ÉTAT ─────────────────────────────────────────────────── */
let _data      = null;   // réponse API complète
let _selCell   = null;   // { profil, flux }
let _catalogue = {};

/* ── UTILS ────────────────────────────────────────────────── */
const fmt  = v => typeof v === "number" && isFinite(v)
  ? v.toLocaleString("fr-FR", { maximumFractionDigits: 0 }) : "—";
const fmt2 = v => typeof v === "number" && isFinite(v)
  ? v.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "—";
function escHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

/* ── COULEUR HEATMAP ──────────────────────────────────────── */
function cellColor(value, maxVal) {
  if (value === 0 || maxVal === 0) return null;
  const t = Math.min(value / maxVal, 1);
  // Gradient bleu-clair → rouge, en passant par orange
  const stops = [
    [76, 124, 243],   // bleu  (t=0)
    [251, 146, 60],   // orange (t=0.5)
    [239, 68,  68],   // rouge  (t=1)
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
function showHmState(id) {
  ["hm-idle","hm-loading","hm-error","hm-result"].forEach(s => {
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
  const x = e.clientX + 14;
  const y = e.clientY + 14;
  tooltip.style.left = Math.min(x, window.innerWidth - tooltip.offsetWidth - 10) + "px";
  tooltip.style.top  = Math.min(y, window.innerHeight - tooltip.offsetHeight - 10) + "px";
}
function hideTooltip() { tooltip.style.display = "none"; }

/* ── LÉGENDE ──────────────────────────────────────────────── */
function buildLegend(maxVal) {
  const scale = document.getElementById("hm-legend-scale");
  const steps = 20;
  scale.innerHTML = "";
  for (let i = 0; i <= steps; i++) {
    const v = (i / steps) * maxVal;
    const div = document.createElement("div");
    div.style.flex = "1";
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

  const { profils, flux, matrix, mean_matrix, max_val } = data;

  // En-têtes
  const thead = document.getElementById("hm-thead");
  thead.innerHTML = "";
  const headerRow = document.createElement("tr");
  const thCorner  = document.createElement("th");
  thCorner.className = "hm-profil-header";
  thCorner.textContent = "Profil \\ Flux";
  headerRow.appendChild(thCorner);
  flux.forEach(f => {
    const th = document.createElement("th");
    // Tronquer les noms longs
    th.title       = f;
    th.textContent = f.length > 18 ? f.slice(0, 17) + "…" : f;
    th.style.maxWidth = "90px";
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Corps
  const tbody = document.getElementById("hm-tbody");
  tbody.innerHTML = "";
  profils.forEach((profil, rowIdx) => {
    const tr = document.createElement("tr");

    // Cellule profil (sticky gauche)
    const tdProfil = document.createElement("td");
    tdProfil.className   = "hm-profil-cell";
    tdProfil.textContent = profil;
    tr.appendChild(tdProfil);

    flux.forEach((f, colIdx) => {
      const count    = matrix[rowIdx][colIdx];
      const meanAbs  = mean_matrix[rowIdx][colIdx];
      const td       = document.createElement("td");

      const cell = document.createElement("div");
      cell.className = "hm-cell" + (count === 0 ? " hm-cell--zero" : "");
      cell.textContent = count > 0 ? count : "";

      if (count > 0) {
        cell.style.background = cellColor(count, max_val);
        cell.style.color = "#fff";

        // Tooltip survol
        cell.addEventListener("mousemove", e => {
          showTooltip(e,
            `<strong>${escHtml(profil)}</strong> / <strong>${escHtml(f)}</strong><br>
             Anomalies : <span style="color:#FCA5A5;font-weight:700;">${count}</span><br>
             Écart moy. abs. : ${fmt(meanAbs)} k€`
          );
          positionTooltip(e);
        });
        cell.addEventListener("mouseleave", hideTooltip);

        // Clic → afficher détail
        cell.addEventListener("click", () => {
          // Désélectionner l'ancienne cellule
          document.querySelectorAll(".hm-cell--selected").forEach(el =>
            el.classList.remove("hm-cell--selected")
          );
          cell.classList.add("hm-cell--selected");
          _selCell = { profil, flux: f };
          showDetail(profil, f, count);
        });
      }

      td.appendChild(cell);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  buildLegend(max_val);
  showHmState("hm-result");

  // Résumé
  const anom_pct = data.n_total > 0
    ? ((data.n_anomalies / data.n_total) * 100).toFixed(1)
    : "0.0";
  document.getElementById("hm-summary").innerHTML = `
    <span class="hm-chip hm-chip--red">🔴 ${data.n_anomalies} anomalie(s)</span>
    <span class="hm-chip hm-chip--blue">${data.n_total} points analysés</span>
    <span class="hm-chip hm-chip--muted">${anom_pct}% taux anomalie</span>
    <span class="hm-chip hm-chip--muted">${profils.length} profils · ${flux.length} flux</span>
  `;
}

/* ── DÉTAIL CELLULE ───────────────────────────────────────── */
function showDetail(profil, fluxName, count) {
  const key     = `${profil}|||${fluxName}`;
  const anomRows = (_data?.details || {})[key] || [];

  document.getElementById("hm-detail-title").textContent =
    `Détail — ${profil} / ${fluxName}`;
  document.getElementById("hm-detail-sub").textContent =
    `${count} anomalie(s) détectée(s)`;

  const body = document.getElementById("hm-detail-body");
  if (!anomRows.length) {
    body.innerHTML = `<div class="hm-state" style="min-height:80px;">
      <span style="color:#64748B;font-size:12px;">Aucun détail disponible.</span>
    </div>`;
    return;
  }

  const rows = anomRows.map(r => {
    const cls     = r.favorable ? "pos" : "neg";
    const ecartFmt = (r.ecart >= 0 ? "+" : "") + fmt(r.ecart);
    return `<tr>
      <td>${escHtml(r.date)}</td>
      <td>${escHtml(fluxName)}</td>
      <td>${escHtml(profil)}</td>
      <td style="text-align:right;">${fmt(r.reel)}</td>
      <td style="text-align:right;">${fmt(r.prev)}</td>
      <td class="${cls}" style="text-align:right;font-weight:600;">${ecartFmt}</td>
    </tr>`;
  }).join("");

  body.innerHTML = `
    <div class="hm-detail-table-wrap">
      <table class="hm-detail-table">
        <thead>
          <tr>
            <th>Date</th><th>Flux</th><th>Profil</th>
            <th style="text-align:right;">Réel (k€)</th>
            <th style="text-align:right;">Prévision (k€)</th>
            <th style="text-align:right;">Écart (k€)</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  // Scroll vers le détail
  document.getElementById("hm-detail-card").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* ── LANCER L'ANALYSE ─────────────────────────────────────── */
async function lancerHeatmap() {
  const section = document.getElementById("hm-section").value;
  const annee   = document.getElementById("hm-annee").value;
  if (!section || !annee) return;

  showHmState("hm-loading");
  document.getElementById("hm-summary").innerHTML = "";
  document.getElementById("hm-title").textContent =
    `Heatmap — ${section} — ${annee}`;

  try {
    const res  = await fetch("/api/heatmap/analyse", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ section, annee }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
    renderHeatmap(data);
  } catch (err) {
    document.getElementById("hm-error-msg").textContent = err.message || "Erreur inconnue";
    showHmState("hm-error");
  }
}

/* ── INIT ─────────────────────────────────────────────────── */
(async () => {
  document.getElementById("btn-hm-lancer").addEventListener("click", lancerHeatmap);

  document.getElementById("btn-export-pdf")?.addEventListener("click", () => {
    window.pulsePDF("Heatmap des anomalies — PULSE");
  });

  document.getElementById("btn-export-excel")?.addEventListener("click", () => {
    if (!_data) { alert("Lancez d'abord la heatmap."); return; }
    const section  = document.getElementById("hm-section")?.value || "Section";
    const allProfils = Object.keys(_data);
    const allFlux    = [...new Set(allProfils.flatMap(p => Object.keys(_data[p] || {})))].sort();
    const headers    = ["Profil", ...allFlux];
    const rows = allProfils.map(p => [p, ...allFlux.map(f => _data[p]?.[f] ?? 0)]);
    window.pulseExcelData(headers, rows, `heatmap_anomalies_${section}`);
  });

  // Charger le catalogue
  try {
    const res = await fetch("/api/catalogue");
    if (!res.ok) throw new Error();
    _catalogue = await res.json();

    const selSection = document.getElementById("hm-section");
    const selAnnee   = document.getElementById("hm-annee");

    Object.keys(_catalogue).sort().forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      selSection.appendChild(o);
    });

    selSection.addEventListener("change", async () => {
      const s = selSection.value;
      selAnnee.innerHTML = `<option value="">—</option>`;
      selAnnee.disabled  = !s;
      document.getElementById("btn-hm-lancer").disabled = true;
      if (!s) return;

      // Charger les années depuis l'API visualisation (premier flux)
      try {
        const firstFlux = (_catalogue[s] || [])[0] || "";
        const r = await fetch(`/api/visualisation?section=${encodeURIComponent(s)}&flux=${encodeURIComponent(firstFlux)}`);
        const d = await r.json().catch(() => ({}));
        (d.annees || []).forEach(y => {
          const o = document.createElement("option");
          o.value = String(y); o.textContent = String(y);
          selAnnee.appendChild(o);
        });
        if ((d.annees || []).length) {
          const now  = new Date().getFullYear();
          const best = [...(d.annees || [])].filter(y => y < now).pop()
                    ?? d.annees[d.annees.length - 1];
          selAnnee.value = String(best);
          document.getElementById("btn-hm-lancer").disabled = false;
        }
      } catch (_) {}
    });

    selAnnee.addEventListener("change", () => {
      document.getElementById("btn-hm-lancer").disabled =
        !selSection.value || !selAnnee.value;
    });

  } catch (_) {
    document.getElementById("hm-error-msg").textContent = "Impossible de charger le catalogue.";
    showHmState("hm-error");
  }
})();
