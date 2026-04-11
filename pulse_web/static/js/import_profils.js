/* import_profils.js — Import des profils Tréso SNCF */
"use strict";

// ── État global ─────────────────────────────────────────────────────────────
let _files        = [];   // [{name, path}] — liste courante (modifiable)
let _reelFound    = false;
let _pollInterval = null;

// ── Éléments DOM ────────────────────────────────────────────────────────────
const selYear       = document.getElementById("sel-year");
const inpFolder     = document.getElementById("inp-folder");
const btnBrowse     = document.getElementById("btn-browse");
const btnScan       = document.getElementById("btn-scan");
const btnReset      = document.getElementById("btn-reset");
const btnLaunch     = document.getElementById("btn-launch");
const btnDeleteSel  = document.getElementById("btn-delete-sel");
const btnDeleteSel2 = document.getElementById("btn-delete-sel-2");

const filesEmptyState  = document.getElementById("files-empty-state");
const filesLoading     = document.getElementById("files-loading");
const filesTableWrap   = document.getElementById("files-table-wrap");
const filesTbody       = document.getElementById("files-tbody");
const filesCount       = document.getElementById("files-count");
const checkAll         = document.getElementById("check-all");

const reelStatus       = document.getElementById("reel-status");
const cardProgress     = document.getElementById("card-progress");
const progressBarFill  = document.getElementById("progress-bar-fill");
const progressPct      = document.getElementById("progress-pct");
const progressMsg      = document.getElementById("progress-msg");
const progressResult   = document.getElementById("progress-result");


// ── Init année ───────────────────────────────────────────────────────────────
(function populateYears() {
  const currentYear = new Date().getFullYear();
  for (let y = 2018; y <= 2035; y++) {
    const opt = document.createElement("option");
    opt.value       = y;
    opt.textContent = y;
    if (y === currentYear) opt.selected = true;
    selYear.appendChild(opt);
  }
})();


// ── Rendu du tableau ─────────────────────────────────────────────────────────
function renderTable() {
  filesTbody.innerHTML = "";

  if (_files.length === 0) {
    showEmpty();
    return;
  }

  _files.forEach((file, idx) => {
    const tr = document.createElement("tr");
    tr.dataset.idx = idx;

    const tdCheck = document.createElement("td");
    const cb = document.createElement("input");
    cb.type      = "checkbox";
    cb.className = "ip-check";
    cb.dataset.idx = idx;
    cb.addEventListener("change", onCheckChange);
    tdCheck.appendChild(cb);
    tr.appendChild(tdCheck);

    const tdName = document.createElement("td");
    tdName.innerHTML = `
      <div class="ip-filename">${escHtml(file.name)}</div>
      <div class="ip-filepath">${escHtml(file.path)}</div>
    `;
    tr.appendChild(tdName);

    filesTbody.appendChild(tr);
  });

  // Afficher le tableau
  filesEmptyState.hidden  = true;
  filesLoading.hidden     = true;
  filesTableWrap.hidden   = false;
  filesCount.hidden       = false;
  filesCount.textContent  = _files.length;

  // Mettre à jour les boutons suppression
  updateDeleteBtns();
  syncCheckAll();
}


function showEmpty() {
  filesEmptyState.hidden  = false;
  filesLoading.hidden     = true;
  filesTableWrap.hidden   = true;
  filesCount.hidden       = true;
  btnDeleteSel.hidden     = true;
  btnDeleteSel2.hidden    = true;
}


function showLoading() {
  filesEmptyState.hidden = true;
  filesLoading.hidden    = false;
  filesTableWrap.hidden  = true;
  filesCount.hidden      = true;
}


// ── Checkbox "tout sélectionner" ────────────────────────────────────────────
checkAll.addEventListener("change", () => {
  const checked = checkAll.checked;
  document.querySelectorAll("#files-tbody .ip-check").forEach(cb => {
    cb.checked = checked;
  });
  updateDeleteBtns();
});


function onCheckChange() {
  syncCheckAll();
  updateDeleteBtns();
}


function syncCheckAll() {
  const cbs = [...document.querySelectorAll("#files-tbody .ip-check")];
  if (cbs.length === 0) { checkAll.checked = false; return; }
  const allChecked = cbs.every(cb => cb.checked);
  const someChecked = cbs.some(cb => cb.checked);
  checkAll.checked       = allChecked;
  checkAll.indeterminate = someChecked && !allChecked;
}


function updateDeleteBtns() {
  const hasChecked = [...document.querySelectorAll("#files-tbody .ip-check")]
    .some(cb => cb.checked);
  btnDeleteSel.hidden  = !hasChecked;
  btnDeleteSel2.hidden = !hasChecked;
}


// ── Supprimer la sélection ───────────────────────────────────────────────────
function deleteSelected() {
  const toRemove = new Set(
    [...document.querySelectorAll("#files-tbody .ip-check")]
      .filter(cb => cb.checked)
      .map(cb => parseInt(cb.dataset.idx, 10))
  );
  _files = _files.filter((_, idx) => !toRemove.has(idx));
  checkAll.checked       = false;
  checkAll.indeterminate = false;
  renderTable();
}

btnDeleteSel.addEventListener("click",  deleteSelected);
btnDeleteSel2.addEventListener("click", deleteSelected);


// ── Parcourir le dossier ─────────────────────────────────────────────────────
btnBrowse.addEventListener("click", async () => {
  btnBrowse.disabled = true;
  const origText = btnBrowse.innerHTML;
  btnBrowse.innerHTML = `<div class="spinner" style="width:12px;height:12px;border-width:2px;"></div>`;

  try {
    const resp = await fetch("/api/import_profils/browse_folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const data = await resp.json();

    if (data.folder) {
      inpFolder.value = data.folder;
    }
    // Si annulé (data.folder === null), on ne fait rien
  } catch (err) {
    // Réseau KO, on ignore silencieusement
  } finally {
    btnBrowse.disabled = false;
    btnBrowse.innerHTML = origText;
  }
});


// ── Scan ─────────────────────────────────────────────────────────────────────
btnScan.addEventListener("click", async () => {
  const year   = parseInt(selYear.value, 10);
  const folder = inpFolder.value.trim();

  if (!folder) {
    setReelStatus("Veuillez saisir un chemin de dossier.", "warning");
    return;
  }

  showLoading();
  setReelStatus("Scan en cours…", "neutral");
  btnScan.disabled   = true;
  btnLaunch.disabled = true;
  _files             = [];
  _reelFound         = false;

  try {
    const resp = await fetch("/api/import_profils/scan", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ year, folder }),
    });

    const data = await resp.json();

    if (!resp.ok) {
      setReelStatus(data.error || "Erreur lors du scan.", "warning");
      showEmpty();
      btnScan.disabled = false;
      return;
    }

    _files = data.files || [];
    renderTable();

    if (_files.length === 0) {
      setReelStatus(
        `Aucun fichier profil trouvé pour ${year} dans ce dossier.`,
        "warning"
      );
    }

    // Statut fichier réel
    if (data.fichier_reel) {
      _reelFound = true;
      setReelStatus(
        `Trouvé : ${data.fichier_reel}`,
        "success"
      );
    } else {
      _reelFound = false;
      setReelStatus(
        `Fichier "Réel ${year}.xlsx" introuvable dans :\n${data.base_donnees_dir}`,
        "warning"
      );
    }

  } catch (err) {
    setReelStatus(`Erreur réseau : ${err.message}`, "warning");
    showEmpty();
  } finally {
    btnScan.disabled   = false;
    btnLaunch.disabled = !(_reelFound && _files.length > 0);
  }
});


// ── Réinitialiser ────────────────────────────────────────────────────────────
btnReset.addEventListener("click", () => {
  _files        = [];
  _reelFound    = false;
  checkAll.checked       = false;
  checkAll.indeterminate = false;
  showEmpty();
  setReelStatus("En attente d'un scan…", "neutral");
  btnLaunch.disabled  = true;
  btnDeleteSel.hidden  = true;
  btnDeleteSel2.hidden = true;
  stopPolling();
  cardProgress.hidden = true;
  progressBarFill.style.width = "0%";
  progressPct.textContent     = "0 %";
  progressMsg.textContent     = "Initialisation…";
  progressResult.hidden       = true;
});


// ── Lancer l'import ──────────────────────────────────────────────────────────
btnLaunch.addEventListener("click", async () => {
  if (!_reelFound || _files.length === 0) return;

  const year  = parseInt(selYear.value, 10);
  const paths = _files.map(f => f.path);

  btnLaunch.disabled = true;
  btnScan.disabled   = true;
  cardProgress.hidden = false;
  progressResult.hidden = true;

  setProgress(0, "Envoi de la requête d'import…");

  try {
    const resp = await fetch("/api/import_profils/launch", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ year, files: paths }),
    });

    const data = await resp.json();

    if (!resp.ok) {
      setProgress(0, data.error || "Erreur au lancement.");
      showResultMsg(data.error || "Erreur au lancement.", "error");
      btnLaunch.disabled = false;
      btnScan.disabled   = false;
      return;
    }

    const jobId = data.job_id;
    startPolling(jobId);

  } catch (err) {
    setProgress(0, `Erreur réseau : ${err.message}`);
    showResultMsg(`Erreur réseau : ${err.message}`, "error");
    btnLaunch.disabled = false;
    btnScan.disabled   = false;
  }
});


// ── Polling ──────────────────────────────────────────────────────────────────
function startPolling(jobId) {
  stopPolling();
  _pollInterval = setInterval(async () => {
    try {
      const resp = await fetch(`/api/import_profils/progress/${jobId}`, {
        cache: "no-store",
      });

      if (!resp.ok) {
        stopPolling();
        showResultMsg("Impossible de récupérer la progression.", "error");
        return;
      }

      const data = await resp.json();

      setProgress(data.progress || 0, data.message || "");

      if (data.done) {
        stopPolling();
        if (data.error) {
          showResultMsg(`Erreur : ${data.error}`, "error");
        } else {
          showResultMsg("Import terminé avec succès !", "success");
          if (window.toast) toast("Import terminé avec succès !", "success");
        }
        btnScan.disabled   = false;
        btnLaunch.disabled = false;
      }
    } catch (err) {
      stopPolling();
      showResultMsg(`Erreur de connexion : ${err.message}`, "error");
      btnScan.disabled   = false;
      btnLaunch.disabled = false;
    }
  }, 800);
}


function stopPolling() {
  if (_pollInterval !== null) {
    clearInterval(_pollInterval);
    _pollInterval = null;
  }
}


// ── Helpers UI ───────────────────────────────────────────────────────────────
function setProgress(pct, msg) {
  const clamped = Math.max(0, Math.min(100, pct));
  progressBarFill.style.width = `${clamped}%`;
  progressPct.textContent     = `${clamped} %`;
  if (msg) progressMsg.textContent = msg;
}


function setReelStatus(msg, tone) {
  reelStatus.className = `ip-status ip-status--${tone}`;
  reelStatus.textContent = msg;
}


function showResultMsg(msg, tone) {
  progressResult.className   = `ip-result-msg ip-result-msg--${tone}`;
  progressResult.textContent = msg;
  progressResult.hidden      = false;
}


function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
