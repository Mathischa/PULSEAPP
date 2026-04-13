/* export_utils.js — Utilitaires d'export PULSE (PDF + Excel) */
"use strict";

/* ── PDF via impression navigateur ──────────────────────────── */
window.pulsePDF = function (title) {
  if (title) {
    const prev = document.title;
    document.title = title;
  }
  
  // Injecter les styles d'impression pour améliorer la qualité
  const printStyle = document.createElement("style");
  printStyle.media = "print";
  printStyle.innerHTML = `
    @media print {
      body { background: white; color: black; }
      canvas { max-height: 100% !important; width: 100% !important; }
      .chart-container { page-break-inside: avoid; }
      table { border-collapse: collapse; width: 100%; }
      th, td { border: 1px solid #333; padding: 8px; text-align: left; }
      th { background-color: #f0f0f0; font-weight: bold; }
      h1, h2, h3 { margin-top: 15px; margin-bottom: 10px; }
      .no-print { display: none !important; }
      img { max-width: 100%; }
    }
  `;
  document.head.appendChild(printStyle);
  
  window.print();
  
  if (title) {
    document.title = prev;
  }
  
  // Nettoyer le style après l'impression
  setTimeout(() => document.head.removeChild(printStyle), 100);
};

/* ── Excel depuis un tableau HTML ───────────────────────────── */
window.pulseExcelTable = function (tableId, filename) {
  if (typeof XLSX === "undefined") {
    alert("Bibliothèque Excel (SheetJS) non disponible.");
    return;
  }
  const el = document.getElementById(tableId);
  if (!el) { alert("Tableau introuvable : " + tableId); return; }
  
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(el);
  
  // Améliorer le formatage: définir la largeur des colonnes
  if (ws['!cols']) {
    ws['!cols'].forEach((col, i) => { col.wch = Math.max(col.wch || 10, 12); });
  }
  
  XLSX.utils.book_append_sheet(wb, ws, "Données");
  XLSX.writeFile(wb, (filename || "export") + ".xlsx");
};

/* ── Excel depuis une instance Chart.js ─────────────────────── */
window.pulseExcelChart = function (chart, filename, sheetName) {
  if (typeof XLSX === "undefined") {
    alert("Bibliothèque Excel (SheetJS) non disponible.");
    return;
  }
  if (!chart || !chart.data) {
    alert("Aucune donnée de graphique disponible.");
    return;
  }
  const labels   = chart.data.labels   || [];
  const datasets = chart.data.datasets || [];
  const headers  = ["", ...datasets.map((d) => d.label || "Série")];
  const rows     = labels.map((lbl, i) => [
    lbl,
    ...datasets.map((d) => {
      const v = d.data[i];
      return v !== null && v !== undefined ? v : "";
    }),
  ]);
  
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  
  // Améliorer le formatage: en-têtes gras, largeur des colonnes
  ws['!cols'] = [{ wch: 15 }, ...datasets.map(() => ({ wch: 14 }))];
  
  // Formater les en-têtes en gras
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_col(col) + "1";
    if (ws[cellRef]) {
      ws[cellRef].s = { 
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "366092" } },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }
  }
  
  XLSX.utils.book_append_sheet(wb, ws, sheetName || "Données");
  XLSX.writeFile(wb, (filename || "export") + ".xlsx");
};

/* ── Excel depuis tableau de données (AOA) ───────────────────── */
window.pulseExcelData = function (headers, rows, filename, sheetName) {
  if (typeof XLSX === "undefined") {
    alert("Bibliothèque Excel (SheetJS) non disponible.");
    return;
  }
  
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  
  // Améliorer le formatage: en-têtes gras, largeur des colonnes
  ws['!cols'] = headers.map(() => ({ wch: 14 }));
  
  // Formater les en-têtes en gras
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_col(col) + "1";
    if (ws[cellRef]) {
      ws[cellRef].s = { 
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "366092" } },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }
  }
  
  XLSX.utils.book_append_sheet(wb, ws, sheetName || "Données");
  XLSX.writeFile(wb, (filename || "export") + ".xlsx");
};

/* ── Activer/désactiver les boutons export ───────────────────── */
window.pulseExportReady = function (enabled) {
  ["btn-export-pdf", "btn-export-excel"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.disabled = !enabled;
  });
};
