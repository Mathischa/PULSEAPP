/* export_utils.js — Utilitaires d'export PULSE (PDF + Excel) */
"use strict";

/* ── PDF via jsPDF + html2canvas (téléchargement direct) ────── */
window.pulsePDF = function (title) {
  const { jsPDF } = window.jspdf;
  
  if (!window.html2canvas || !jsPDF) {
    alert("Bibliothèques PDF (html2canvas/jsPDF) non disponibles.");
    return;
  }
  
  // Récupérer l'élément principal
  const element = document.querySelector(".main") || document.body;
  
  // Masquer éléments non-imprimables
  const hideElements = [
    ".header", "#sidebar", ".sidebar", ".sidebar-toggle-btn",
    "#pulse-splash", ".filters-bar", ".filters-actions",
    ".page-header-actions", ".export-btn-group", ".ctrl-footer",
    ".btn--export-pdf", ".btn--export-excel",
    "#btn-export-pdf", "#btn-export-excel", "#btn-export", "#btn-reset-zoom",
    ".ctrl-panel"
  ];
  
  const hidden = [];
  hideElements.forEach(selector => {
    document.querySelectorAll(selector).forEach(el => {
      if (el.style.display !== "none") {
        hidden.push({ el, display: el.style.display });
        el.style.display = "none";
      }
    });
  });
  
  // Créer PDF
  const btn = document.getElementById("btn-export-pdf");
  if (btn) btn.disabled = true;
  
  html2canvas(element, {
    scale: 2,
    useCORS: true,
    allowTaint: false,
    backgroundColor: "#ffffff"
  }).then(canvas => {
    try {
      const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4"
      });
      
      const imgData = canvas.toDataURL("image/jpeg", 0.95);
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 10;
      const imgWidth = pageWidth - (margin * 2);
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      
      let heightLeft = imgHeight;
      let position = 0;
      
      // Première page
      pdf.addImage(imgData, "JPEG", margin, margin, imgWidth, imgHeight);
      heightLeft -= pageHeight - (margin * 2);
      
      // Pages suivantes
      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, "JPEG", margin, position, imgWidth, imgHeight);
        heightLeft -= pageHeight - (margin * 2);
      }
      
      // Télécharger
      pdf.save((title || "export") + ".pdf");
    } catch (e) {
      console.error("Erreur PDF:", e);
      alert("Erreur lors de la génération du PDF: " + e.message);
    } finally {
      // Restaurer éléments
      hidden.forEach(({ el, display }) => {
        el.style.display = display;
      });
      if (btn) btn.disabled = false;
    }
  }).catch(err => {
    console.error("Erreur html2canvas:", err);
    alert("Erreur lors de la capture: " + err.message);
    hidden.forEach(({ el, display }) => {
      el.style.display = display;
    });
    if (btn) btn.disabled = false;
  });
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
