/* export_utils.js — Utilitaires d'export PULSE (PDF + Excel) */
"use strict";

const _toast = (msg, type) => window.toast ? window.toast(msg, type) : console.warn(msg);

/* ── PDF via jsPDF + html2canvas ────────────────────────────────
   Principe : on adapte les dimensions du PDF au ratio réel du
   contenu capturé, sans jamais écraser l'image.
   - Paysage automatique si contenu plus large que haut
   - Page unique si le contenu tient en une page
   - Multi-page découpé proprement sinon
─────────────────────────────────────────────────────────────── */
window.pulsePDF = function (title, targetSelector) {
  const { jsPDF } = window.jspdf;

  if (!window.html2canvas || !jsPDF) {
    _toast("Bibliothèques PDF (html2canvas/jsPDF) non disponibles.", "error");
    return;
  }

  _toast("Génération du PDF en cours…", "info");

  // Élément à capturer : priorité au sélecteur passé, sinon .content, sinon body
  const element =
    (targetSelector && document.querySelector(targetSelector)) ||
    document.querySelector(".content") ||
    document.querySelector("main") ||
    document.body;

  // Masquer les éléments non imprimables
  const hideSelectors = [
    ".header", "#sidebar", ".sidebar", ".sidebar-toggle-btn",
    "#pulse-splash", ".filters-bar", ".filters-actions",
    ".page-header-actions", ".export-btn-group", ".ctrl-footer",
    ".btn--export-pdf", ".btn--export-excel",
    "#btn-export-pdf", "#btn-export-excel", "#btn-export", "#btn-reset-zoom",
    ".ctrl-panel",
  ];
  const hidden = [];
  hideSelectors.forEach(sel => {
    document.querySelectorAll(sel).forEach(el => {
      if (el.style.display !== "none") {
        hidden.push({ el, prev: el.style.display });
        el.style.display = "none";
      }
    });
  });

  const btn = document.querySelector("#btn-export-pdf, .btn--export-pdf");
  if (btn) { btn.disabled = true; btn.classList.add("loading"); }

  const restore = () => {
    hidden.forEach(({ el, prev }) => { el.style.display = prev; });
    if (btn) { btn.disabled = false; btn.classList.remove("loading"); }
  };

  html2canvas(element, {
    scale: 2,
    useCORS: true,
    allowTaint: false,
    backgroundColor: "#07090F",   // fond sombre de l'appli
    logging: false,
    // Capturer le scroll complet de l'élément
    scrollX: 0,
    scrollY: -window.scrollY,
    windowWidth:  element.scrollWidth,
    windowHeight: element.scrollHeight,
  }).then(canvas => {
    try {
      // ── Ratio réel du canvas capturé ──────────────────────────
      const cW = canvas.width;
      const cH = canvas.height;
      const ratio = cW / cH;

      // ── Choix orientation selon le ratio ──────────────────────
      const landscape = ratio > 1.0;
      const pdf = new jsPDF({
        orientation: landscape ? "landscape" : "portrait",
        unit: "mm",
        format: "a4",
      });

      const margin   = 8;                                // mm
      const pageW    = pdf.internal.pageSize.getWidth();
      const pageH    = pdf.internal.pageSize.getHeight();
      const usableW  = pageW  - margin * 2;
      const usableH  = pageH  - margin * 2;

      // Calculer les dimensions de l'image dans la page
      // On contraint à la largeur et on recalcule la hauteur
      const imgW = usableW;
      const imgH = imgW / ratio;          // hauteur proportionnelle RÉELLE

      const imgData = canvas.toDataURL("image/jpeg", 0.92);

      if (imgH <= usableH) {
        // ── Cas 1 : tient en une seule page, centré verticalement ──
        const offsetY = margin + (usableH - imgH) / 2;
        pdf.addImage(imgData, "JPEG", margin, offsetY, imgW, imgH);
      } else {
        // ── Cas 2 : multi-page — on découpe l'image en tranches ──
        // chaque tranche = usableH en coordonnées image
        const sliceH   = usableH * (cW / usableW);   // hauteur d'une tranche en px canvas
        let srcY = 0;
        let page = 0;

        while (srcY < cH) {
          if (page > 0) pdf.addPage();

          // Créer un canvas temporaire pour la tranche
          const sliceCanvas = document.createElement("canvas");
          sliceCanvas.width  = cW;
          sliceCanvas.height = Math.min(sliceH, cH - srcY);
          const sliceCtx = sliceCanvas.getContext("2d");
          sliceCtx.drawImage(canvas,
            0, srcY, cW, sliceCanvas.height,    // source
            0, 0,    cW, sliceCanvas.height     // dest
          );

          const sliceData  = sliceCanvas.toDataURL("image/jpeg", 0.92);
          const sliceImgH  = sliceCanvas.height * (usableW / cW);
          pdf.addImage(sliceData, "JPEG", margin, margin, usableW, sliceImgH);

          srcY += sliceCanvas.height;
          page++;
        }
      }

      // ── Téléchargement ────────────────────────────────────────
      const filename = (title || "export") + ".pdf";
      const pdfBlob  = pdf.output("blob");
      const url      = URL.createObjectURL(pdfBlob);
      const link     = document.createElement("a");
      link.href      = url;
      link.download  = filename;
      link.style.display = "none";
      document.body.appendChild(link);
      link.click();
      setTimeout(() => { URL.revokeObjectURL(url); link.remove(); }, 200);
      _toast("PDF exporté avec succès !", "success");

    } catch (e) {
      console.error("[PULSE PDF]", e);
      _toast("Erreur lors de la génération du PDF : " + e.message, "error");
    } finally {
      restore();
    }
  }).catch(err => {
    console.error("[PULSE PDF html2canvas]", err);
    _toast("Erreur lors de la capture : " + err.message, "error");
    restore();
  });
};

/* ── Export PDF d'un graphique Chart.js (direct canvas → PDF) ──
   Exporte le canvas Chart.js directement en PDF en conservant
   exactement le ratio width/height du canvas — aucune distorsion.
   chartOrCanvasId : instance Chart.js OU id du canvas HTML
─────────────────────────────────────────────────────────────── */
window.pulseChartPDF = function (chartOrCanvasId, title) {
  const { jsPDF } = window.jspdf;
  if (!jsPDF) { _toast("jsPDF non disponible.", "error"); return; }

  _toast("Génération du PDF en cours…", "info");

  // Récupérer le canvas
  let canvas;
  if (typeof chartOrCanvasId === "string") {
    const el = document.getElementById(chartOrCanvasId);
    canvas = el?.tagName === "CANVAS" ? el : el?.querySelector("canvas");
  } else if (chartOrCanvasId?.canvas) {
    canvas = chartOrCanvasId.canvas;           // instance Chart.js
  } else if (chartOrCanvasId?.tagName === "CANVAS") {
    canvas = chartOrCanvasId;
  }

  // Fallback : prendre le premier canvas visible dans #chart-wrap
  if (!canvas) {
    canvas = document.querySelector("#chart-wrap canvas, .chart-wrap canvas, canvas");
  }
  if (!canvas) { _toast("Canvas introuvable.", "error"); return; }

  const cW = canvas.width;
  const cH = canvas.height;
  if (!cW || !cH) { _toast("Le graphique n'a pas encore été rendu.", "error"); return; }

  const ratio     = cW / cH;
  const landscape = ratio > 1.0;

  const pdf      = new jsPDF({ orientation: landscape ? "landscape" : "portrait", unit: "mm", format: "a4" });
  const margin   = 10;
  const pageW    = pdf.internal.pageSize.getWidth();
  const pageH    = pdf.internal.pageSize.getHeight();
  const usableW  = pageW - margin * 2;
  const usableH  = pageH - margin * 2;

  const imgW     = usableW;
  const imgH     = imgW / ratio;
  const offsetY  = margin + Math.max(0, (usableH - imgH) / 2);

  // Fond sombre pour correspondre au thème
  pdf.setFillColor(7, 9, 15);
  pdf.rect(0, 0, pageW, pageH, "F");

  const imgData = canvas.toDataURL("image/png");
  pdf.addImage(imgData, "PNG", margin, offsetY, imgW, Math.min(imgH, usableH));

  const filename = (title || "export") + ".pdf";
  const blob     = pdf.output("blob");
  const url      = URL.createObjectURL(blob);
  const link     = document.createElement("a");
  link.href = url; link.download = filename; link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  setTimeout(() => { URL.revokeObjectURL(url); link.remove(); }, 200);
  _toast("PDF exporté avec succès !", "success");
};

/* ── Export PDF d'un graphique Plotly (qualité max via SVG) ────
   Utilise l'API native Plotly.downloadImage pour un rendu parfait
   sans distorsion html2canvas.
─────────────────────────────────────────────────────────────── */
window.pulsePlotlyPDF = function (divId, title, widthPx, heightPx) {
  const el = document.getElementById(divId);
  if (!el || !window.Plotly) {
    // Fallback vers html2canvas si Plotly absent
    window.pulsePDF(title, "#" + divId);
    return;
  }
  _toast("Génération du PDF en cours…", "info");
  Plotly.downloadImage(el, {
    format:   "pdf",
    filename: title || "export",
    width:    widthPx  || Math.round(el.offsetWidth  * 1.5) || 1600,
    height:   heightPx || Math.round(el.offsetHeight * 1.5) || 900,
  }).then(() => {
    _toast("PDF exporté avec succès !", "success");
  }).catch(() => {
    // Plotly PDF non supporté dans ce browser → fallback PNG via html2canvas
    window.pulsePDF(title, "#" + divId);
  });
};

/* ── Export PNG haute résolution d'un graphique Plotly ──────── */
window.pulsePlotlyPNG = function (divId, title, widthPx, heightPx) {
  const el = document.getElementById(divId);
  if (!el || !window.Plotly) return;
  _toast("Export PNG en cours…", "info");
  Plotly.downloadImage(el, {
    format:   "png",
    filename: title || "export",
    width:    widthPx  || Math.round(el.offsetWidth  * 2) || 1920,
    height:   heightPx || Math.round(el.offsetHeight * 2) || 1080,
    scale:    2,
  }).then(() => _toast("PNG exporté !", "success"));
};

/* ── Excel depuis un tableau HTML ───────────────────────────── */
window.pulseExcelTable = function (tableId, filename) {
  if (typeof XLSX === "undefined") {
    _toast("Bibliothèque Excel (SheetJS) non disponible.", "error");
    return;
  }
  const el = document.getElementById(tableId);
  if (!el) { _toast("Tableau introuvable : " + tableId, "error"); return; }

  _toast("Export Excel en cours…", "info");
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(el);

  if (ws["!cols"]) {
    ws["!cols"].forEach(col => { col.wch = Math.max(col.wch || 10, 12); });
  }

  XLSX.utils.book_append_sheet(wb, ws, "Données");
  XLSX.writeFile(wb, (filename || "export") + ".xlsx");
  _toast("Excel exporté avec succès !", "success");
};

/* ── Excel depuis une instance Chart.js ─────────────────────── */
window.pulseExcelChart = function (chart, filename, sheetName) {
  if (typeof XLSX === "undefined") { _toast("SheetJS non disponible.", "error"); return; }
  if (!chart?.data) { _toast("Aucune donnée disponible.", "error"); return; }

  _toast("Export Excel en cours…", "info");
  const labels   = chart.data.labels   || [];
  const datasets = chart.data.datasets || [];
  const headers  = ["", ...datasets.map(d => d.label || "Série")];
  const rows     = labels.map((lbl, i) => [
    lbl,
    ...datasets.map(d => d.data[i] != null ? d.data[i] : ""),
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = [{ wch: 15 }, ...datasets.map(() => ({ wch: 14 }))];

  for (let col = 0; col < headers.length; col++) {
    const ref = XLSX.utils.encode_col(col) + "1";
    if (ws[ref]) ws[ref].s = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "366092" } },
      alignment: { horizontal: "center" },
    };
  }

  XLSX.utils.book_append_sheet(wb, ws, sheetName || "Données");
  XLSX.writeFile(wb, (filename || "export") + ".xlsx");
  _toast("Excel exporté avec succès !", "success");
};

/* ── Excel depuis tableau de données (AOA) ───────────────────── */
window.pulseExcelData = function (headers, rows, filename, sheetName) {
  if (typeof XLSX === "undefined") { _toast("SheetJS non disponible.", "error"); return; }

  _toast("Export Excel en cours…", "info");
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = headers.map(() => ({ wch: 14 }));

  for (let col = 0; col < headers.length; col++) {
    const ref = XLSX.utils.encode_col(col) + "1";
    if (ws[ref]) ws[ref].s = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "366092" } },
      alignment: { horizontal: "center" },
    };
  }

  XLSX.utils.book_append_sheet(wb, ws, sheetName || "Données");
  XLSX.writeFile(wb, (filename || "export") + ".xlsx");
  _toast("Excel exporté avec succès !", "success");
};

/* ── Activer/désactiver les boutons export ───────────────────── */
window.pulseExportReady = function (enabled) {
  ["btn-export-pdf", "btn-export-excel"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.disabled = !enabled;
  });
};
