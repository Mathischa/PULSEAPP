# -*- coding: utf-8 -*-
"""
api/heatmap.py — Heatmap des anomalies par profil × flux.

POST /api/heatmap/analyse  { section, annee }
    → matrice (profils × flux), anomalies, mean écart abs
"""
from __future__ import annotations
import math, re
from collections import defaultdict
from flask import Blueprint, jsonify, request
from pulse_v2.data.cache import CACHE, TOKENS

bp = Blueprint("heatmap", __name__, url_prefix="/api/heatmap")

FLUX_A_EXCLURE = {
    "Cash flow de financement", "Cash flow net", "Sous total financier",
    "Sous total Investissements nets et ACE", "Free cash Flow",
    "Sous total recettes", "Sous total dépenses",
}

ENCAISSEMENTS = {
    "Trafic Voyageurs", "Subventions", "Redevances d'infrastructure",
    "Enc. Autres Produits", "Sous total recettes", "Subventions d'investissements"
}
DECAISSEMENTS = {"Péages", "Charges de personnel", "ACE & Investissements"}


def _profil_sort_key(name: str) -> tuple:
    s = str(name).strip()
    m = re.search(r"(\d{1,2})[/-](\d{1,2})", s)
    if not m:
        return (99, 99, s.lower())
    jj, mm = int(m.group(1)), int(m.group(2))
    if 1 <= jj <= 31 and 1 <= mm <= 12:
        return (mm, jj, s.lower())
    return (99, 99, s.lower())


def _est_favorable(flux_nom: str, reel: float, prev: float) -> bool:
    if flux_nom in ENCAISSEMENTS:
        return reel >= prev
    if flux_nom in DECAISSEMENTS:
        return abs(reel) <= abs(prev)
    return reel >= prev


@bp.route("/analyse", methods=["POST"])
def analyse():
    import numpy as np
    from sklearn.ensemble import IsolationForest

    body    = request.get_json(force=True) or {}
    section = body.get("section", "").strip()
    annee_s = body.get("annee",   "")
    annee   = int(annee_s) if str(annee_s).isdigit() else None

    if not section or annee is None:
        return jsonify({"error": "Paramètres section et annee requis"}), 400

    # ── Collecte des rows (date, flux, profil, reel, prev, écart) ──────────
    rows = []
    for (sec, flux_name), bucket in CACHE.items():
        if sec != section:
            continue
        if flux_name in FLUX_A_EXCLURE:
            continue

        dates     = bucket.get("dates",        [])
        reel_list = bucket.get("reel",         [])
        prev_vals = bucket.get("prev_vals",    [])
        headers   = bucket.get("prev_headers", [])

        for i, d in enumerate(dates):
            if not (hasattr(d, "year") and d.year == annee):
                continue
            r = reel_list[i] if i < len(reel_list) else None
            if r is None or not isinstance(r, (int, float)) or math.isnan(r):
                continue

            for p_idx, prev_serie in enumerate(prev_vals):
                pv = prev_serie[i] if i < len(prev_serie) else None
                if pv is None or not isinstance(pv, (int, float)) or math.isnan(pv):
                    continue
                nom_profil = (str(headers[p_idx]).strip()
                              if p_idx < len(headers) else f"Profil {p_idx + 1}")
                rows.append({
                    "date":      d.strftime("%d/%m/%Y"),
                    "flux":      flux_name,
                    "profil":    nom_profil,
                    "reel":      round(float(r),  2),
                    "prev":      round(float(pv), 2),
                    "ecart":     round(float(r - pv), 2),
                    "favorable": _est_favorable(flux_name, r, pv),
                })

    if not rows:
        return jsonify({"error": f"Aucune donnée pour {section} en {annee}"}), 422

    # ── Détection d'anomalies par IsolationForest sur l'ensemble ──────────
    ecarts = [row["ecart"] for row in rows]
    arr    = np.array(ecarts).reshape(-1, 1)
    seuil  = 2 * float(np.std(arr))
    contam = min(0.05, max(0.01, float(np.mean(np.abs(arr) > seuil))))

    if len(rows) >= 12 and len(set(ecarts)) > 1:
        model  = IsolationForest(contamination=contam, random_state=42)
        preds  = model.fit_predict(arr)
        for row, pred in zip(rows, preds):
            row["anomalie"] = int(pred == -1)
    else:
        q90 = float(np.quantile(np.abs(arr), 0.90)) if len(rows) > 1 else float(np.max(np.abs(arr)))
        for row in rows:
            row["anomalie"] = int(abs(row["ecart"]) >= q90)

    # ── Construction de la matrice ────────────────────────────────────────
    # Profils et flux qui ont au moins une anomalie
    anom_rows = [r for r in rows if r["anomalie"]]
    if not anom_rows:
        return jsonify({"error": "Aucune anomalie détectée pour cette sélection"}), 422

    profils_set = sorted({r["profil"] for r in anom_rows}, key=_profil_sort_key)
    flux_set    = sorted({r["flux"]   for r in anom_rows})

    # count et mean_abs_ecart par (profil, flux)
    count_map = defaultdict(int)
    mean_map  = defaultdict(list)
    for r in anom_rows:
        key = (r["profil"], r["flux"])
        count_map[key] += 1
        mean_map[key].append(abs(r["ecart"]))

    matrix      = []
    mean_matrix = []
    for profil in profils_set:
        row_counts = []
        row_means  = []
        for flux in flux_set:
            key   = (profil, flux)
            cnt   = count_map.get(key, 0)
            m_abs = round(sum(mean_map[key]) / len(mean_map[key]), 0) if mean_map.get(key) else 0
            row_counts.append(cnt)
            row_means.append(m_abs)
        matrix.append(row_counts)
        mean_matrix.append(row_means)

    max_val = max((max(row) for row in matrix if row), default=1)

    # Détails par cellule (profil, flux) — tous les rows anomalies
    details: dict[str, list] = {}
    for r in anom_rows:
        key = f"{r['profil']}|||{r['flux']}"
        if key not in details:
            details[key] = []
        details[key].append({
            "date":      r["date"],
            "reel":      r["reel"],
            "prev":      r["prev"],
            "ecart":     r["ecart"],
            "favorable": r["favorable"],
        })

    # Trier les détails par |écart| décroissant
    for key in details:
        details[key].sort(key=lambda x: abs(x["ecart"]), reverse=True)

    return jsonify({
        "profils":     profils_set,
        "flux":        flux_set,
        "matrix":      matrix,
        "mean_matrix": mean_matrix,
        "max_val":     max_val,
        "details":     details,
        "n_anomalies": len(anom_rows),
        "n_total":     len(rows),
    })
