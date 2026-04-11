# -*- coding: utf-8 -*-
"""
api/heatmap_ecarts.py — Heatmap des écarts significatifs (>=40%) par profil × mois.

POST /api/heatmap_ecarts/analyse  { section, annee, flux }
    → matrice (profils × mois), compte des écarts
"""
from __future__ import annotations
import re, math
from collections import defaultdict
from flask import Blueprint, jsonify, request
from pulse_v2.data.cache import CACHE, TOKENS
import pandas as pd

bp = Blueprint("heatmap_ecarts", __name__, url_prefix="/api/heatmap_ecarts")

FLUX_A_EXCLURE = {
    "Cash flow de financement", "Cash flow net", "Sous total financier",
    "Sous total Investissements nets et ACE", "Free cash Flow",
    "Sous total recettes", "Sous total dépenses", "C/C - Groupe",
    "Emprunts", "Tirages Lignes CT", "Variation de collatéral",
    "Créances CDP", "Placements", "CC financiers",
    "Emprunts / Prêts - Groupe", "Cashpool", "Encours de financement",
    "Endettement Net",
}


def _profil_sort_key(name: str) -> tuple:
    s = str(name).strip()
    m = re.search(r"(\d{1,2})[/-](\d{1,2})", s)
    if not m:
        return (99, 99, s.lower())
    jj, mm = int(m.group(1)), int(m.group(2))
    if 1 <= jj <= 31 and 1 <= mm <= 12:
        return (mm, jj, s.lower())
    return (99, 99, s.lower())


def _to_number(x) -> float:
    if x is None:
        return None
    if isinstance(x, str):
        s = x.strip().replace("\xa0", " ").replace(" ", "")
        if s in {"", "-", "—", "NA", "N/A"}:
            return None
        s = s.replace(",", ".")
        try:
            return float(s)
        except:
            return None
    try:
        return float(x)
    except:
        return None


def _year_of(d) -> int:
    if d is None:
        return None
    if hasattr(d, "year"):
        try:
            return int(d.year)
        except:
            return None
    if isinstance(d, (int, float)):
        y = int(d)
        return y if 1900 <= y <= 2100 else None
    if isinstance(d, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%Y", "%d/%m/%y", "%Y/%m/%d"):
            try:
                return int(pd.to_datetime(d, format=fmt).year)
            except:
                pass
        m = re.search(r"(20\d{2}|19\d{2})", d)
        if m:
            return int(m.group(1))
    return None


@bp.route("/analyse", methods=["POST"])
def analyse():
    body = request.get_json(force=True) or {}
    section = body.get("section", "").strip()
    annee_s = body.get("annee", "")
    flux_filtre = body.get("flux", "Tous flux").strip()
    
    annee = int(annee_s) if str(annee_s).isdigit() else None

    if not section or annee is None:
        return jsonify({"error": "Paramètres section et annee requis"}), 400

    # ── Collecte des rows (date, profil, mois, reel, prev) ────────
    rows = []
    for (sec, flux_name), bucket in CACHE.items():
        if sec != section:
            continue
        if flux_name in FLUX_A_EXCLURE:
            continue
        if flux_filtre != "Tous flux" and flux_name != flux_filtre:
            continue

        dates = bucket.get("dates", [])
        reel_list = bucket.get("reel", [])
        prev_vals = bucket.get("prev_vals", [])
        headers = bucket.get("prev_headers", [])

        for i, d in enumerate(dates):
            if not (hasattr(d, "year") and d.year == annee):
                continue
            r = reel_list[i] if i < len(reel_list) else None
            if r is None or not isinstance(r, (int, float)) or math.isnan(r):
                continue

            try:
                mois = pd.to_datetime(d).strftime("%Y-%m")
            except:
                continue

            for p_idx, prev_serie in enumerate(prev_vals):
                pv = prev_serie[i] if i < len(prev_serie) else None
                if pv is None or not isinstance(pv, (int, float)) or math.isnan(pv):
                    continue
                
                try:
                    ecart_pct = abs((pv - r) / r) if r != 0 else 0
                except:
                    continue

                # Seuil: écart >= 40%
                if ecart_pct >= 0.4:
                    nom_profil = (str(headers[p_idx]).strip()
                                  if p_idx < len(headers) else f"Profil {p_idx + 1}")
                    rows.append({
                        "mois": mois,
                        "profil": nom_profil,
                        "reel": round(float(r), 2),
                        "prev": round(float(pv), 2),
                        "ecart_pct": round(ecart_pct * 100, 1),
                    })

    if not rows:
        return jsonify({"error": f"Aucune donnée pour {section} en {annee}"}), 422

    # ── Pivot table (Profils × Mois) ────────────────────────────
    data_dict = defaultdict(lambda: defaultdict(int))
    for row in rows:
        key = (row["profil"], row["mois"])
        data_dict[row["profil"]][row["mois"]] += 1

    if not data_dict:
        return jsonify({"error": "Aucun écart significatif trouvé"}), 422

    heatmap_df = pd.DataFrame(data_dict).T.fillna(0).astype(int)
    heatmap_df = heatmap_df.reindex(sorted(heatmap_df.columns), axis=1)
    heatmap_df = heatmap_df.loc[sorted(heatmap_df.index, key=_profil_sort_key)]

    matrix = heatmap_df.values.tolist()
    profils = heatmap_df.index.tolist()
    mois_list = heatmap_df.columns.tolist()
    max_val = int(heatmap_df.max().max()) if not heatmap_df.empty else 1

    # ── Détails par cellule (profil, mois) ────────────────────────
    details = {}
    for row in rows:
        key = f"{row['profil']}|||{row['mois']}"
        if key not in details:
            details[key] = []
        details[key].append({
            "reel": row["reel"],
            "prev": row["prev"],
            "ecart_pct": row["ecart_pct"],
        })

    return jsonify({
        "profils": profils,
        "mois": mois_list,
        "matrix": matrix,
        "max_val": max_val,
        "details": details,
        "n_ecarts": len(rows),
    })
