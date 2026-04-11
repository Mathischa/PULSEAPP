# -*- coding: utf-8 -*-
"""
api/visualisation.py — Données brutes pour la page Visualisation graphique.

GET /api/visualisation?section=SA_VOYAGEURS&flux=Trafic%20Voyageurs&annee=2025
    → dates, réel, prévisions (toutes), années disponibles pour ce flux.
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request

from pulse_v2.data.cache import CACHE

bp = Blueprint("visualisation", __name__, url_prefix="/api")

# Flux dont les valeurs sont des encours → convertis en flux (différence successive)
FLUX_A_CONVERTIR = {
    "Emprunts", "Tirages Lignes CT", "Variation de collatéral",
    "Créances CDP", "Placements", "CC financiers",
    "Emprunts / Prêts - Groupe", "Cashpool",
    "Encours de financement", "Endettement Net",
}


def _en_flux(values: list) -> list:
    """Convertit une série d'encours en flux (différences successives)."""
    out = [0.0 if values[0] is not None else None]
    for i in range(1, len(values)):
        v, vp = values[i], values[i - 1]
        out.append(v - vp if v is not None and vp is not None else None)
    return out


@bp.route("/visualisation")
def get_visualisation():
    section   = request.args.get("section", "").strip()
    flux_name = request.args.get("flux",    "").strip()
    annee_s   = request.args.get("annee",   "").strip()
    annee     = int(annee_s) if annee_s.isdigit() else None

    if not section or not flux_name:
        return jsonify({"error": "Paramètres section et flux requis"}), 400

    bucket = CACHE.get((section, flux_name))
    if bucket is None:
        return jsonify({"error": f"Flux '{flux_name}' introuvable pour la section '{section}'"}), 404

    all_dates   = bucket.get("dates",        [])
    all_reel    = bucket.get("reel",         [])
    all_prevs   = bucket.get("prev_vals",    [])
    all_headers = bucket.get("prev_headers", [])

    # Années disponibles
    all_years = sorted({d.year for d in all_dates if hasattr(d, "year")})

    # Filtre par année
    if annee is not None:
        indices = [i for i, d in enumerate(all_dates) if hasattr(d, "year") and d.year == annee]
    else:
        indices = list(range(len(all_dates)))

    dates = [all_dates[i].strftime("%Y-%m-%d") for i in indices]
    reel  = [float(all_reel[i]) if i < len(all_reel) and all_reel[i] is not None else None
             for i in indices]

    # Conversion encours → flux si nécessaire
    is_flux_type = flux_name in FLUX_A_CONVERTIR
    if is_flux_type and reel:
        reel = _en_flux(reel)

    # Profils
    profils = []
    for idx, prev_serie in enumerate(all_prevs):
        nom = (str(all_headers[idx]).strip()
               if idx < len(all_headers) else f"Profil {idx + 1}")
        vals = [float(prev_serie[i]) if i < len(prev_serie) and prev_serie[i] is not None else None
                for i in indices]
        if is_flux_type and vals:
            vals = _en_flux(vals)
        profils.append({"nom": nom, "valeurs": vals})

    return jsonify({
        "dates":        dates,
        "reel":         reel,
        "profils":      profils,
        "annees":       all_years,
        "is_flux_type": is_flux_type,
    })
