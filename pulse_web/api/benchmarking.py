# -*- coding: utf-8 -*-
"""
api/benchmarking.py — Matrice Réel/Prévision + Performance filiales.
Utilise le même pattern que tendance.py : TOKENS + extraire_valeurs.
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request

from pulse_v2.config import FLUX_DECAISSEMENTS, FLUX_ENCAISSEMENTS, FLUX_MIXTES
from pulse_v2.data.cache import TOKENS
from pulse_v2.data.extractor import extraire_valeurs

bp = Blueprint("benchmarking", __name__, url_prefix="/api")


def _favorable(flux_nom: str, reel: float, prev: float) -> bool:
    if flux_nom in FLUX_ENCAISSEMENTS:
        return reel >= prev
    if flux_nom in FLUX_DECAISSEMENTS:
        return abs(reel) <= abs(prev)
    if flux_nom in FLUX_MIXTES:
        return reel >= prev if prev >= 0 else abs(reel) <= abs(prev)
    return reel >= prev


def _get_flux_type(flux_nom: str) -> str:
    if flux_nom in FLUX_ENCAISSEMENTS:
        return "recettes"
    if flux_nom in FLUX_DECAISSEMENTS:
        return "depenses"
    if flux_nom in FLUX_MIXTES:
        return "investissements"
    return "autres"


@bp.route("/benchmarking", methods=["GET"])
def get_benchmarking():
    try:
        annee_s      = request.args.get("annee", "").strip()
        filiale_f    = request.args.get("filiale", "").strip()
        flux_type_f  = request.args.get("flux_type", "").strip()
        annee_filter = int(annee_s) if annee_s else None

        points: list[dict] = []

        for section, flux_list in TOKENS.items():
            if filiale_f and section != filiale_f:
                continue

            for flux_name, _col in flux_list:
                if flux_type_f and _get_flux_type(flux_name) != flux_type_f:
                    continue

                dates, reel_serie, previsions, _labels = extraire_valeurs(
                    section, flux_name, 0, annee=None
                )

                for i, date in enumerate(dates):
                    y = getattr(date, "year", None)
                    if annee_filter and y != annee_filter:
                        continue

                    if i >= len(reel_serie) or reel_serie[i] is None:
                        continue

                    r = float(reel_serie[i])

                    for prev_serie in previsions:
                        if i >= len(prev_serie) or prev_serie[i] is None:
                            continue

                        prev_val = float(prev_serie[i])
                        if r == 0 and prev_val == 0:
                            continue

                        denom    = prev_val if prev_val != 0 else 1
                        ecart_pct = (r - prev_val) / denom * 100

                        points.append({
                            "prevision":  round(prev_val, 2),
                            "reel":       round(r, 2),
                            "ecart_pct":  round(ecart_pct, 1),
                            "filiale":    section,
                            "flux":       flux_name,
                            "favorable":  _favorable(flux_name, r, prev_val),
                            "annee":      y,
                        })

        return jsonify(points)

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500
