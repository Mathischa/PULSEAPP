# -*- coding: utf-8 -*-
"""
api/ecarts.py — Endpoint REST pour l'analyse des écarts PULSE.

GET /api/ecarts
    Retourne la liste des écarts ≥ 40 % (réel vs prévision),
    triés par valeur absolue d'écart décroissante.

Champs retournés par ligne :
    date        str   "YYYY-MM-DD"
    annee       int
    profil      str   label du profil de prévision
    filiale     str   nom de la section/filiale
    flux        str   nom du flux trésorerie
    reel        float valeur réelle (k€)
    prevision   float valeur prévisionnelle (k€)
    ecart_k     float écart absolu (k€)
    ecart_pct   float écart en % (arrondi 1 décimale)
    favorable   bool  True si l'écart va dans le bon sens métier
"""
from __future__ import annotations

from flask import Blueprint, jsonify

from pulse_v2.config import FLUX_DECAISSEMENTS, FLUX_ENCAISSEMENTS, FLUX_MIXTES
from pulse_v2.data.cache import sections
from pulse_v2.data.extractor import charger_donnees, extraire_valeurs

bp = Blueprint("ecarts", __name__, url_prefix="/api")


# ---------------------------------------------------------------------------
# Helper : favorabilité métier
# ---------------------------------------------------------------------------

def _favorable(flux_nom: str, reel: float, prev: float) -> bool:
    """Détermine si un écart est favorable selon la nature du flux."""
    if flux_nom in FLUX_ENCAISSEMENTS:
        return reel >= prev
    if flux_nom in FLUX_DECAISSEMENTS:
        return abs(reel) <= abs(prev)
    if flux_nom in FLUX_MIXTES:
        return reel >= prev if prev >= 0 else abs(reel) <= abs(prev)
    return reel >= prev


# ---------------------------------------------------------------------------
# Route
# ---------------------------------------------------------------------------

@bp.route("/ecarts")
def get_ecarts():
    """Optimized: iterate directly on CACHE instead of charger_donnees/extraire_valeurs"""
    result: list[dict] = []

    # Direct CACHE iteration - avoid redundant function calls
    for (section, flux_name), bucket in CACHE.items():
        dates = bucket.get("dates", [])
        reel_serie = bucket.get("reel", [])
        previsions = bucket.get("prev_vals", [])
        noms_profils = bucket.get("prev_headers", [])

        for i, date in enumerate(dates):
            if i >= len(reel_serie) or reel_serie[i] is None:
                continue

            r = reel_serie[i]

            for idx, prev_serie in enumerate(previsions):
                if i >= len(prev_serie) or prev_serie[i] is None:
                    continue

                p = prev_serie[i]

                if r == 0 and p == 0:
                    continue

                denom = p if p != 0 else 1
                ecart = (r - p) / denom

                if abs(ecart) < 0.4:
                    continue

                profil = (
                    noms_profils[idx]
                    if idx < len(noms_profils)
                    else f"Profil {idx + 1}"
                )

                result.append(
                    {
                        "date": date.strftime("%Y-%m-%d"),
                        "annee": int(date.year),
                        "profil": profil,
                        "filiale": section,
                        "flux": flux_name,
                        "reel": round(float(r), 2),
                        "prevision": round(float(p), 2),
                        "ecart_k": round(float(r - p), 2),
                        "ecart_pct": round(float(ecart * 100), 1),
                        "favorable": _favorable(flux_name, r, p),
                    }
                )

    result.sort(key=lambda x: abs(x["ecart_pct"]), reverse=True)
    return jsonify(result)
