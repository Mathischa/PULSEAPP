# -*- coding: utf-8 -*-
"""
api/ecarts.py — Écarts significatifs réel vs prévision.
Utilise le même pattern que tendance.py : TOKENS + extraire_valeurs.
"""
from __future__ import annotations

from flask import Blueprint, jsonify

from pulse_v2.config import FLUX_DECAISSEMENTS, FLUX_ENCAISSEMENTS, FLUX_MIXTES
from pulse_v2.data.cache import TOKENS
from pulse_v2.data.extractor import extraire_valeurs

bp = Blueprint("ecarts", __name__, url_prefix="/api")


def _favorable(flux_nom: str, reel: float, prev: float) -> bool:
    if flux_nom in FLUX_ENCAISSEMENTS:
        return reel >= prev
    if flux_nom in FLUX_DECAISSEMENTS:
        return abs(reel) <= abs(prev)
    if flux_nom in FLUX_MIXTES:
        return reel >= prev if prev >= 0 else abs(reel) <= abs(prev)
    return reel >= prev


@bp.route("/ecarts")
def get_ecarts():
    result: list[dict] = []

    for section, flux_list in TOKENS.items():
        for flux_name, _col in flux_list:
            dates, reel_serie, previsions, noms_profils = extraire_valeurs(
                section, flux_name, 0, annee=None
            )

            for i, date in enumerate(dates):
                if i >= len(reel_serie) or reel_serie[i] is None:
                    continue

                r = float(reel_serie[i])

                for idx, prev_serie in enumerate(previsions):
                    if i >= len(prev_serie) or prev_serie[i] is None:
                        continue

                    p = float(prev_serie[i])
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

                    date_str = (
                        date.strftime("%Y-%m-%d")
                        if hasattr(date, "strftime")
                        else str(date)
                    )
                    annee = int(date.year) if hasattr(date, "year") else None

                    result.append({
                        "date":      date_str,
                        "annee":     annee,
                        "profil":    profil,
                        "filiale":   section,
                        "flux":      flux_name,
                        "reel":      round(r, 2),
                        "prevision": round(p, 2),
                        "ecart_k":   round(r - p, 2),
                        "ecart_pct": round(ecart * 100, 1),
                        "favorable": _favorable(flux_name, r, p),
                    })

    result.sort(key=lambda x: abs(x["ecart_pct"]), reverse=True)
    return jsonify(result)
