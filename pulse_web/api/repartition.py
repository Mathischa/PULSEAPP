# -*- coding: utf-8 -*-
"""
api/repartition.py — Distribution des écarts par filiale.

GET /api/repartition?annee=2023
    → agrégation du nb d'écarts ≥40%, favorable/défavorable, par filiale.
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request

from pulse_v2.config import FLUX_DECAISSEMENTS, FLUX_ENCAISSEMENTS, FLUX_MIXTES
from pulse_v2.data.cache import sections
from pulse_v2.data.extractor import charger_donnees, extraire_valeurs

bp = Blueprint("repartition", __name__, url_prefix="/api")


def _favorable(flux_nom: str, reel: float, prev: float) -> bool:
    if flux_nom in FLUX_ENCAISSEMENTS:
        return reel >= prev
    if flux_nom in FLUX_DECAISSEMENTS:
        return abs(reel) <= abs(prev)
    if flux_nom in FLUX_MIXTES:
        return reel >= prev if prev >= 0 else abs(reel) <= abs(prev)
    return reel >= prev


@bp.route("/repartition")
def get_repartition():
    annee_s = request.args.get("annee", "").strip()
    annee   = int(annee_s) if annee_s.isdigit() else None

    par_filiale: list[dict] = []
    total = 0

    for feuille in sections.values():
        _, noms_colonnes = charger_donnees(feuille, 0)

        nb_fav = nb_def = 0
        ecarts_pct: list[float] = []

        for nom_flux, col_start in noms_colonnes:
            dates, reel_serie, previsions, _ = extraire_valeurs(
                feuille, nom_flux, 0, annee=annee
            )

            for i, date in enumerate(dates):
                if i >= len(reel_serie) or reel_serie[i] is None:
                    continue
                r = reel_serie[i]

                for prev_serie in previsions:
                    if i >= len(prev_serie) or prev_serie[i] is None:
                        continue
                    p = prev_serie[i]
                    if r == 0 and p == 0:
                        continue
                    denom = p if p != 0 else 1
                    ecart = (r - p) / denom

                    if abs(ecart) < 0.4:
                        continue

                    ecarts_pct.append(abs(ecart) * 100)
                    if _favorable(nom_flux, r, p):
                        nb_fav += 1
                    else:
                        nb_def += 1

        nb = nb_fav + nb_def
        total += nb

        if nb == 0:
            continue

        par_filiale.append({
            "filiale": feuille,
            "nb_ecarts": nb,
            "nb_favorables": nb_fav,
            "nb_defavorables": nb_def,
            "pct_favorables": round(nb_fav / nb * 100, 1),
            "ecart_moy_pct": round(sum(ecarts_pct) / len(ecarts_pct), 1),
            "ecart_max_pct": round(max(ecarts_pct), 1),
        })

    par_filiale.sort(key=lambda x: x["nb_ecarts"], reverse=True)

    return jsonify({
        "annee": annee,
        "total_ecarts": total,
        "nb_filiales": len(par_filiale),
        "par_filiale": par_filiale,
    })
