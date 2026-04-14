# -*- coding: utf-8 -*-
"""
api/benchmarking.py — Endpoint REST pour Matrice Réel/Prévision et Performance Filiales.

GET /api/benchmarking
    Retourne les données pour scatter plot (Réel vs Prévision)
    et radar chart (Performance par filiale).

Query params:
    annee: int (optionnel)
    filiale: str (optionnel)
    flux_type: str in ['recettes', 'depenses', 'investissements'] (optionnel)

Champs retournés:
    prevision: float valeur prévisionnelle (k€)
    reel: float valeur réelle (k€)
    ecart_pct: float écart en %
    filiale: str nom filiale
    favorable: bool True si écart favorable
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request

from pulse_v2.config import FLUX_DECAISSEMENTS, FLUX_ENCAISSEMENTS, FLUX_MIXTES
from pulse_v2.data.cache import CACHE, sections
from pulse_v2.data.extractor import charger_donnees, extraire_valeurs

bp = Blueprint("benchmarking", __name__, url_prefix="/api")


def _favorable(flux_nom: str, reel: float, prev: float) -> bool:
    """Détermine si un écart est favorable selon la nature du flux."""
    if flux_nom in FLUX_ENCAISSEMENTS:
        return reel >= prev
    if flux_nom in FLUX_DECAISSEMENTS:
        return abs(reel) <= abs(prev)
    if flux_nom in FLUX_MIXTES:
        return reel >= prev if prev >= 0 else abs(reel) <= abs(prev)
    return reel >= prev


def _get_flux_type(flux_nom: str) -> str:
    """Détermine le type de flux."""
    if flux_nom in FLUX_ENCAISSEMENTS:
        return "recettes"
    if flux_nom in FLUX_DECAISSEMENTS:
        return "depenses"
    if flux_nom in FLUX_MIXTES:
        return "investissements"
    return "autres"


def _year_of(d):
    """Extrait l'année d'une date."""
    if d is None:
        return None
    if hasattr(d, "year"):
        try:
            return int(d.year)
        except Exception:
            pass
    if isinstance(d, str):
        import re
        m = re.search(r"(20\d{2}|19\d{2})", d)
        if m:
            return int(m.group(1))
    return None


@bp.route("/benchmarking", methods=["GET"])
def get_benchmarking():
    """Retourne données pour scatter + radar."""
    try:
        annee_filter = request.args.get("annee", "", type=str)
        filiale_filter = request.args.get("filiale", "", type=str)
        flux_type_filter = request.args.get("flux_type", "", type=str)

        annee_filter = int(annee_filter) if annee_filter else None
        
        points = []
        feuilles_all = list(sections.values())

        for feuille in feuilles_all:
            if filiale_filter and feuille != filiale_filter:
                continue

            try:
                ws, noms_colonnes = charger_donnees(feuille)
            except TypeError:
                continue

            for nom_flux, col_start in noms_colonnes:
                # Filtrer par type de flux
                if flux_type_filter:
                    if _get_flux_type(nom_flux) != flux_type_filter:
                        continue

                try:
                    dates, reel, previsions, noms_profils = extraire_valeurs(
                        ws, col_start, annee=None
                    )
                except TypeError:
                    try:
                        dates, reel, previsions, noms_profils = extraire_valeurs(ws, col_start)
                    except Exception:
                        continue

                for i, date in enumerate(dates):
                    y = _year_of(date)
                    if annee_filter and y != annee_filter:
                        continue

                    if i >= len(reel) or reel[i] is None:
                        continue

                    r = reel[i]

                    for idx, prev_list in enumerate(previsions):
                        if i >= len(prev_list) or prev_list[i] is None:
                            continue

                        prev_val = prev_list[i]

                        if r == 0 and prev_val == 0:
                            continue
                        if prev_val == 0:
                            prev_val = 1

                        ecart_pct = ((r - prev_val) / prev_val) * 100

                        points.append({
                            "prevision": round(prev_val, 2),
                            "reel": round(r, 2),
                            "ecart_pct": round(ecart_pct, 1),
                            "filiale": feuille,
                            "flux": nom_flux,
                            "favorable": _favorable(nom_flux, r, prev_val),
                        })

        return jsonify(points)

    except Exception as e:
        print(f"[ERROR] benchmarking API: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
