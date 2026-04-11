# -*- coding: utf-8 -*-
"""
api/repartition_flux.py — Répartition des écarts par flux.

GET /api/repartition_flux?section=SA_VOYAGEURS&annee=2025&profil=Prévision+01/01
    → agrégation par flux : nb_ecarts, nb_previsions, pct_ecarts, valeur_ecarts.
    Le filtre profil s'applique uniquement à la valorisation signée.

GET /api/repartition_flux/profils?section=SA_VOYAGEURS&annee=2025
    → liste des profils disponibles pour le filtre.
"""
from __future__ import annotations

import re
from flask import Blueprint, jsonify, request

from pulse_v2.data.cache import CACHE, sections

bp = Blueprint("repartition_flux", __name__, url_prefix="/api")

FLUX_EXCLUS = {
    "Cash flow de financement", "Cash flow net", "Sous total financier",
    "Sous total Investissements nets et ACE", "Free cash Flow",
    "Sous total recettes", "Sous total dépenses", "C/C - Groupe",
}


def _sort_profil_key(name: str):
    m = re.search(r"(\d{1,2})[/-](\d{1,2})", str(name))
    if not m:
        return (99, 99, name.lower())
    jj, mm = int(m.group(1)), int(m.group(2))
    if 1 <= jj <= 31 and 1 <= mm <= 12:
        return (mm, jj, name.lower())
    return (99, 99, name.lower())


@bp.route("/repartition_flux")
def get_repartition_flux():
    section_filter = request.args.get("section", "").strip()
    annee_s = request.args.get("annee", "").strip()
    profil_filter = request.args.get("profil", "").strip()

    annee = int(annee_s) if annee_s.isdigit() else None

    # Sections valides dans la configuration
    sections_valides = set(sections.values())

    nb_ecarts: dict[str, int] = {}
    nb_previsions: dict[str, int] = {}
    valeur_ecarts: dict[str, float] = {}

    for (s, flux_name), bucket in CACHE.items():
        # Filtrage section
        if s not in sections_valides:
            continue
        if section_filter and s != section_filter:
            continue
        if flux_name in FLUX_EXCLUS:
            continue

        dates = bucket.get("dates", [])
        reel_serie = bucket.get("reel", [])
        previsions = bucket.get("prev_vals", [])
        noms_profils = bucket.get("prev_headers", [])

        if flux_name not in nb_ecarts:
            nb_ecarts[flux_name] = 0
            nb_previsions[flux_name] = 0
            valeur_ecarts[flux_name] = 0.0

        for i, d in enumerate(dates):
            # Filtre annee
            if annee is not None:
                try:
                    if d.year != annee:
                        continue
                except Exception:
                    continue

            r = reel_serie[i] if i < len(reel_serie) else None
            r = float(r) if r is not None else 0.0

            for idx, prev_serie in enumerate(previsions):
                if i >= len(prev_serie) or prev_serie[i] is None:
                    continue
                p = float(prev_serie[i])

                # Comptage previsions (pas de filtre profil)
                nb_previsions[flux_name] += 1

                # Calcul ecart
                if p == 0:
                    if r == 0:
                        continue
                    denom = 1.0
                else:
                    denom = p

                ecart = (r - p) / denom
                if abs(ecart) < 0.4:
                    continue

                # Comptage ecarts (pas de filtre profil)
                nb_ecarts[flux_name] += 1

                # Valorisation signée (filtre profil appliqué ici seulement)
                pnom = str(noms_profils[idx]).strip() if idx < len(noms_profils) else ""
                if not profil_filter or pnom == profil_filter:
                    valeur_ecarts[flux_name] += (r - p)

    # Construire et trier la liste (garder uniquement les flux avec ecarts)
    flux_list = []
    for f in nb_ecarts:
        if nb_ecarts[f] == 0:
            continue
        nb_p = nb_previsions[f]
        nb_e = nb_ecarts[f]
        flux_list.append({
            "nom": f,
            "nb_previsions": nb_p,
            "nb_ecarts": nb_e,
            "pct_ecarts": round(nb_e / nb_p * 100, 1) if nb_p > 0 else 0.0,
            "valeur_ecarts": round(valeur_ecarts[f], 2),
        })

    flux_list.sort(key=lambda x: x["nb_ecarts"], reverse=True)

    return jsonify({
        "annee": annee,
        "section": section_filter or None,
        "profil": profil_filter or None,
        "flux": flux_list,
    })


@bp.route("/repartition_flux/profils")
def get_profils():
    """Retourne les profils disponibles pour une section et une année données."""
    section_filter = request.args.get("section", "").strip()
    annee_s = request.args.get("annee", "").strip()
    annee = int(annee_s) if annee_s.isdigit() else None

    sections_valides = set(sections.values())
    profils_set: set[str] = set()

    for (s, flux_name), bucket in CACHE.items():
        if s not in sections_valides:
            continue
        if section_filter and s != section_filter:
            continue
        if flux_name in FLUX_EXCLUS:
            continue

        dates = bucket.get("dates", [])
        noms_profils = bucket.get("prev_headers", [])
        prev_vals = bucket.get("prev_vals", [])

        for idx, prev_serie in enumerate(prev_vals):
            has_data = False
            for i, d in enumerate(dates):
                if annee is not None:
                    try:
                        if d.year != annee:
                            continue
                    except Exception:
                        continue
                if i < len(prev_serie) and prev_serie[i] is not None:
                    has_data = True
                    break
            if has_data and idx < len(noms_profils):
                p = str(noms_profils[idx]).strip()
                if p:
                    profils_set.add(p)

    profils = sorted(profils_set, key=_sort_profil_key)
    return jsonify({"profils": profils})
