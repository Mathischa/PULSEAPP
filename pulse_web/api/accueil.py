# -*- coding: utf-8 -*-
"""
api/accueil.py — Endpoint REST pour le tableau de bord PULSE.

GET /api/accueil
    Retourne les KPIs généraux : nombre de fichiers, dernière MAJ,
    sections disponibles, années indexées, entrées en cache.
"""
from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path

from flask import Blueprint, jsonify

from pulse_v2.config import FICHIER_EXCEL_DIR
from pulse_v2.data.cache import CACHE, YEAR_INDEX, sections

bp = Blueprint("accueil", __name__, url_prefix="/api")

_MENSUEL_RE = re.compile(r"^Historique_prev_reel_filiales_(\d{4})_(\d{2})\.xlsx$")


@bp.route("/accueil")
def get_accueil():
    # ------------------------------------------------------------------
    # Fichiers Excel mensuels disponibles
    # ------------------------------------------------------------------
    excel_dir = Path(FICHIER_EXCEL_DIR)
    fichiers: list[dict] = []

    if excel_dir.exists():
        for root, _, files in os.walk(excel_dir):
            for fname in files:
                if _MENSUEL_RE.match(fname):
                    full = Path(root) / fname
                    stat = full.stat()
                    fichiers.append(
                        {
                            "nom": fname,
                            "taille_ko": round(stat.st_size / 1024),
                            "modifie": datetime.fromtimestamp(stat.st_mtime).strftime(
                                "%d/%m/%Y %H:%M"
                            ),
                        }
                    )

    fichiers.sort(key=lambda f: f["nom"], reverse=True)

    # ------------------------------------------------------------------
    # Années indexées (depuis YEAR_INDEX)
    # ------------------------------------------------------------------
    annees: set[int] = set()
    for info in YEAR_INDEX.values():
        annees.update(info.get("years", {}).keys())

    # ------------------------------------------------------------------
    # Réponse JSON
    # ------------------------------------------------------------------
    return jsonify(
        {
            "nb_fichiers": len(fichiers),
            "derniere_maj": fichiers[0]["modifie"] if fichiers else "N/A",
            "nb_sections": len(sections),
            "sections": sorted(sections.keys()),
            "annees": sorted(annees),
            "nb_entrees_cache": len(CACHE),
            "fichiers_recents": fichiers[:5],
        }
    )
