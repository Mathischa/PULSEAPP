# -*- coding: utf-8 -*-
"""
api/accueil.py — Endpoint REST pour le tableau de bord PULSE.

GET /api/accueil
    Retourne les KPIs, signaux de priorité et recommandations d'analyse.
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

    annees_list = sorted(annees)

    # ------------------------------------------------------------------
    # Fraîcheur des données
    # ------------------------------------------------------------------
    days_ago: int | None = None
    freshness: str = "unknown"

    if fichiers:
        try:
            last_dt = datetime.strptime(fichiers[0]["modifie"], "%d/%m/%Y %H:%M")
            days_ago = (datetime.now() - last_dt).days
            if days_ago <= 7:
                freshness = "recent"
            elif days_ago <= 45:
                freshness = "normal"
            else:
                freshness = "stale"
        except ValueError:
            pass

    # ------------------------------------------------------------------
    # Années manquantes
    # ------------------------------------------------------------------
    missing_years: list[int] = []
    if len(annees_list) >= 2:
        missing_years = [
            y for y in range(min(annees_list), max(annees_list) + 1)
            if y not in annees_list
        ]

    year_span = (max(annees_list) - min(annees_list) + 1) if len(annees_list) >= 2 else len(annees_list)

    # ------------------------------------------------------------------
    # Signaux de priorité (🔴 critique / 🟡 surveillance / 🟢 normal)
    # ------------------------------------------------------------------
    signals: list[dict] = []

    if len(sections) == 0:
        signals.append({
            "type": "critical",
            "label": "Aucune donnée chargée",
            "detail": "Importez des fichiers pour commencer l'analyse",
        })
    elif len(sections) < 3:
        signals.append({
            "type": "warning",
            "label": f"Couverture partielle — {len(sections)} section(s)",
            "detail": "Certaines filiales peuvent manquer",
        })
    else:
        signals.append({
            "type": "ok",
            "label": f"{len(sections)} sections actives",
            "detail": "Couverture complète du périmètre",
        })

    if freshness == "stale" and days_ago is not None:
        signals.append({
            "type": "warning",
            "label": f"Données vieilles de {days_ago} jours",
            "detail": "Synchronisation recommandée",
        })
    elif freshness == "recent" and days_ago is not None:
        signals.append({
            "type": "ok",
            "label": f"Données fraîches ({days_ago} j)",
            "detail": "Synchronisation récente",
        })
    else:
        signals.append({
            "type": "ok",
            "label": "Synchronisation normale",
            "detail": fichiers[0]["modifie"] if fichiers else "—",
        })

    if missing_years:
        signals.append({
            "type": "warning",
            "label": f"{len(missing_years)} année(s) manquante(s)",
            "detail": f"Données absentes : {', '.join(str(y) for y in missing_years[:4])}",
        })
    elif len(annees_list) >= 3:
        signals.append({
            "type": "ok",
            "label": f"Série complète sur {year_span} ans",
            "detail": f"{min(annees_list)} — {max(annees_list)}",
        })
    else:
        signals.append({
            "type": "ok",
            "label": f"{len(annees_list)} année(s) disponible(s)",
            "detail": ", ".join(str(y) for y in annees_list) if annees_list else "—",
        })

    # ------------------------------------------------------------------
    # Recommandations intelligentes (max 3)
    # ------------------------------------------------------------------
    recommendations: list[dict] = []

    if len(annees_list) >= 2:
        years_label = f"{min(annees_list)} — {max(annees_list)}"
        recommendations.append({
            "id": "tendance",
            "title": "Explorer les tendances",
            "detail": f"Données sur {len(annees_list)} années ({years_label}) — idéal pour identifier les évolutions",
            "url": "/tendance",
            "priority": "normal",
            "icon": "trend",
        })

    if len(sections) >= 3:
        recommendations.append({
            "id": "benchmarking",
            "title": "Comparer les filiales",
            "detail": f"{len(sections)} sections disponibles — analysez les écarts de performance inter-filiales",
            "url": "/benchmarking",
            "priority": "normal",
            "icon": "compare",
        })

    if freshness == "stale":
        recommendations.append({
            "id": "anomalies",
            "title": "Vérifier les anomalies",
            "detail": f"Données vieilles de {days_ago} jours — passez en revue les écarts récents",
            "url": "/heatmap",
            "priority": "warning",
            "icon": "alert",
        })
    else:
        recommendations.append({
            "id": "ecarts",
            "title": "Analyser les écarts",
            "detail": "Identifiez les déviations significatives entre réel et prévision",
            "url": "/ecarts",
            "priority": "info",
            "icon": "alert",
        })

    recommendations = recommendations[:3]

    # ------------------------------------------------------------------
    # Répartition fichiers par année (pour sparkline)
    # ------------------------------------------------------------------
    from collections import Counter
    files_by_year_counter: Counter = Counter()
    for f in fichiers:
        m = _MENSUEL_RE.match(f["nom"])
        if m:
            files_by_year_counter[int(m.group(1))] += 1

    fichiers_par_annee = [
        {"annee": y, "count": files_by_year_counter[y]}
        for y in sorted(files_by_year_counter)
    ]

    # ------------------------------------------------------------------
    # Réponse JSON
    # ------------------------------------------------------------------
    return jsonify(
        {
            "nb_fichiers":      len(fichiers),
            "derniere_maj":     fichiers[0]["modifie"] if fichiers else "N/A",
            "nb_sections":      len(sections),
            "sections":         sorted(sections.keys()),
            "annees":           annees_list,
            "nb_entrees_cache": len(CACHE),
            "fichiers_recents": fichiers[:5],
            # Enrichissements
            "days_ago":         days_ago,
            "freshness":        freshness,
            "year_span":        year_span,
            "missing_years":    missing_years,
            "signals":            signals,
            "recommendations":    recommendations,
            "fichiers_par_annee": fichiers_par_annee,
        }
    )
