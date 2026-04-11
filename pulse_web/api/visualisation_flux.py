# -*- coding: utf-8 -*-
"""
api/visualisation_flux.py — Agrégation des flux réels par année pour superposition.

GET /api/visualisation_flux
    ?section=SA_VOYAGEURS
    &flux=Trafic%20Voyageurs        (ou "Tous les flux")
    &month_start=1
    &month_end=12

Retourne les agrégations weekly / monthly / annual par année.
"""
from __future__ import annotations

import datetime as dt
from collections import defaultdict

from flask import Blueprint, jsonify, request

from pulse_v2.data.cache import CACHE, TOKENS

bp = Blueprint("visualisation_flux", __name__, url_prefix="/api")

MONTHS_LABELS = [
    "Jan", "Fév", "Mar", "Avr", "Mai", "Jun",
    "Jul", "Aoû", "Sep", "Oct", "Nov", "Déc",
]

_TOUS_LES_FLUX = {"tous les flux", "tous", "all", ""}


# =========================================================
# HELPERS
# =========================================================

def _to_float(x) -> float | None:
    if x is None:
        return None
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def _to_date(d) -> dt.date | None:
    """Normalise n'importe quelle valeur date vers dt.date."""
    if d is None:
        return None
    # datetime ou date objet
    if hasattr(d, "date") and callable(d.date):
        try:
            return d.date()
        except Exception:
            return None
    if hasattr(d, "year") and hasattr(d, "month") and hasattr(d, "day"):
        try:
            return dt.date(int(d.year), int(d.month), int(d.day))
        except Exception:
            return None
    # chaîne ISO
    if isinstance(d, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
            try:
                return dt.datetime.strptime(d.strip(), fmt).date()
            except ValueError:
                pass
    return None


# =========================================================
# AGRÉGATION
# =========================================================

def _collect(
    section: str,
    flux_nom: str,
    month_start: int,
    month_end: int,
) -> dict[int, dict]:
    """
    Retourne un dict {année: {weekly, monthly, annual, n}} agrégé
    depuis le CACHE pour les paires (section, flux) concernées.
    """
    flux_lower = flux_nom.strip().lower()
    is_all_flux = flux_lower in _TOUS_LES_FLUX

    if is_all_flux:
        pairs = [
            (section, name)
            for name, _ in TOKENS.get(section, [])
            if (section, name) in CACHE
        ]
    else:
        pairs = [(section, flux_nom)] if (section, flux_nom) in CACHE else []

    if not pairs:
        return {}

    acc: dict[int, dict] = defaultdict(lambda: {
        "weekly":  defaultdict(float),
        "monthly": defaultdict(float),
        "annual":  0.0,
        "n":       0,
    })

    for (sec, flux) in pairs:
        bucket = CACHE.get((sec, flux)) or {}
        dates = bucket.get("dates", [])
        reel  = bucket.get("reel",  [])

        for d_raw, r_raw in zip(dates, reel):
            d = _to_date(d_raw)
            r = _to_float(r_raw)

            if d is None or r is None:
                continue
            if not (month_start <= d.month <= month_end):
                continue

            y = d.year
            w = d.isocalendar()[1]
            m = d.month

            acc[y]["weekly"][w]  += r
            acc[y]["monthly"][m] += r
            acc[y]["annual"]     += r
            acc[y]["n"]          += 1

    return dict(acc)


# =========================================================
# ROUTE
# =========================================================

@bp.route("/visualisation_flux")
def get_visualisation_flux():
    section     = request.args.get("section",     "").strip()
    flux_nom    = request.args.get("flux",        "").strip()
    month_start = request.args.get("month_start", "1")
    month_end   = request.args.get("month_end",   "12")

    if not section:
        return jsonify({"error": "Paramètre 'section' requis"}), 400
    if not flux_nom:
        return jsonify({"error": "Paramètre 'flux' requis"}), 400

    try:
        month_start = max(1, min(12, int(month_start)))
        month_end   = max(1, min(12, int(month_end)))
    except ValueError:
        month_start, month_end = 1, 12

    raw = _collect(section, flux_nom, month_start, month_end)

    if not raw:
        return jsonify({
            "error": f"Aucune donnée pour le flux '{flux_nom}' dans la section '{section}'."
        }), 404

    annees = sorted(raw.keys())

    # Sérialisation JSON (clés str)
    weekly  = {str(y): {str(w): round(v, 2) for w, v in raw[y]["weekly"].items()}  for y in annees}
    monthly = {str(y): {str(m): round(v, 2) for m, v in raw[y]["monthly"].items()} for y in annees}
    annual  = {str(y): round(raw[y]["annual"], 2)                                   for y in annees}

    # KPIs globaux
    total_global  = sum(raw[y]["annual"] for y in annees)
    annee_peak    = max(annees, key=lambda y: raw[y]["annual"])
    annee_trough  = min(annees, key=lambda y: raw[y]["annual"])
    nb_points     = sum(raw[y]["n"] for y in annees)

    return jsonify({
        "section":     section,
        "flux":        flux_nom,
        "month_start": month_start,
        "month_end":   month_end,
        "annees":      annees,
        "weekly":      weekly,
        "monthly":     monthly,
        "annual":      annual,
        "mois_labels": MONTHS_LABELS,
        "kpis": {
            "nb_annees":    len(annees),
            "total_global": round(total_global, 2),
            "nb_points":    nb_points,
            "annee_peak": {
                "annee":  annee_peak,
                "valeur": round(raw[annee_peak]["annual"], 2),
            },
            "annee_trough": {
                "annee":  annee_trough,
                "valeur": round(raw[annee_trough]["annual"], 2),
            },
        },
    })
