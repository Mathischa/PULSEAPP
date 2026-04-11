# -*- coding: utf-8 -*-
"""
api/catalogue.py — Référentiel sections / flux disponibles.

GET /api/catalogue
    Retourne toutes les sections et leurs flux (pour peupler les dropdowns).
"""
from flask import Blueprint, jsonify

from pulse_v2.config import FLUX_A_EXCLURE
from pulse_v2.data.cache import TOKENS, sections

bp = Blueprint("catalogue", __name__, url_prefix="/api")


@bp.route("/catalogue")
def get_catalogue():
    result: dict[str, list[str]] = {}
    for feuille in sections.values():
        flux_list = [
            name
            for name, _ in TOKENS.get(feuille, [])
            if name not in FLUX_A_EXCLURE
        ]
        if flux_list:
            result[feuille] = flux_list
    return jsonify(result)
