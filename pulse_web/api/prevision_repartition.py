# -*- coding: utf-8 -*-
"""
api/prevision_repartition.py — Répartition des écarts par profil de prévision.
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request
from collections import defaultdict
import re

from pulse_v2.data.cache import CACHE, TOKENS, sections
from pulse_v2.data.extractor import charger_donnees, extraire_valeurs
from pulse_v2.config import FLUX_A_EXCLURE

bp = Blueprint("prevision_repartition", __name__, url_prefix="/api")


def _to_number(x):
    """Convertit un nombre potentiellement mal formaté en float"""
    if x is None:
        return None
    if isinstance(x, str):
        s = x.strip().replace("\xa0", " ").replace(" ", "")
        if s in {"", "-", "—", "NA", "N/A"}:
            return None
        s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return None
    try:
        return float(x)
    except Exception:
        return None


def _is_filled(x):
    """Vérifie si une valeur est remplie"""
    return _to_number(x) is not None


def _parse_jj_mm(nom: str):
    """Extrait jour/mois d'un nom de profil (ex: "Prévision 15/02")"""
    m = re.search(r"(\d{1,2})[/-](\d{1,2})", str(nom))
    if not m:
        return 99, 99
    jj = int(m.group(1))
    mm = int(m.group(2))
    if 1 <= jj <= 31 and 1 <= mm <= 12:
        return mm, jj
    return 99, 99


@bp.route("/prevision_repartition")
def get_prevision_repartition():
    """
    Retourne les données de répartition des écarts par profil.
    
    Query params:
        filiale: "Toutes filiales" ou nom d'une section
        annee: année (int) ou "Toutes années"
        flux: "Tous les flux" ou nom d'un flux
    
    Returns JSON avec:
        - filiales: liste des sections
        - annees: liste des années disponibles
        - flux_list: liste des flux disponibles
        - data: { profils: [...], taux: [...], valorisation: [...], table: [...] }
    """
    
    # Params
    filiale_param = request.args.get("filiale", "Toutes filiales").strip()
    annee_str = request.args.get("annee", "Toutes années").strip()
    flux_param = request.args.get("flux", "Tous les flux").strip()
    
    annee = None if annee_str in {"", "Toutes années"} else int(annee_str) if annee_str.isdigit() else None
    
    # Filiales à analyser
    filiales_calc = list(sections.values()) if filiale_param == "Toutes filiales" else [filiale_param]
    
    # Dicts de comptage
    compteur_ecarts = defaultdict(int)
    valorisation_ecarts = defaultdict(float)
    compteur_total = defaultdict(int)
    
    # ========== PASS 1: écarts + valorisation ==========
    for feuille in filiales_calc:
        try:
            _, noms_colonnes = charger_donnees(feuille, 0)
        except Exception:
            continue
        
        for nom_flux, col_idx in noms_colonnes:
            # Filtre flux
            if flux_param != "Tous les flux" and nom_flux != flux_param:
                continue
            if nom_flux in FLUX_A_EXCLURE:
                continue
            
            try:
                dates, reel, previsions, noms_profils = extraire_valeurs(feuille, nom_flux, 0, annee=None)
            except Exception:
                continue
            
            for p_idx, nom_profil in enumerate(noms_profils):
                if p_idx >= len(previsions):
                    continue
                prev_list = previsions[p_idx]
                
                for i, d in enumerate(dates):
                    # Filtre année
                    if annee is not None:
                        try:
                            y = d.year if hasattr(d, "year") else None
                        except Exception:
                            y = None
                        if y is not None and y != annee:
                            continue
                    
                    pv_raw = prev_list[i] if i < len(prev_list) else None
                    prev_val = _to_number(pv_raw)
                    r_val = _to_number(reel[i] if i < len(reel) else None)
                    
                    if prev_val is None:
                        continue
                    if r_val is None:
                        r_val = 0.0
                    
                    if prev_val == 0:
                        if r_val == 0:
                            continue
                        prev_val = 1.0
                    
                    ecart = (r_val - prev_val) / prev_val
                    if abs(ecart) >= 0.4:
                        compteur_ecarts[nom_profil] += 1
                        valorisation_ecarts[nom_profil] += abs((prev_val or 0.0) - (r_val or 0.0))
    
    # ========== PASS 2: nombre de prévisions non vides ==========
    for feuille in filiales_calc:
        try:
            _, noms_colonnes = charger_donnees(feuille, 0)
        except Exception:
            continue
        
        for nom_flux, col_idx in noms_colonnes:
            if flux_param != "Tous les flux" and nom_flux != flux_param:
                continue
            if nom_flux in FLUX_A_EXCLURE:
                continue
            
            try:
                dates, reel, previsions, noms_profils = extraire_valeurs(feuille, nom_flux, 0, annee=None)
            except Exception:
                continue
            
            for p_idx, nom_profil in enumerate(noms_profils):
                if p_idx >= len(previsions):
                    continue
                prev_list = previsions[p_idx]
                
                nb_prev_non_vides = 0
                for i, d in enumerate(dates):
                    if annee is not None:
                        try:
                            y = d.year if hasattr(d, "year") else None
                        except Exception:
                            y = None
                        if y is not None and y != annee:
                            continue
                    if i < len(prev_list) and _is_filled(prev_list[i]):
                        nb_prev_non_vides += 1
                
                compteur_total[nom_profil] += nb_prev_non_vides
    
    # ========== TRI + FILTRAGE ==========
    rows = []
    all_profils = set(compteur_total.keys()) | set(compteur_ecarts.keys())
    
    for nom in all_profils:
        total = compteur_total.get(nom, 0)
        ecarts = compteur_ecarts.get(nom, 0)
        
        if annee is not None and total == 0 and ecarts == 0:
            continue
        
        taux = (ecarts / total * 100) if total > 0 else 0.0
        valo = valorisation_ecarts.get(nom, 0.0)
        mm, jj = _parse_jj_mm(nom)
        rows.append({
            "mm": mm,
            "jj": jj,
            "profil": nom,
            "taux": taux,
            "valorisation": valo,
            "total": total,
            "ecarts": ecarts
        })
    
    rows.sort(key=lambda r: (r["mm"], r["jj"], str(r["profil"]).casefold()))
    
    # ========== BUILD RESPONSE ==========
    table_data = []
    total_prev_global = 0
    total_ecarts_global = 0
    total_valo_global = 0.0
    
    for row in rows:
        table_data.append([
            row["profil"],
            row["total"],
            row["ecarts"],
            round(row["taux"], 2),
            round(row["valorisation"], 0)
        ])
        total_prev_global += row["total"]
        total_ecarts_global += row["ecarts"]
        total_valo_global += row["valorisation"]
    
    # Ligne de total
    if rows:
        taux_global = (total_ecarts_global / total_prev_global * 100) if total_prev_global > 0 else 0.0
        table_data.append([
            "TOTAL",
            total_prev_global,
            total_ecarts_global,
            round(taux_global, 2),
            round(total_valo_global, 0)
        ])
    
    return jsonify({
        "filiale": filiale_param,
        "annee": annee_str,
        "flux": flux_param,
        "profils": [r["profil"] for r in rows],
        "taux": [r["taux"] for r in rows],
        "valorisation": [r["valorisation"] for r in rows],
        "table": table_data,
        "empty": len(rows) == 0
    })


@bp.route("/prevision_repartition/config")
def get_prevision_repartition_config():
    """Retourne la configuration: filiales, années, flux disponibles"""
    
    filiales = ["Toutes filiales"] + list(sections.values())
    
    # Collecte les années
    annees_set = set()
    for cache_key in CACHE.keys():
        try:
            bucket = CACHE[cache_key]
            dates = bucket.get("dates", [])
            for d in dates:
                try:
                    y = d.year if hasattr(d, "year") else None
                    if y:
                        annees_set.add(y)
                except Exception:
                    pass
        except Exception:
            pass
    
    annees = ["Toutes années"] + sorted(annees_set, reverse=True)
    
    # Collecte les flux
    flux_set = set()
    for cache_key in CACHE.keys():
        section, flux_name = cache_key
        if flux_name not in FLUX_A_EXCLURE:
            flux_set.add(flux_name)
    
    flux_list = ["Tous les flux"] + sorted(flux_set)
    
    return jsonify({
        "filiales": filiales,
        "annees": annees,
        "flux": flux_list
    })
