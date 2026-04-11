# -*- coding: utf-8 -*-
"""
extractor.py — Extraction de données du cache PULSE.

Fonctions de haut niveau pour extraire et filtrer les données de prévision/réel
depuis le cache chargé en mémoire.
"""
from __future__ import annotations

from typing import Any

from .cache import CACHE, TOKENS


def charger_donnees(section: str, offset: int = 0) -> tuple[str, list[tuple[str, int]]]:
    """
    Charge les données de base pour une section/feuille.
    
    Args:
        section: Nom de la section/feuille
        offset: Décalage de colonne (non utilisé pour l'instant)
    
    Returns:
        (feuille, [(nom_flux, col_idx), ...])
    
    Récupère la liste des flux disponibles dans une feuille depuis TOKENS.
    """
    # Récupère les flux pour cette section depuis TOKENS
    flux_list = TOKENS.get(section, [])
    # Retourne la feuille et la liste des (nom_flux, col_start)
    return section, flux_list


def extraire_valeurs(
    section: str,
    flux_name: str,  # C'est le nom du flux, pas une colonne!
    offset: int = 0,
    annee: int | None = None,
) -> tuple[list, list, list[list], list]:
    """
    Extrait les valeurs réelles et prévisionnelles pour un flux dans une section.
    
    Args:
        section: Nom de la section/feuille
        flux_name: Nom du flux (ex: "Encaissements", "Décaissements")
        offset: Décalage supplémentaire (non utilisé)
        annee: Année de filtrage (optionnel)
    
    Returns:
        (dates, reel_serie, previsions, labels)
        où:
        - dates: liste des timestamps
        - reel_serie: liste des valeurs réelles
        - previsions: liste de listes (une par profil prévisionnel)
        - labels: liste des labels/profils
    """
    # Récupère les données du cache
    cache_key = (section, flux_name)
    if cache_key not in CACHE:
        # Flux non trouvé
        return [], [], [], []
    
    bucket = CACHE[cache_key]
    
    # Récupère les données brutes
    dates = bucket.get("dates", [])
    reel = bucket.get("reel", [])
    prev_headers = bucket.get("prev_headers", [])
    prev_vals = bucket.get("prev_vals", [])
    
    # Filtre par année si spécifiée
    if annee is not None:
        selected_indices = [i for i, d in enumerate(dates) if d.year == annee]
    else:
        selected_indices = list(range(len(dates)))
    
    if not selected_indices:
        return [], [], [], []
    
    # Construit les séries filtrées
    filtered_dates = [dates[i] for i in selected_indices]
    filtered_reel = [reel[i] if i < len(reel) else None for i in selected_indices]
    
    # Construit les séries de prévisions filtrées
    filtered_previsions = []
    for serie in prev_vals:
        filtered_serie = [serie[i] if i < len(serie) else None for i in selected_indices]
        filtered_previsions.append(filtered_serie)
    
    # Labels des profils
    labels = [str(h) for h in prev_headers] if prev_headers else []
    
    return filtered_dates, filtered_reel, filtered_previsions, labels
