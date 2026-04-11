# -*- coding: utf-8 -*-
"""
test_data.py — Données de test pour développement PULSE.

Charge des données de test dans le cache pour simulation.
"""
from __future__ import annotations

import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

from .cache import CACHE, TOKENS, sections, YEAR_INDEX, STRUCT


def load_test_data() -> None:
    """Crée des données de test dans le cache pour la démonstration."""
    
    # Sections/filiales de test
    test_sections = {
        "Filiale A": "Filiale A",
        "Filiale B": "Filiale B",
        "Filiale C": "Filiale C",
    }
    sections.update(test_sections)
    
    # Flux de test pour chaque filiale
    test_flux = [
        ("Encaissements", 1),
        ("Décaissements", 2),
        ("Trésorerie", 3),
    ]
    
    # Crée les données pour chaque section/flux
    base_date = pd.Timestamp("2024-01-01")
    
    for sec_name in sections.keys():
        TOKENS[sec_name] = test_flux.copy()
        
        for flux_name, col_idx in test_flux:
            # Génère 24 mois de données
            dates = [base_date + timedelta(days=30*i) for i in range(24)]
            
            # Données réelles (bruitées)
            import random
            reel = [max(0, 100 + random.uniform(-20, 20) + i*2) for i in range(24)]
            
            # Données prévisionnelles (3 profils)
            prev_headers = [
                "",
                "Prévision 01/24",
                "Prévision 02/24",
                "Prévision 03/24",
            ]
            
            prev_vals = [
                [],  # placeholder pour la première colonne
                [105 + i*1.5 for i in range(24)],  # Profil 1
                [102 + i*2 for i in range(24)],    # Profil 2
                [108 + i*1.8 for i in range(24)],  # Profil 3
            ]
            
            cache_key = (sec_name, flux_name)
            CACHE[cache_key] = {
                "dates": dates,
                "reel": reel,
                "prev_headers": prev_headers[1:],  # Exclut le placeholder
                "prev_vals": prev_vals[1:],         # Exclut le placeholder
            }
    
    print("[TEST] Données de test chargées:", len(CACHE), "séries")
