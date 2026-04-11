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
        "Groupe": "Groupe",
    }
    sections.update(test_sections)
    
    # Flux de test pour chaque filiale
    test_flux = [
        ("Encaissements", 1),
        ("Décaissements", 2),
        ("Trésorerie nette", 3),
        ("Cash Flow opérationnel", 4),
    ]
    
    # Crée les données pour chaque section/flux
    base_date = pd.Timestamp("2024-01-01")
    
    for sec_name in sections.keys():
        TOKENS[sec_name] = test_flux.copy()
        
        for flux_name, col_idx in test_flux:
            # Génère 36 mois de données (2024, 2025, 2026)
            dates = [base_date + pd.DateOffset(months=i) for i in range(36)]
            
            # Données réelles (bruitées)
            import random
            reel = [max(0, 100 + random.uniform(-20, 20) + i*2) for i in range(36)]
            
            # Données prévisionnelles (3 profils)
            prev_headers = [
                "Prévision P1",
                "Prévision P2",
                "Prévision P3",
            ]
            
            prev_vals = [
                [105 + i*1.5 for i in range(36)],  # Profil 1
                [102 + i*2 for i in range(36)],    # Profil 2
                [108 + i*1.8 for i in range(36)],  # Profil 3
            ]
            
            cache_key = (sec_name, flux_name)
            CACHE[cache_key] = {
                "dates": dates,
                "reel": reel,
                "prev_headers": prev_headers,
                "prev_vals": prev_vals,
            }
            
            # Remplir YEAR_INDEX avec les années disponibles
            year_index_key = cache_key
            YEAR_INDEX[year_index_key] = {
                "years": {
                    2024: {"row_idx": 0, "prof_idx": 0, "headers": prev_headers},
                    2025: {"row_idx": 12, "prof_idx": 0, "headers": prev_headers},
                    2026: {"row_idx": 24, "prof_idx": 0, "headers": prev_headers},
                }
            }
    
    # Remplir STRUCT avec la structure des filiales
    for sec_name in sections.keys():
        if sec_name not in STRUCT:
            from collections import OrderedDict
            STRUCT[sec_name] = OrderedDict()
        for flux_name, col_idx in test_flux:
            STRUCT[sec_name][flux_name] = {
                "col_start": col_idx,
                "col_end": col_idx + 1,
                "prev_headers": ["Prévision P1", "Prévision P2", "Prévision P3"],
            }
    
    print("[TEST] Données de test chargées:", len(CACHE), "séries")
    print("[TEST] Sections:", sorted(sections.keys()))
    print("[TEST] Années:", sorted(set(y for info in YEAR_INDEX.values() for y in info["years"].keys())))
