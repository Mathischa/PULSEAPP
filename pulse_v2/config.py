# -*- coding: utf-8 -*-
"""
config.py — Configuration PULSE (chemins, colonnes Excel, etc.).
"""
from __future__ import annotations

import os
import unicodedata
from pathlib import Path

# =============================================================================
# RÉSOLUTION DU CHEMIN DE BASE (DEV_PATH)
# =============================================================================

def _norm(s: str) -> str:
    """Normalise pour comparaisons: minuscules, accents retirés, espaces/underscores compressés."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("_", " ").strip()
    s = " ".join(s.split())
    return s


def _match_tail(path: Path, tail: list[str]) -> bool:
    """Vérifie que 'path' se termine par les éléments de 'tail'."""
    parts = list(path.parts)
    if len(parts) < len(tail):
        return False
    for i in range(1, len(tail) + 1):
        if _norm(parts[-i]) != _norm(tail[-i]):
            return False
    return True


def find_dev_path() -> Path:
    """
    Trouve le chemin DEV_PATH en cherchant la queue standard:
    [...]/Partage - Invités/Projet PULSE/4. Données historiques/Développement
    
    Pour production (Render), utilise DEV_PATH depuis l'env ou depuis le répertoire courant.
    """
    # 1) Si DEV_PATH est défini en variable d'environnement, l'utiliser
    if "DEV_PATH" in os.environ:
        path = Path(os.environ["DEV_PATH"])
        if path.exists():
            return path
    
    # 2) Essai avec le répertoire parent du projet (relatif)
    # Depuis pulse_v2/config.py → remonte à pulse_v2 → remonte à root
    current_dir = Path(__file__).resolve().parent.parent
    if (current_dir / "Données").exists():
        return current_dir
    
    # 3) Sinon, chercher sur le système local (Windows dev)
    USER_ID = Path.home().name
    BASE_DIR = Path(fr"C:\Users\{USER_ID}\SNCF")
    
    REQUIRED_TAIL = [
        "Partage - Invités",
        "Projet PULSE",
        "4. Données historiques",
        "Développement",
    ]

    # Essai direct sur le chemin "canonique"
    default_path = (
        BASE_DIR 
        / "DCF GROUPE (Grp. O365) GrpO365 - Reporting et prévisions" 
        / Path(*REQUIRED_TAIL)
    )
    if default_path.exists():
        return default_path

    # Scan générique : on ne regarde que la fin de chemin (tail)
    candidates = []
    if BASE_DIR.exists():
        for root, dirs, _ in os.walk(BASE_DIR):
            if any(_norm(d) == _norm(REQUIRED_TAIL[-1]) for d in dirs):
                for d in dirs:
                    full = Path(root) / d
                    if _match_tail(full, REQUIRED_TAIL):
                        candidates.append(full)

    if candidates:
        best = sorted(candidates, key=lambda p: len(p.parts), reverse=True)[0]
        return best

    raise FileNotFoundError(
        f"Impossible de localiser DEV_PATH. "
        f"Cherche: .../{'/'.join(REQUIRED_TAIL)}"
    )


# Trouve le chemin de base
DEV_PATH = find_dev_path()

# =============================================================================
# CHEMINS ET FICHIERS DE CONFIGURATION
# =============================================================================

# Répertoire contenant les fichiers historiques mensuels
FICHIER_EXCEL_DIR = str(DEV_PATH / "Données" / "Historique Prévisions Réel Filiales")

# Fichier Excel de configuration des sections (dans le dossier Données)
FICHIER_CONFIG_SECTIONS = str(DEV_PATH / "Données" / "Filiales Analysées.xlsx")

# Dossier de données réelles
BASE_DONNEES_DIR = str(DEV_PATH / "Données" / "Données Réelles")

# Chemin défaut pour les images
IMAGE_PATH = str(DEV_PATH / "Données" / "Images" / "logo_Pulse.png")

# =============================================================================
# CONFIGURATION EXCEL
# =============================================================================

# Nom de la feuille dans FICHIER_CONFIG_SECTIONS
FEUILLE_CONFIG_SECTIONS = "CONFIG_SECTIONS"

# Colonnes dans la feuille de configuration (1-indexed)
COL_DEST = 1     # Colonne A: Destination
COL_SOURCE = 2   # Colonne B: Source
COL_PREV = 3     # Colonne C: Prévisions

# =============================================================================
# CLASSIFICATION DES FLUX TRÉSORERIE
# =============================================================================

# Flux dont une diminution est favorable (charges, décaissements)
FLUX_DECAISSEMENTS = {
    "Décaissements", "Dépenses", "Coûts", 
    "Charges", "Paiements", "Achat", "Frais",
}

# Flux dont une augmentation est favorable (produits, encaissements)
FLUX_ENCAISSEMENTS = {
    "Encaissements", "Revenus", "Ventes",
    "Produits", "Chiffre d'affaires", "Bénéfice",
}

# Flux mixtes (comportement dépend du contexte)
FLUX_MIXTES = {
    "Flux net", "Trésorerie", "Variation",
    "Différentiel", "Solde",
}

# Flux à exclure du catalogue (ex: colonnes de configuration)
FLUX_A_EXCLURE = {
    "Token", "Index", "ID", "Configuration",
    "Total", "Somme", "Agrégé",
}

# =============================================================================
# VALIDATION AU DÉMARRAGE
# =============================================================================

if __name__ != "__main__":
    # Valide que les fichiers/dossiers existent
    import warnings
    
    if not Path(FICHIER_EXCEL_DIR).exists():
        warnings.warn(
            f"⚠️  Dossier 'Historique Prévisions Réel Filiales' introuvable : {FICHIER_EXCEL_DIR}",
            RuntimeWarning
        )
    
    if not Path(FICHIER_CONFIG_SECTIONS).exists():
        warnings.warn(
            f"⚠️  Fichier 'Filiales Analysées.xlsx' introuvable : {FICHIER_CONFIG_SECTIONS}",
            RuntimeWarning
        )
    
    if not Path(BASE_DONNEES_DIR).exists():
        warnings.warn(
            f"⚠️  Dossier 'Données Réelles' introuvable : {BASE_DONNEES_DIR}",
            RuntimeWarning
        )
