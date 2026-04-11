from openpyxl import Workbook
import os

# === Chemin du fichier à créer ===
fichier_destination = r"C:\Users\0304336A\SNCF\DCF GROUPE (Grp. O365) GrpO365 - Reporting et prévisions\Partage - Invités\Projet PULSE\4. Données historiques\Développement\Résultats\Historique_prev_reel_filiales2.xlsx"

# === Sections avec noms de feuilles ===
sections = [
    {"source": "1_-_SA_SNCF", "prev": "SA_SNCF", "dest": "SNCF_SA"},
    {"source": "3_-_SA_VOYAGEURS", "prev": "SA_VOYAGEURS", "dest": "SA_VOYAGEURS"},
    {"source": "2_-_RESEAU", "prev": "RESEAU", "dest": "RESEAU"},
    {"source": "4_-_G&C", "prev": "G&C", "dest": "G&C"},
    {"source": "5_-_GEODIS", "prev": "GEODIS", "dest": "GEODIS"},
]

# === Crée le dossier s'il n'existe pas ===
dossier_resultats = os.path.dirname(fichier_destination)
os.makedirs(dossier_resultats, exist_ok=True)

# === Supprime le fichier s'il existe déjà ===
if os.path.exists(fichier_destination):
    os.remove(fichier_destination)
    print(f"🗑️ Fichier existant supprimé : {fichier_destination}")

# === Crée un nouveau classeur ===
wb = Workbook()

# Supprime la feuille par défaut
default_sheet = wb.active
wb.remove(default_sheet)

# Crée toutes les feuilles définies dans sections
for s in sections:
    wb.create_sheet(title=s["dest"])

# Sauvegarde le fichier
wb.save(fichier_destination)

print(f"✅ Nouveau fichier Excel créé avec les feuilles : {[s['dest'] for s in sections]}")
