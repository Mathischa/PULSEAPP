from spire.pdf.common import *
from spire.pdf import *
import os
import sys
import traceback

# 📂 Dossier contenant les fichiers PDF
pdf_folder = r"C:\Users\0304336A\SNCF\DCF GROUPE (Grp. O365) GrpO365 - Reporting et prévisions\Partage - Invités\Projet PULSE\4. Données historiques\Développement\pdf à convertir"

# 📂 Nouveau dossier pour stocker les fichiers Excel
excel_folder = os.path.join(pdf_folder, "Excels")
os.makedirs(excel_folder, exist_ok=True)

# 📄 Liste de tous les fichiers PDF dans le dossier
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]

if not pdf_files:
    print("⚠️ Aucun fichier PDF trouvé dans le dossier.")
    sys.exit(0)

for pdf_file in pdf_files:
    pdf_path = os.path.abspath(os.path.join(pdf_folder, pdf_file))
    excel_file = os.path.abspath(os.path.join(excel_folder, os.path.splitext(pdf_file)[0] + ".xlsx"))

    print(f"🔄 Conversion en cours : {pdf_file} ...")

    pdf = PdfDocument()
    try:
        # Charger le fichier PDF
        pdf.LoadFromFile(pdf_path)

        # Vérifier s'il contient des pages
        if pdf.Pages.Count == 0:
            print(f"⚠️ Fichier vide ignoré : {pdf_file}")
            continue

        # Définir les options de conversion (améliorées pour la mise en page)
        options = XlsxLineLayoutOptions(
            True,   # Conserver la largeur des colonnes
            True,   # Conserver les polices
            True,   # Conserver les styles
            True,   # Conserver les images
            False   # Ne pas fusionner les cellules (meilleure précision des tableaux)
        )
        pdf.ConvertOptions.SetPdfToXlsxOptions(options)

        # Sauvegarder le fichier Excel
        pdf.SaveToFile(excel_file, FileFormat.XLSX)
        print(f"✅ Converti avec succès : {pdf_file} -> {excel_file}")

    except Exception as e:
        print(f"❌ Erreur lors de la conversion de {pdf_file} : {e}")
        traceback.print_exc()

    finally:
        pdf.Close()

print(f"\n🎯 Conversion terminée. Les fichiers Excel sont stockés dans : {excel_folder}") 