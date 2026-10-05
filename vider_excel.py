import os
import sys
import openpyxl

# Configuration de la sortie console pour Windows
if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Charge le fichier original de la préparatrice physique
file_path = "ENTRAINEMENT 080926.xlsx"

if not os.path.exists(file_path):
    downloads_path = os.path.expanduser(r"~\Downloads\ENTRAINEMENT 080926.xlsx")
    if os.path.exists(downloads_path):
        file_path = downloads_path

try:
    print(f"Chargement du fichier : {file_path}...")
    wb = openpyxl.load_workbook(file_path)

    # 1. Vider l'onglet 'données brutes' (à partir de la ligne 3)
    if 'données brutes' in wb.sheetnames:
        ws_brutes = wb['données brutes']
        if ws_brutes.max_row >= 3:
            for row in ws_brutes.iter_rows(min_row=3, max_row=ws_brutes.max_row):
                for cell in row:
                    cell.value = None
        print("Onglet 'données brutes' vidé à partir de la ligne 3.")

    # 2. Vider l'onglet 'données indiv. joueurs' (à partir de la ligne 3)
    if 'données indiv. joueurs' in wb.sheetnames:
        ws_indiv = wb['données indiv. joueurs']
        if ws_indiv.max_row >= 3:
            for row in ws_indiv.iter_rows(min_row=3, max_row=ws_indiv.max_row):
                for cell in row:
                    cell.value = None
        print("Onglet 'données indiv. joueurs' vidé à partir de la ligne 3.")

    # Sauvegarder sous le nouveau nom attendu par l'application
    output_filename = "Modele_Vierge_STAPS.xlsx"
    wb.save(output_filename)
    print(f"Le fichier {output_filename} a été créé avec succès !")

except FileNotFoundError:
    print(f"Erreur : Le fichier {file_path} est introuvable dans ce dossier.")
except Exception as e:
    print(f"Erreur inattendue : {e}")
