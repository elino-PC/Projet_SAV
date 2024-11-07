# -*- coding: utf-8 -*-
"""
Created on Fri Oct 11 12:12:53 2024

@author: elino
"""

import sys
import os
import pandas as pd
import argparse

# Ajouter le chemin vers le répertoire parent pour l'importation des modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from Core.Installation_Definition import SolarInstallationMC
from Report_Generation.Report_Generator import generate_report
from Data_Collection.collectors.MC_Collector import get_site_list


def get_report_type(argument, master_path):
    """
    Récupère le type de rapport (1 mois ou 3 mois) à partir du fichier maître.
    """
    worksheet_name = "meteocontrol"
    try:
        df = pd.read_excel(master_path, sheet_name=worksheet_name)
        recap_values_array = df[df.iloc[:, 0] == argument]  # Utilisez 'argument' pour le nom du site
        if not recap_values_array.empty:
            report_type = recap_values_array.iloc[:, 4].values[0]
        else:
            print(f"Aucune donnée trouvée pour le site {argument}")
            return None

        # Conversion du type de rapport
        if report_type == "1 mois":
            return "1m"
        elif report_type == "3 mois":
            return "3m"
        else:
            print(f"Type de rapport inconnu : {report_type}")
            return None
    except FileNotFoundError:
        print(f"Fichier maître introuvable à {master_path}")
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier maître : {e}")
    return None


def main(argument):
    name_site_list = ['Antana Production', 'EPSILON', 'Epsilon - site 2', 'SOCOTA - PHASE 1', 'Actual Textile', 'Menakao']
    site_ids = get_site_list(name_site_list)

    if argument not in site_ids:
        print("Nom de site incorrect")
        return

    site_id = site_ids[argument]

    master_path = r"G:/.shortcut-targets-by-id/12F2rxDjhgWoKdVhzLQxaPsc1FL1nWr9H/17- Technique/3 - SAV/3 - Rapports de production/3 - Outil de rapport/Projet_SAV/#Master_Report Generator.xlsm"
    template_path = r"G:/.shortcut-targets-by-id/12F2rxDjhgWoKdVhzLQxaPsc1FL1nWr9H/17- Technique/3 - SAV/3 - Rapports de production/3 - Outil de rapport/Projet_SAV/Projet_SAV/Report_Generation/templates/Modèle_Rapport_MC.xlsx"
    result_path = r"G:/.shortcut-targets-by-id/12F2rxDjhgWoKdVhzLQxaPsc1FL1nWr9H/17- Technique/3 - SAV/3 - Rapports de production/3 - Outil de rapport/Projet_SAV/Rapports de production"
    
    print(f"Argument passé à get_report_type: {argument}, master_path: {master_path}")
    report_type = get_report_type(argument, master_path)
    print(f"Type de rapport obtenu : {report_type}")
    
    # Vérifier que report_type est valide
    if report_type not in ['1m', '3m']:
        print(f"Type de rapport invalide : {report_type}. Il doit être '1m' ou '3m'.")
        return

    data_path = "G:/.shortcut-targets-by-id/1FQfz_wNk7M-PeQeUyVAHoy9ay4UY1_62/3 - SAV/3 - Rapports de production/3 - Outil de rapport/Projet_SAV/Projet_SAV/Data_Processing/Data_Folders/Antana Production"  # Remplacez par le chemin réel
    year = 2024  # Remplacez par l'année appropriée si nécessaire

    # Passer les arguments requis à la classe SolarInstallationMC
    test_installation = SolarInstallationMC(argument, report_type, data_path, year)

    generate_report(master_path, template_path, result_path, test_installation)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script pour générer rapport Meteocontrol")
    parser.add_argument("arg", type=str, help="Argument = site solaire")
    args = parser.parse_args()
    
    main(args.arg)
