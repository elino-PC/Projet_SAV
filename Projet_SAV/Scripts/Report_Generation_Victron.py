import sys
import os
import pandas as pd
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))) 

from Core.Installation_Definition import SolarInstallation
from Report_Generation.Report_Generator import generate_report
from Data_Collection.collectors.Victron_Collector import SolarInstallationVictron, get_site_list


import argparse

def get_report_type (site, master_path):
    worksheet_name = "victron energy"
    try:
        df = pd.read_excel(master_path, sheet_name=worksheet_name)
        recap_values_array = df[df.iloc[:,0] == site]
        if not recap_values_array.empty:
            report_type = recap_values_array.iloc[:, 4].values[0]
        else:
            print(f"No data found for site {site}")
        if report_type == "1 mois":
            report_type = "1m"
        elif report_type == "3 mois":
            report_type = "3m"
        return report_type
    except FileNotFoundError:
        print("Master Report Generator was not found.")
        return None

def main(argument):
    site_list, name_list, phone = get_site_list()
    try:
        index = name_list.index(argument)
    except:
        print("Wrong site name")
    site_id = site_list[index]
    master_path = r"G:\.shortcut-targets-by-id\12F2rxDjhgWoKdVhzLQxaPsc1FL1nWr9H\17- Technique\3 - SAV\3 - Rapports de production\3 - Outil de rapport\Projet_SAV\#Master_Report Generator.xlsm"
    template_path = r"G:\.shortcut-targets-by-id\12F2rxDjhgWoKdVhzLQxaPsc1FL1nWr9H\17- Technique\3 - SAV\3 - Rapports de production\3 - Outil de rapport\Projet_SAV\Projet_SAV\Report_Generation\templates\Modèle_Rapport_Victron.xlsx"
    result_path = r"G:\.shortcut-targets-by-id\12F2rxDjhgWoKdVhzLQxaPsc1FL1nWr9H\17- Technique\3 - SAV\3 - Rapports de production\3 - Outil de rapport\Projet_SAV\Rapports de production"
    
    report_type = get_report_type(argument, master_path)
    print(f"Type de rapport obtenu : {report_type}")
    
    # Vérifier que report_type est valide
    if report_type not in ['1m', '3m']:
        print(f"Type de rapport invalide : {report_type}. Il doit être '1m' ou '3m'.")
        return

    test_installation = SolarInstallationVictron(argument, id=site_id, report_type=report_type)

    generate_report(master_path, template_path, result_path,test_installation)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script pour générer rapport Victron")
    parser.add_argument("arg", type=str, help="Argument = site solaire")
    args = parser.parse_args()
    
    main(args.arg)