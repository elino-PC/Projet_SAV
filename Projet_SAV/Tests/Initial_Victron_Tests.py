import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))) 

from Core.Installation_Definition import SolarInstallation
from Report_Generation.Report_Generator import generate_report
from Data_Collection.collectors.Victron_Collector import SolarInstallationVictron


test_installation = SolarInstallationVictron("Hôtel Sarimanok", id="178826", report_type="1m")
master_path = r"C:\Users\danno\Documents\Projet SAV\Projet_SAV\Projet_SAV\Report_Generation\templates\#Master_Report Generator.xlsx"
template_path = r"C:\Users\danno\Documents\Projet SAV\Projet_SAV\Projet_SAV\Report_Generation\templates\Modèle_Rapport_Victron.xlsx"

generate_report(master_path, template_path, test_installation)