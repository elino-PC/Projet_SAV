
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))) 

from Data_Collection.collectors.PV_Gis import get_irradiance_pv_gis, process_data
from Report_Generation.Report_Generator import get_recap_values
from Data_Collection.collectors.Victron_Collector import SolarInstallationVictron

import pandas as pd
import json

test_installation = SolarInstallationVictron("Centre Anani", id="93302", report_type="1m")
master_path = r"C:\Users\danno\Documents\Projet SAV\Projet_SAV\Projet_SAV\Report_Generation\templates\#Master_Report Generator.xlsx"
recap_values=get_recap_values(master_path, test_installation)

m_pv_gis = pd.Timestamp(recap_values[6]).month
lat = recap_values[18]
lon = recap_values[19]
pvtechchoice = recap_values[20]

# List of parameters to check for missing values and their default values
params = [
    (recap_values[7], 0, "peakpower"),
    (recap_values[21], 0, "loss"),
    (recap_values[22], 0, "mountingplace"),
    (recap_values[23], 0, "angle"),
    (recap_values[24], 0, "azimut"),
    (recap_values[25], 0, "startyear"),
    (recap_values[26], 0, "endyear")
]

# Dictionary to store the parameter values
param_values = {}

# Check for missing values and assign default values if necessary
for value, default, name in params:
    if isinstance(value, str):
        param_values[name]=value
    else:
        param_values[name] = round(value) if not pd.isna(value) else default

# Extract the parameter values from the dictionary
peakpower = param_values["peakpower"]
loss = param_values["loss"]
mountingplace = param_values["mountingplace"]
angle = param_values["angle"]
azimut = param_values["azimut"]
startyear = param_values["startyear"]
endyear = param_values["endyear"]

json_file_path = r"C:\Users\danno\Downloads\Timeseries_-18.916_47.534_SA2_10kWp_crystSi_14_25deg_0deg_2005_2020.json"
with open(json_file_path, 'r') as file:
    data = json.load(file)
daily_mean_df, monthly_data = process_data(data, m_pv_gis)