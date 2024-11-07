import pandas as pd
import numpy as np

class SolarInstallation:
    # Parameters common to all installations
    def __init__(self, name):
        self.name = name
        self.type = None
        self.data = None
        self.start = None
        self.end = None
        self.month_type = None
        self.writer = None

    def load_dates_from_report(self, master_path):
        """
        Charge les dates de début et de fin à partir du fichier Excel maître.
        """
        worksheet_name = self.type
        try:
            df = pd.read_excel(master_path, sheet_name=worksheet_name)
            site_data = df[df.iloc[:, 0] == self.name]

            if not site_data.empty:
                self.start = pd.to_datetime(site_data.iloc[:, 5].values[0], format='%d/%m/%Y')
                self.end = pd.to_datetime(site_data.iloc[:, 6].values[0], format='%d/%m/%Y')
            else:
                print(f"Aucune donnée trouvée pour le site {self.name}")
        except FileNotFoundError:
            print(f"Fichier maître introuvable à {master_path}")
        except Exception as e:
            print(f"Erreur lors du chargement des dates : {e}")

    # Méthodes à implémenter dans les sous-classes
    def get_all_data(self, start, end):
        raise NotImplementedError("Should be implemented by subclasses!")

    def get_data_previous_month(self):
        raise NotImplementedError("Should be implemented by subclasses!")

    def get_data_12_months(self):
        raise NotImplementedError("Should be implemented by subclasses!")

    def get_and_analyze_bv_and_sy(self, start, end):
        raise NotImplementedError("Should be implemented by subclasses!")

    def get_soc(self):
        raise NotImplementedError("Should be implemented by subclasses!")

    def get_alarms(self):
        raise NotImplementedError("Should be implemented by subclasses!")


class SolarInstallationMC(SolarInstallation):
    def __init__(self, name, report_type, data_path, year):
        super().__init__(name)
        self.report_type = report_type
        self.data_path = data_path
        self.year = year
        self.type = "meteocontrol"
        
   
    def load_and_process_day_data(self):
        # Lecture des fichiers max_production_day.csv et max_consumption_day.csv
        df_sun = pd.read_csv(f"{self.data_path}/max_production_day.csv", encoding='utf-16', sep='\t', skiprows=2, on_bad_lines='skip')
        df_conso = pd.read_csv(f"{self.data_path}/max_consumption_day.csv", encoding='utf-16', sep='\t', skiprows=2, on_bad_lines='skip')
        print("colonnes du fichier df_sun :", df_sun.columns)
        # Manipulation des données
        day_data_sun = df_sun.copy()
        day_data_conso = df_conso.copy()
        
        # Formatage de la colonne Date
        day_data_sun['Date'] = pd.to_datetime(day_data_sun['Date'])
        day_data_conso['Date'] = pd.to_datetime(day_data_conso['Date'])
        
        return day_data_sun, day_data_conso

    def get_all_data(self, start=None, end=None):
        df_month = pd.read_csv(f"{self.data_path}/month_report.csv", encoding='utf-16', sep='\t', skiprows=2, on_bad_lines='skip')
        df_previous_month = pd.read_csv(f"{self.data_path}/previous_month_report.csv", encoding='utf-16', sep='\t', skiprows=2, on_bad_lines='skip')
        df_sun = pd.read_csv(f"{self.data_path}/max_production_day.csv", encoding='utf-16', sep='\t', skiprows=2, on_bad_lines='skip')
        df_conso = pd.read_csv(f"{self.data_path}/max_consumption_day.csv", encoding='utf-16', sep='\t', skiprows=2, on_bad_lines='skip')
        
        
        # Afficher les colonnes pour vérifier leur nom exact
        print("Colonnes du fichier df_month:", df_month.columns)

        df_month['Date'] = pd.to_datetime(df_month['Date'], errors='coerce', dayfirst=True)

        # Filtrer les données si `start` et `end` sont fournis
        if start is not None and end is not None:
            start = pd.to_datetime(start)
            end = pd.to_datetime(end)
            df_filtered = df_month[(df_month['Date'] >= start) & (df_month['Date'] <= end)].copy()
        else:
            df_filtered = df_month.copy()

        # Trouver la date et les données de production maximale
        if not df_filtered.empty:
            target_day_sun = df_filtered.loc[df_filtered['Energie PV totale [kWh]'].idxmax()]
            day_data_sun = target_day_sun[['Date', 'Energie PV totale [kWh]']]

        else:
            target_day_sun = None
            day_data_sun = None

        # Trouver la date et les données de consommation maximale
        if not df_filtered.empty:
            target_day_conso = df_filtered.loc[df_filtered['Energie consommées totale [kWh]'].idxmax()]
            day_data_conso = target_day_conso[['Date', 'Energie consommées totale [kWh]']]
        else:
            target_day_conso = None
            day_data_conso = None
        
        # Grouper les données par jour (moyenne, somme, etc., selon le besoin)
        df_grouped_days = df_month.drop(columns=['Date']).groupby(df_month['Date'].dt.date).sum().reset_index()

        # Retourner les variables demandées
        return day_data_sun, day_data_conso, df_grouped_days, target_day_conso, target_day_sun

class SolarInstallationFronius(SolarInstallation):
    def __init__(self, name):
        super().__init__(name)


class SolarInstallationSMA(SolarInstallation):
    def __init__(self, name):
        super().__init__(name)
