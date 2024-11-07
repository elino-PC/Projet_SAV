# -*- coding: utf-8 -*-
"""
Created on Wed Oct 23 10:03:55 2024

@author: elino
"""

    def get_all_data(self, start, end):
        # Conversion des timestamps UNIX en datetime si nécessaire
        if isinstance(start, (int, float)):  # Vérifie si start est un timestamp
            start = pd.to_datetime(start, unit='s')
        if isinstance(end, (int, float)):
            end = pd.to_datetime(end, unit='s')
        
        # Charger les données des fichiers CSV
        production_csv = self.data_path + '/max_production_day.csv'
        consommation_csv = self.data_path + '/max_consumption_day.csv'

        # Charger les fichiers CSV avec pandas
        df_sun = pd.read_csv(production_csv, encoding='utf-16', sep='\t', skiprows=2, header=0, on_bad_lines='skip')
        df_conso = pd.read_csv(consommation_csv, encoding='utf-16', sep='\t', skiprows=2, header=0, on_bad_lines='skip')

        # Supprimer les espaces autour des noms de colonnes
        df_sun.columns = df_sun.columns.str.strip()
        df_conso.columns = df_conso.columns.str.strip()       

        # Convertir les colonnes "Date" en datetime, si elles ne le sont pas déjà
        df_sun['Date'] = pd.to_datetime(df_sun['Date'], format='%H:%M', errors='coerce')
        df_conso['Date'] = pd.to_datetime(df_conso['Date'], format='%H:%M', errors='coerce')

        # Filtrer les colonnes numériques pour éviter la somme sur des colonnes datetime
        numeric_cols_sun = df_sun.select_dtypes(include=[np.number]).columns
        numeric_cols_conso = df_conso.select_dtypes(include=[np.number]).columns

        # Vérifier que les colonnes numériques existent
        if not numeric_cols_sun.empty:
            df_grouped_sun = df_sun.groupby(df_sun['Date'].dt.date)[numeric_cols_sun].sum()
        else:
            df_grouped_sun = pd.DataFrame()  # Créer un DataFrame vide si aucune colonne numérique

        if not numeric_cols_conso.empty:
            df_grouped_conso = df_conso.groupby(df_conso['Date'].dt.date)[numeric_cols_conso].sum()
        else:
            df_grouped_conso = pd.DataFrame()  # Créer un DataFrame vide si aucune colonne numérique

        # Identifier le jour de production maximale
        target_day_sun = df_grouped_sun['Energie PV totale [kWh]'].idxmax() if not df_grouped_sun.empty else None
        # Identifier le jour de consommation maximale
        target_day_conso = df_grouped_conso['Energie consommées totale [kWh]'].idxmax() if not df_grouped_conso.empty else None

        # Retourner les valeurs
        return df_sun, df_conso, df_grouped_sun, target_day_conso, target_day_sun
