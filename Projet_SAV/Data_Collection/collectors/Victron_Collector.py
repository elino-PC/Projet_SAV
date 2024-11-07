
import requests
import http.client
import os
import json
import datetime
from time import * 
from datetime import *
import pandas as pd
pd.options.mode.chained_assignment = None
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
from openpyxl import load_workbook
from openpyxl.chart import  BarChart, PieChart, LineChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font 
from openpyxl.drawing.image import Image  
import numpy as np
import pytz
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import ssl
import io
from sklearn.linear_model import LinearRegression

# Import Solar Installation base classe
from Core.Installation_Definition import SolarInstallation



class SolarInstallationVictron(SolarInstallation):

    ### CLASS VARIABLES ###
    # connection with victron VRM API
    connection = http.client.HTTPSConnection("vrmapi.victronenergy.com")
    headers = {
            'Content-Type': "application/json",
            'x-authorization': "Token fbb54a457a39c7c00785818b42193c5952ebb3652fe1c3f3ca2da035de524ff3"
            }

    def __init__(self, name, id, report_type):
        super().__init__(name)
        self.id = id

        # Put this in base class with option for None implies 1m
        self.type = "victron energy"
        if report_type not in ['1m', '3m']:
            raise ValueError("Le rapport est soit de 1 mois soit de 3 mois")
        self.report_type = report_type

    def fetch_data(self, url):
        """Fetches data from the given URL using the provided connection and headers."""
        retries = 3
        for retry in range(3):
            try:
                self.connection.request("GET", url, headers=self.headers)
                res = self.connection.getresponse()
                data = res.read()    
                data_str = data.decode("utf-8")
                
                return data_str
                break
            except requests.exceptions.RequestException as e:
                print(f"Request Exception occurred: {e}")
                if retry < retries - 1:
                    print("Retrying...")
                    sleep(1)  # Wait before retrying
                else:
                    print("Maximum retries reached.")
                    break
            except Exception as e:
                print(f"An unexpected error occurred: {e}")
                break
        # Reformat data to be usable



    

    # Get raw data from site
    def reformat_data(self, data):

        lines = data.split("\n")
        lines[-1] = lines[-1].rstrip(",")
        data_fixed = "\n".join(lines)
        return data_fixed



    def get_all_data(self, start, end):

        url = f"/v2/installations/{self.id}/data-download?start={start}&end={end}&datatype=kwh&format=csv&debug=true"
        data =self.fetch_data(url)
        data = self.reformat_data(data)
        
        df1 = pd.DataFrame({'timestamp': pd.date_range(start=pd.to_datetime(start, unit='s'), end=pd.to_datetime(end, unit='s'), freq='15min')})
    #Make better
        df = pd.read_csv(io.StringIO(data))
        df = df.drop(0)
        df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric)
        df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
        df_filled = df.fillna(0)
        df_filled['timestamp'] = pd.to_datetime(df_filled['timestamp'], errors='coerce')
        df_filled = df_filled.set_index('timestamp')
        df_filled['Fonctionnement GE'] = df_filled.apply(lambda row: 15*60 if row['Genset to consumers'] > 0 or row['Genset to battery'] > 0 else 0, axis=1)
        
        df2 = pd.merge(df1,df_filled, on='timestamp', how='left')
        df2 = df2.fillna(0)
        df2 = df2.set_index('timestamp')
        df_grouped_hours = df2.groupby(pd.Grouper(freq='h')).sum()
    # End make better


    # Rename categories in French
        df_grouped_hours['total_solar'] = df_grouped_hours['PV to consumers'] + df_grouped_hours['PV to battery'] + df_grouped_hours['PV to grid'] 
        df_grouped_hours['total_battery'] = df_grouped_hours['Battery to consumers'] + df_grouped_hours['Battery to grid'] 
        df_grouped_hours['total_grid'] = df_grouped_hours['Grid to consumers'] + df_grouped_hours['Grid to battery'] 
        df_grouped_hours['total_gen'] = df_grouped_hours['Genset to consumers'] + df_grouped_hours['Genset to battery']
        df_grouped_hours['total_conso'] = df_grouped_hours['PV to consumers'] + df_grouped_hours['Battery to consumers'] + df_grouped_hours['Grid to consumers'] + df_grouped_hours['Genset to consumers']
        df_grouped_hours['Groupe'] = df_grouped_hours['Genset to consumers']
        df_grouped_hours['Jirama'] = df_grouped_hours['Grid to consumers']
        df_grouped_hours['Batterie'] = df_grouped_hours['Battery to consumers']
        df_grouped_hours['PV'] = df_grouped_hours['PV to consumers']
        
        df_grouped_days = df_grouped_hours.groupby(pd.Grouper(freq='D')).sum()
        df_grouped_days['Fonctionnement GE'] = df_grouped_days.pop('Fonctionnement GE')
        df_grouped_days['Fonctionnement GE'] = pd.to_datetime(df_grouped_days['Fonctionnement GE'], unit='s').dt.strftime('%H:%M:%S')
        
        target_day_sun = df_grouped_days['total_solar'].idxmax()
        day_data_sun = df_grouped_hours[df_grouped_hours.index.date == pd.to_datetime(target_day_sun).date()]
        day_data_sun["PV directement consommée"] = day_data_sun['PV to consumers']
        day_data_sun["PV exporté sur le réseau"] = day_data_sun['PV to grid']
        day_data_sun["PV rechargeant les batteries"] = day_data_sun['PV to battery']
        day_data_sun["Production PV totale"] = day_data_sun['total_solar']
        day_data_sun["Jirama alimentant les charges"] = day_data_sun['Grid to consumers']
        day_data_sun["Groupe alimentant les charges"] = day_data_sun['Genset to consumers']
        day_data_sun["Batterie alimentant les charges"] = day_data_sun['Battery to consumers']
        day_data_sun["Consommation totale"] = day_data_sun['total_conso']
        day_data_sun.pop('Fonctionnement GE')

        target_day_conso = df_grouped_days['total_conso'].idxmax()
        day_data_conso = df_grouped_hours[df_grouped_hours.index.date == pd.to_datetime(target_day_conso).date()]
        day_data_conso["PV directement consommée"] = day_data_conso['PV to consumers']
        day_data_conso["PV exporté sur le réseau"] = day_data_conso['PV to grid']
        day_data_conso["PV rechargeant les batteries"] = day_data_conso['PV to battery']
        day_data_conso["Production PV totale"] = day_data_conso['total_solar']
        day_data_conso["Jirama alimentant les charges"] = day_data_conso['Grid to consumers']
        day_data_conso["Groupe alimentant les charges"] = day_data_conso['Genset to consumers']
        day_data_conso["Batterie alimentant les charges"] = day_data_conso['Battery to consumers']
        day_data_conso["Consommation totale"] = day_data_conso['total_conso']
        day_data_conso.pop('Fonctionnement GE')

    # Return needed data to write into report
        return day_data_sun, day_data_conso, df_grouped_days, target_day_conso, target_day_sun

    def get_soc(self, s, e):
        try:
            # Adjust the start and end times by subtracting 3 hours
            s = s - 3600 * 3
            e = e - 3600 * 3
        
            # Create a DataFrame with hourly timestamps between the adjusted start and end times
            df1 = pd.DataFrame({'timestamp': pd.date_range(start=pd.to_datetime(s, unit='s'), end=pd.to_datetime(e, unit='s'), freq='h')})
        
            # Construct the API URL to fetch SOC data
            type_h = 'venus'
            url_soc = f"/v2/installations/{self.id}/stats?start={s}&end={e}&type={type_h}&interval=hours"
            
            # Send the API request
            data = self.fetch_data(url_soc)
            data_json_value = json.loads(data)
            # Extract the 'bs' data from the JSON response
            value = data_json_value['records']['bs']
            # Define columns and create a DataFrame from the SOC data
            columns = ['timestamp', 'bs_moy', 'bs_min', 'bs_max']
            df_temp = pd.DataFrame(value, columns=columns)

            df = pd.DataFrame()
            df['timestamp'] = pd.to_datetime(df_temp['timestamp'], unit='ms') 
            df['bs_moy'] = df_temp['bs_moy']
            
            # Merge the hourly timestamp DataFrame with the SOC data DataFrame
            df_export = pd.merge(df1, df, on='timestamp', how='left')
            df_export.pop('timestamp')
            # Write the merged DataFrame to an Excel file
            return df_export
            
        except Exception as e:
            # Print any exceptions that occur
            print(f"Une erreur s'est produite : {e}")

    def get_data_previous_month(self, start, end):
        url = f"/v2/installations/{self.id}/data-download?start={start}&end={end}&datatype=kwh&format=csv&debug=true"

        # Get data for the previous month from VRM API
        data = self.reformat_data(self.fetch_data(url))

        # Make base dataframe with timestamp column
        df1 = pd.DataFrame({'timestamp': pd.date_range(start=pd.to_datetime(start, unit='s'), end=pd.to_datetime(end, unit='s'), freq='15min')})

        # Make dataframe with correct timestamps
        df = pd.read_csv(io.StringIO(data))
        df = df.drop(0)
        df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric)
        df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
        df_filled = df.fillna(0)
        df_filled['timestamp'] = pd.to_datetime(df_filled['timestamp'], errors='coerce')
        df_filled = df_filled.set_index('timestamp')

        # Fill dataframe GE data
        df_filled['Fonctionnement GE'] = df_filled.apply(lambda row: 15*60 if row['Genset to consumers'] > 0 or row['Genset to battery'] > 0 else 0, axis=1)
        
        # Merge dataframes according to timestamps
        df2 = pd.merge(df1,df_filled, on='timestamp', how='left')
        df2 = df2.fillna(0)
        df2 = df2.set_index('timestamp')

        # Make new dataframe grouped by hours with correct colmun names        
        df_grouped_hours = df2.groupby(pd.Grouper(freq='h')).sum()
        df_grouped_hours['total_solar'] = df_grouped_hours['PV to consumers'] + df_grouped_hours['PV to battery'] + df_grouped_hours['PV to grid'] 
        df_grouped_hours['total_battery'] = df_grouped_hours['Battery to consumers'] + df_grouped_hours['Battery to grid'] 
        df_grouped_hours['total_grid'] = df_grouped_hours['Grid to consumers'] + df_grouped_hours['Grid to battery'] 
        df_grouped_hours['total_gen'] = df_grouped_hours['Genset to consumers'] + df_grouped_hours['Genset to battery']
        df_grouped_hours['total_conso'] = df_grouped_hours['PV to consumers'] + df_grouped_hours['Battery to consumers'] + df_grouped_hours['Grid to consumers'] + df_grouped_hours['Genset to consumers']
        df_grouped_hours['Groupe'] = df_grouped_hours['Genset to consumers']
        df_grouped_hours['Jirama'] = df_grouped_hours['Grid to consumers']
        df_grouped_hours['Batterie'] = df_grouped_hours['Battery to consumers']
        df_grouped_hours['PV'] = df_grouped_hours['PV to consumers']

        #Make dataframe grouped by days
        df_grouped_days = df_grouped_hours.groupby(pd.Grouper(freq='D')).sum()
        df_grouped_days['Fonctionnement GE'] = df_grouped_days.pop('Fonctionnement GE')
        df_grouped_days['Fonctionnement GE'] = pd.to_datetime(df_grouped_days['Fonctionnement GE'], unit='s').dt.strftime('%H:%M:%S')
        
        # Return dataframe used for report generation
        return df_grouped_days
    
    def get_timestamps(self, records, keys):
        """Processes records to extract specified keys and unique timestamps."""
        data = {key: records.get(key, []) for key in keys}
        timestamps = set()
        for key in keys:
            if isinstance(data[key], list):
                for elem in data[key]:
                    timestamps.add(elem[0])
        return data, timestamps
    
    def build_dataframe(self, timestamps, data_keys):
        """Builds a DataFrame from the given timestamps and data keys."""
        data = []
        for ts in timestamps:
            row = {"Timestamp": ts}
            for key, values in data_keys.items():
                if isinstance(values, (list, tuple)):
                    val = next((val for val in values if val[0] == ts), None)
                    if val:
                        row[key] = val[1]
                    else:
                        row[key] = None
                else:
                    row[key] = values  # Directly assign the value if it's not iterable
            data.append(row)

        df = pd.DataFrame(data)
        return df.sort_values(by='Timestamp', ascending=True)
    
    def get_data_12_months(self, start, end):
        # Construct URL for solar yield
        type_m = 'solar_yield'
        url_12m = f"/v2/installations/{self.id}/stats?start={start}&end={end}&type={type_m}&interval=months"

        # Get data and process it for solar yield to make a dataframe
        data_json_value = json.loads(self.fetch_data(url_12m))
        solar_records = data_json_value['records']
        solar_keys = ['Pc', 'Pb']
        solar_values, solar_timestamps = self.get_timestamps(solar_records, solar_keys)

        df_sy = self.build_dataframe(solar_timestamps, solar_values)

        # Construct URL for live feed
        type_m = 'live_feed'
        url_12m = f"/v2/installations/{self.id}/stats?start={start}&end={end}&type={type_m}&interval=months"

        # Get data and process it for solar yield
        data_json_value = json.loads(self.fetch_data(url_12m))
        live_records = data_json_value['records']
        live_keys = ['total_consumption', 'total_genset', 'grid_history_from']
        live_values, live_timestamps = self.get_timestamps(live_records, live_keys)

        all_timestamps = solar_timestamps.union(live_timestamps)
        df_lf = self.build_dataframe(all_timestamps, live_values)

        # Merge dataframes
        df_temp = pd.merge(df_sy, df_lf, on='Timestamp', how='left')

        # Create final dataframe
        df = pd.DataFrame()
        df['Date'] = pd.to_datetime(df_temp['Timestamp'] + 3600000*3, unit='ms')
        df['PV consommée'] = df_temp['Pc']
        df['Production PV totale'] = df_temp['Pc'] + df_temp['Pb']
        df['Jirama'] = df_temp['grid_history_from']
        df['Production Groupe'] = df_temp['total_genset']
        df['Consommation totale'] = df_temp['total_consumption']

        return df
    
    def fit_model(self, df, columns):
        regressor = LinearRegression()
        for column in columns:
            X = df['Timestamp'].values.reshape(-1, 1)
            y = df[column].values
            regressor.fit(X, y)
            df[f"{column}_pred"] = regressor.predict(X)
        return df

    def detect_peaks(self, values, mean_value):
        """Detect local maxima and minima that are greater/less than the mean value."""
        max_peak_indices = np.where((values[1:-1] > values[:-2]) & (values[1:-1] > values[2:]) & (values[1:-1] > mean_value))[0] + 1
        min_peak_indices = np.where((values[1:-1] < values[:-2]) & (values[1:-1] < values[2:]) & (values[1:-1] < mean_value))[0] + 1
        return max_peak_indices, min_peak_indices

    def calculate_z_scores(self, peaks, mean_value, std_dev):
        """Calculate Z-scores for the detected peaks."""
        return (peaks - mean_value) / std_dev

    def identify_anomalies(self, peaks, peak_indices, mean_value, threshold, feature):
        z_scores = self.calculate_z_scores(peaks, mean_value, peaks.std())
        df_peaks = pd.DataFrame()
        df_peaks['Timestamp'] = peak_indices
        df_peaks[feature] = peaks
        df_peaks['Z_score'] = z_scores
        return df_peaks[abs(df_peaks['Z_score']) > threshold]

    def plot_data(self, df, feature_col, df_export):
        """Plot the original and predicted data, highlighting anomalies."""
        plt.plot(df['Timestamp'], df[feature_col], label=f'{feature_col}')
        plt.plot(df['Timestamp'], df[f'{feature_col}_pred'], label=f'{feature_col}_pred')
        plt.scatter(df_export['Timestamp'], df_export[feature_col], color='red')
        plt.legend()
        plt.show()

    def get_and_analyze_data(self, start, end, feature_col):
        """Fetch, process, analyze data, and write results to an Excel sheet."""
        url = f"/v2/installations/{self.id}/stats?start={start}&end={end}&type=live_feed"

        data = json.loads(self.fetch_data(url))['records'][feature_col]
        if feature_col == "bv":
            feature_col = [ "BV_moy", "BV_min", "BV_max"]
        elif feature_col == 'total_solar_yield':
            feature_col = ["SY"]
        columns = ["Timestamp"]
        columns+=feature_col
        df = pd.DataFrame(data, columns=columns)

        col_no_timestamp = columns[1:]
        df = self.fit_model(df, col_no_timestamp)

        values = df[columns[1]].values
        mean_value = values.mean()
        max_peak_indices, min_peak_indices = self.detect_peaks(values, mean_value)
        max_peaks = values[max_peak_indices]
        min_peaks = values[min_peak_indices]

        max_peaks_moy = max_peaks.mean()
        min_peaks_moy = min_peaks.mean()

        threshold = 0.5
        df_export = pd.DataFrame()
        # Initial population of df_export with anomalies
        anomalies_max = self.identify_anomalies(max_peaks, max_peak_indices, max_peaks_moy, threshold, columns[1])
        anomalies_min = self.identify_anomalies(min_peaks, min_peak_indices, min_peaks_moy, threshold, columns[1])

        df_export = pd.concat([anomalies_max, anomalies_min], axis=0).sort_values(by='Timestamp', ascending=True)

        while (len(df_export) > 5) and (threshold <= 2):
            anomalies_max = self.identify_anomalies(max_peaks, max_peak_indices, max_peaks_moy, threshold, columns[1])
            anomalies_min = self.identify_anomalies(min_peaks, min_peak_indices, min_peaks_moy, threshold, columns[1])

            df_export = pd.concat([anomalies_max, anomalies_min], axis=0).sort_values(by='Timestamp', ascending=True)
            #self.plot_data(df, feature_col, df_export)

            threshold += 0.05
        # Attempt to convert 'Timestamp' to datetime format
        try:
            df_export['Timestamp'] = pd.to_datetime(df_export['Timestamp'], unit='ms')

        except Exception as e:
            print(f"Error during conversion: {e}")
        return df_export

    def get_and_analyze_bv_and_sy(self, start, end):
        bv_df = self.get_and_analyze_data(start, end, 'bv')
        sy_df = self. get_and_analyze_data(start, end, 'total_solar_yield')
        return bv_df, sy_df
    


    def process_alarm_data(self, alarm_data, alarm_meta, time_format):
        processed_data = {}
        for main_key, sub_dict in alarm_data.items():
            inner_dict = {}
            for sub_key, values in sub_dict.items():
                if int(values['0']) > 0:
                    list_value = list(values.values())
                    list_value[1] = datetime.fromtimestamp(list_value[1]).strftime(time_format)
                    list_value[2] = datetime.fromtimestamp(list_value[2]).strftime(time_format)
                    time1 = datetime.strptime(list_value[1], time_format)
                    time2 = datetime.strptime(list_value[2], time_format)
                    list_value.append(time2 - time1)
                    inner_dict[sub_key] = list_value
            processed_data[main_key] = inner_dict
        return processed_data
    
    def prepare_dataframe(self, alarm_data, alarm_meta):
        df = pd.DataFrame(data=alarm_data)
        temp_ser_list = []
        for key in df:
            columns_name = [f'{key} Error Class', f'{key} Start', f'{key} End', f'{key} Duration']
            temp_ser = pd.DataFrame(columns=columns_name)
            for kk in range(len(df)):
                if isinstance(df[key].iloc[kk], list):
                    temp_ser.loc[len(temp_ser)] = df[key][kk]
                else:
                    temp_ser.loc[len(temp_ser)] = [np.nan] * 4
            temp_ser_list.append(temp_ser)
        if temp_ser_list:
            return pd.concat(temp_ser_list, axis=1)
        else:
            return pd.DataFrame()

    def summarize_alarms(self, df):
        count_w = []
        count_alarm = []
        tag_list = []
        count_duration = []

        for key in df.columns:
            key_str = str(key)
            if 'Error Class' in key_str:
                tag_list.append(''.join(filter(str.isdigit, key)))
                count_w.append((df[key] == '1').sum())
                count_alarm.append((df[key] == '2').sum())
            if 'Duration' in key_str:
                count = df[key].sum()
                if isinstance(count, timedelta):
                    days = count.days
                    hours, remainder = divmod(count.seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)
                    formatted_time = f"{days} jours, {hours:02d}:{minutes:02d}:{seconds:02d}"
                else:
                    formatted_time = '00 00:00:00'
                count_duration.append(formatted_time)
            else:
                print(f"Unexpected key type: {type(key)} with value {key}")
        summary = pd.DataFrame(tag_list)
        summary['Nombre Warning'] = count_w
        summary['Nombre Alarmes'] = count_alarm
        summary['Durée cumulée'] = count_duration
        return summary[summary['Durée cumulée'] != '00 00:00:00']

### NEEDS TO BE TRANSFERED TO PLOT FILE ? ###
    def generate_plots(self, data_graph, ve_meta, start, end, time_format, time_cond):
        plot_files = []
        for alarm in ve_meta[0]:
            if alarm in data_graph['alarm'].values:
                alarm_data = data_graph[data_graph['alarm'] == alarm].copy()
                alarm_data['colors'] = alarm_data['data_class'].map({1: 'C1', 2: 'C3'})
                time_diff = alarm_data['end'] - alarm_data['start']
                valid_data = alarm_data[time_diff > time_cond].copy()
                plt.barh(valid_data['data_class'], valid_data['end'] - valid_data['start'],
                        left=valid_data['start'], color=valid_data['colors'],
                        edgecolor='k')
                plt.title(f"Alarme {alarm} - {valid_data['alarm_title'].iloc[0]}")
                plt.gca().xaxis.set_major_formatter(mdates.DateFormatter(time_format))
                plt.xticks(rotation=45)
                plt.xlim(mdates.date2num(datetime.fromtimestamp(start)), mdates.date2num(datetime.fromtimestamp(end)))
                plt.tick_params(axis='x', which='both', direction='out', pad=15)
                plt.yticks([1, 2], ['Warning', 'Alarme'])
                plot_file = f"plot_{self.id}_{alarm}.png"
                plt.savefig(plot_file, dpi=300, bbox_inches='tight')
                plt.close()
                plot_files.append(plot_file)
        return plot_files


    def get_alarms(self, start, end):
        time_format = "%d-%m-%Y %H:%M:%S"
        time_cond = timedelta(minutes=2)
        instances = {
            "ve": [276, 512],
            "inv": [20, 21, 22]
        }
        l_df_meta = []
        l_df_summary = []

        for category, instance_list in instances.items():
            for id_instance in instance_list:
                count = 0
                success = False
                while not success and count < 2:
                    
                    try:
                        if category == "ve":
                            url = f"/v2/installations/{self.id}/widgets/VeBusWarningsAndAlarms?instance={id_instance}&start={start}&end={end}"
                        else:
                            url = f"/v2/installations/{self.id}/widgets/InverterChargerWarningsAndAlarms?instance={id_instance}&start={start}&end={end}"
                        
                        data = self.fetch_data(url)
                        alarm_data = json.loads(data)['records']['data']
                        alarm_meta = json.loads(data)['records']['meta']
                        alarm_meta_df = pd.DataFrame([[key, value['description']] for key, value in alarm_meta.items()])
                        processed_data = self.process_alarm_data(alarm_data, alarm_meta, time_format)
                        alarm_df = self.prepare_dataframe(processed_data, alarm_meta_df)
                        
                        if not alarm_df.empty:
                            alarm_df = pd.concat([pd.DataFrame(alarm_meta_df), alarm_df], axis=1)
                        else:
                            alarm_df = pd.DataFrame(alarm_meta_df)
                        alarm_summary = self.summarize_alarms(alarm_df)
                        l_df_meta.append(alarm_meta_df)
                        l_df_summary.append(alarm_summary)
                        data_graph = []
                        for alarm, sub_dict in processed_data.items():
                            for data_class, values in sub_dict.items():
                                data_graph.append({
                                    'alarm': alarm,
                                    'data_class': int(values[0]),
                                    'start': pd.to_datetime(values[1], format=time_format),
                                    'end': pd.to_datetime(values[2], format=time_format)
                                })
                        
                        data_graph_df = pd.DataFrame(data_graph)
                        data_graph_df = data_graph_df.merge(alarm_meta_df.rename(columns={0: 'alarm', 1: 'alarm_title'}), on='alarm')
                        #plot_files = self.generate_plots(data_graph_df, alarm_meta_df, self.id, start, end, time_format, time_cond)
                        plot_files="not done yet"
                        success = True
                    
                    except Exception as e:
                        print(f"fail {e}")
                        count += 1
        df_meta = pd.concat(l_df_meta, axis=0, ignore_index=True).drop_duplicates()
        df_alarm_summary = pd.concat(l_df_summary, axis=0, ignore_index=True)

        return df_meta, df_alarm_summary, plot_files




# get site list, with names and phone numbers    
def get_site_list():
    connection = http.client.HTTPSConnection("vrmapi.victronenergy.com")
    headers = {
            'Content-Type': "application/json",
            'x-authorization': "Token fbb54a457a39c7c00785818b42193c5952ebb3652fe1c3f3ca2da035de524ff3"
            }
    connection.request("GET", "/v2/users/280026/installations", headers=headers)
    result = connection.getresponse()
    data = result.read()
    data = data.decode("utf-8")
    
    data_json_site = json.loads(data)
    site = data_json_site['records']
    site_list =[]
    name_list =[]
    phonenumber_list =[]

    for i in site:
        site_list.append(i['idSite'])
        name_list.append(i['name'])
        phonenumber_list.append(i['phonenumber'])

    return site_list, name_list, phonenumber_list