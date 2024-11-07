import requests
import pandas as pd

def fetch_data(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"An error occurred during the GET request: {e}")
        return None

def process_data(json_data, m_pv_gis):
    # Load the JSON data into a DataFrame
    df = pd.DataFrame(json_data)
    output_hourly = pd.DataFrame(df['outputs']['hourly'])
    
    # Convert the time column to datetime format and create new columns
    output_hourly['time'] = pd.to_datetime(output_hourly['time'], format="%Y%m%d:%H%M")
    output_hourly['total_irradiance'] = output_hourly[['Gb(i)', 'Gd(i)', 'Gr(i)']].sum(axis=1)
    
    # Group by day and calculate the sum of Irradiance and Energy (Wh) for each day
    df_to_be_grouped = output_hourly[['time', 'total_irradiance', 'P']].copy()
    df_to_be_grouped.columns = ['Time', 'Irradiance', 'Energie (Wh)']
    df_to_be_grouped.set_index('Time', inplace=True)
    grouped_by_day = df_to_be_grouped.resample('D').sum().reset_index()
    
    # Extract day and month from the Time column
    grouped_by_day['Day'] = grouped_by_day['Time'].dt.day
    grouped_by_day['Month'] = grouped_by_day['Time'].dt.month
    
    # Calculate daily mean for the specified month
    daily_mean_df_temp = grouped_by_day.groupby(['Month', 'Day']).mean().reset_index()
    daily_mean_df_temp['Date'] = pd.to_datetime(daily_mean_df_temp[['Month', 'Day']].assign(year=2000))
    daily_mean_df_parsed = daily_mean_df_temp[daily_mean_df_temp['Month'] == m_pv_gis]
    
    # Create the final DataFrame
    daily_mean_df = daily_mean_df_parsed[['Date', 'Irradiance', 'Energie (Wh)']].copy()
    
    return daily_mean_df, daily_mean_df_temp


def get_irradiance_pv_gis(m_pv_gis, lat, lon, pvtechchoice, peakpower, loss, mountingplace, angle, azimut, startyear, endyear):
    url = f"https://re.jrc.ec.europa.eu/api/v5_2/seriescalc?lat={lat}&lon={lon}&raddatabase=PVGIS-SARAH2&browser=1&outputformat=json&userhorizon=&usehorizon=1&angle={angle}&aspect={azimut}&startyear={startyear}&endyear={endyear}&mountingplace={mountingplace}&optimalinclination=0&optimalangles=0&js=1&select_database_hourly=PVGIS-SARAH2&hstartyear={startyear}&hendyear={endyear}&trackingtype=0&hourlyangle={angle}&hourlyaspect={azimut}&pvcalculation=1&pvtechchoice={pvtechchoice}&peakpower={peakpower}&loss={loss}&components=1"
    print(url)

    json_data = fetch_data(url)
    if json_data:
        daily_mean_df, monthly_data = process_data(json_data, m_pv_gis)
    return daily_mean_df, monthly_data