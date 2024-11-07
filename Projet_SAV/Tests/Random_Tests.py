import requests
import http.client
import os
import json
import datetime
from time import * 
from datetime import *
import pandas as pd
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

pd.set_option('future.no_silent_downcasting', True)

pd.options.mode.chained_assignment = None  # disable the warning
ssl._create_default_https_context = ssl._create_unverified_context

# path for the master file
path_src = "G:/.shortcut-targets-by-id/1FQfz_wNk7M-PeQeUyVAHoy9ay4UY1_62/3 - SAV/6 - Surveillance quotidienne/3 - Outil optimisation/Scripts Report Generator/"
path_master = path_src + "#Master_Report Generator.xlsx"

# connection with victron VRM API
conn = http.client.HTTPSConnection("vrmapi.victronenergy.com")
headers = {
        'Content-Type': "application/json",
        'x-authorization': "Token fbb54a457a39c7c00785818b42193c5952ebb3652fe1c3f3ca2da035de524ff3"
        }


# get site list, with names and phone numbers    
conn.request("GET", "/v2/users/280026/installations", headers=headers)
res = conn.getresponse()
data = res.read()
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

print(name_list.index("Centre Anani"))
print(site_list[24])