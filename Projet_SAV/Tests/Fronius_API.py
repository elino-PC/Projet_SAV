import requests
import http.client
import os
import json


# Replace with your demo API key
api_key = 'YOUR_DEMO_API_KEY'
headers = {
    'accept': 'application/json',
    'Authorization': f'Bearer {api_key}'
}

# Example endpoint for demo data
url = "https://api.solarweb.com/v1/DemoEndpoint"

response = requests.get(url, headers=headers)
data = response.json()