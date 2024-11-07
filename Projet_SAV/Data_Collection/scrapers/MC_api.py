# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 10:59:21 2024

@author: elino
"""

import requests
from flask import Flask, jsonify, request

app = Flask(__name__)

# URL de base de l'API Meteocontrol
BASE_URL = "https://www1.meteocontrol.de/vcom/evaluation/index/index/systemId/2125533"
LOGIN_URL = "https://www1.meteocontrol.de/vcom/login"

# Informations d'authentification
credentials = {
    'login': 'MDG_maintenance',
    'password': 'MGPAdmin2022!'
}

# Session pour maintenir l'authentification
session = requests.Session()

# Authentification avec des en-têtes supplémentaires
def authenticate():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
        'Referer': 'https://www1.meteocontrol.de/vcom/login'
    }
    response = session.post(LOGIN_URL, data=credentials, headers=headers)
    if response.ok:
        print("Authentification réussie")
    else:
        print("Échec de l'authentification:", response.text)


@app.route('/api/energie', methods=['GET'])
def get_energie():
    input_date = request.args.get('date')
    key = "RPP09"  # Clé pour récupérer les données d'énergie
    url = f"{BASE_URL}?key={key}&date={input_date}T00%3A00%3A00%2B03%3A00&endDate={input_date}T23%3A59%3A59%2B03%3A00"

    try:
        response = session.get(url)
        response.raise_for_status()  # Vérifier si la requête a réussi
        data = response.json()  # Suppose que les données sont en JSON
        return jsonify(data), 200
    except requests.exceptions.RequestException as e:
        print(f"Erreur de requête : {e}")  # Loguer l'erreur de requête
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        print(f"Erreur inattendue : {e}")  # Loguer l'erreur inattendue
        return jsonify({"error": "Une erreur s'est produite: " + str(e)}), 500

if __name__ == '__main__':
    authenticate()  # Appel de la fonction d'authentification
    try:
        print("Démarrage du serveur Flask sur le port 5000...")
        app.run(port=5000, debug=False)  # Utilisation du port 5001
    except SystemExit as e:
        print(f"Erreur lors du démarrage du serveur : {e}")
    except Exception as e:
        print(f"Erreur inattendue : {e}")
