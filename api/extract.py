"""
API Serverless pour extraire les données ZEFIX
Compatible avec Vercel Serverless Functions
"""

from http.server import BaseHTTPRequestHandler
import json
import requests
from datetime import datetime, timedelta
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import base64


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            # Lire le body
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            cantons = data.get('cantons', ['GE', 'VD'])
            days = data.get('days', 7)
            
            # Extraire les données
            entreprises = self.extract_zefix(cantons, days)
            
            # Créer le fichier Excel
            excel_data = self.create_excel(entreprises)
            
            # Encoder en base64 pour le retour
            excel_b64 = base64.b64encode(excel_data).decode('utf-8')
            
            # Préparer la réponse
            response = {
                'success': True,
                'count': len(entreprises),
                'filename': f'Nouvelles_Entreprises_{datetime.now().strftime("%Y-%m-%d")}.xlsx',
                'data': excel_b64
            }
            
            # Envoyer la réponse
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            self.wfile.write(json.dumps(response).encode('utf-8'))
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            error_response = {
                'success': False,
                'error': str(e)
            }
            
            self.wfile.write(json.dumps(error_response).encode('utf-8'))
    
    def do_OPTIONS(self):
        """Handle CORS preflight"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def extract_zefix(self, cantons, days):
        """Extraire les données depuis l'API ZEFIX"""
        entreprises = []
        
        # Date de début
        date_from = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        # API ZEFIX publique
        api_url = "https://www.zefix.admin.ch/ZefixPublicREST/api/v1/shab/search"
        
        for canton in cantons:
            # Recherche par canton
            params = {
                'canton': canton,
                'registrationDateFrom': date_from,
                'legalForms': ['0106', '0107', '0108'],
                'page': 0,
                'pageSize': 100
            }
            
            try:
                response = requests.get(api_url, params=params, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    
                    for item in data.get('list', []):
                        entreprise = {
                            'nom': item.get('name', ''),
                            'forme_juridique': self.get_forme_juridique(item.get('legalForm', '')),
                            'canton': canton,
                            'ville': item.get('city', ''),
                            'n<span class="cursor">█</span>
