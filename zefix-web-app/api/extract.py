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
                'legalForms': ['0106', '0107', '0108'],  # SA, Sàrl, EI
                'page': 0,
                'pageSize': 100
            }
            
            try:
                response = requests.get(api_url, params=params, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    
                    for item in data.get('list', []):
                        # Extraire les informations
                        entreprise = {
                            'nom': item.get('name', ''),
                            'forme_juridique': self.get_forme_juridique(item.get('legalForm', '')),
                            'canton': canton,
                            'ville': item.get('city', ''),
                            'npa': item.get('zipCode', ''),
                            'adresse': item.get('address', ''),
                            'date_inscription': item.get('registrationDate', ''),
                            'uid': item.get('uid', ''),
                            'numero_rc': item.get('registerNumber', '')
                        }
                        
                        entreprises.append(entreprise)
                        
            except Exception as e:
                print(f"Erreur canton {canton}: {e}")
                continue
        
        # Trier par priorité
        entreprises = self.prioritize(entreprises)
        
        return entreprises
    
    def get_forme_juridique(self, code):
        """Convertir le code en forme juridique"""
        mapping = {
            '0106': 'SA',
            '0107': 'Sàrl',
            '0108': 'EI'
        }
        return mapping.get(code, 'Autre')
    
    def prioritize(self, entreprises):
        """Prioriser par forme juridique"""
        priority_map = {'SA': 1, 'Sàrl': 2, 'EI': 3}
        
        return sorted(entreprises, key=lambda x: priority_map.get(x['forme_juridique'], 99))
    
    def create_excel(self, entreprises):
        """Créer le fichier Excel"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Nouvelles Entreprises"
        
        # Couleurs
        colors = {
            'SA': 'FFCCCC',
            'Sàrl': 'FFE5CC',
            'EI': 'FFFFCC'
        }
        
        # En-têtes
        headers = [
            'Priorité', 'Nom entreprise', 'Forme juridique', 'Canton', 
            'Ville', 'NPA', 'Adresse', 'Date publication', 
            'Téléphone', 'Email', 'Site web', 'LinkedIn', 
            'Statut', 'Notes', 'Date dernier contact', 'Numéro RC', 'UID'
        ]
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color='FFFFFF', size=11)
            cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Données
        for row_idx, ent in enumerate(entreprises, 2):
            # Priorité
            priorite = 'Haute' if ent['forme_juridique'] == 'SA' else 'Moyenne' if ent['forme_juridique'] == 'Sàrl' else 'Basse'
            
            row_data = [
                priorite,
                ent['nom'],
                ent['forme_juridique'],
                ent['canton'],
                ent['ville'],
                ent['npa'],
                ent['adresse'],
                ent['date_inscription'],
                '',  # Téléphone
                '',  # Email
                '',  # Site web
                '',  # LinkedIn
                'Nouveau',  # Statut
                '',  # Notes
                '',  # Date dernier contact
                ent['numero_rc'],
                ent['uid']
            ]
            
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                
                # Couleur priorité
                if col_idx == 1 and priorite in ['Haute', 'Moyenne', 'Basse']:
                    color_map = {'Haute': 'FFCCCC', 'Moyenne': 'FFE5CC', 'Basse': 'FFFFCC'}
                    cell.fill = PatternFill(start_color=color_map[priorite], 
                                           end_color=color_map[priorite], 
                                           fill_type='solid')
                    cell.font = Font(bold=True)
        
        # Largeurs colonnes
        column_widths = {
            'A': 12, 'B': 40, 'C': 15, 'D': 10, 'E': 20, 
            'F': 8, 'G': 35, 'H': 15, 'I': 18, 'J': 30, 
            'K': 35, 'L': 35, 'M': 12, 'N': 40, 'O': 18, 
            'P': 20, 'Q': 20
        }
        
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # Figer + filtres
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        
        # Sauvegarder en mémoire
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output.read()
