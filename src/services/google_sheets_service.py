from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os
import sys

def resource_path(relative_path):
    """Obter o caminho absoluto para o recurso, funciona para dev e PyInstaller"""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho nela
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = resource_path(os.path.join('credentials', 'service-account-file.json'))

def authenticate_google_sheets():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

def save_data_to_sheet(spreadsheet_id, range_name, values):
    service = authenticate_google_sheets()
    body = {
        'values': values
    }
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    return result.get('updates').get('updatedCells')