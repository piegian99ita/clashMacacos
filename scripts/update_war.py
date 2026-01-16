import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

# 1. Configurazione Credenziali
json_creds = json.loads(os.environ['GCP_SERVICE_ACCOUNT_KEY'])
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(json_creds, scope)
client = gspread.authorize(creds)

# 2. Apri il foglio Google (Spreadsheet principale)
spreadsheet_id = os.environ['SPREADSHEET_ID_WAR']
spreadsheet = client.open_by_key(spreadsheet_id)

# 3. Leggi il file Excel con tutti i suoi fogli
file_path = 'rewards.xlsx'  # Assicurati che il nome coincida con quello generato
excel_data = pd.ExcelFile(file_path)

for sheet_name in excel_data.sheet_names:
    print(f"Elaborazione foglio: {sheet_name}...")
    
    # Leggi i dati dello sheet corrente
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    df = df.fillna('') # Sostituisce i valori NaN con stringhe vuote
    
    # Prepara i dati (Header + Righe)
    data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
    
    try:
        # Prova ad aprire la tab esistente
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # Se non esiste, crea una nuova tab
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        print(f"Creata nuova tab: {sheet_name}")

    # Pulisce la tab esistente e carica i nuovi dati
    worksheet.clear()
    worksheet.update(data_to_upload)
    print(f"Tab '{sheet_name}' aggiornata con successo!")

print("Aggiornamento completo di tutti i fogli!")