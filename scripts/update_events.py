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
spreadsheet_id = os.environ['SPREADSHEET_ID_EVENTS']
spreadsheet = client.open_by_key(spreadsheet_id)

# 3. Leggi il file Excel con tutti i suoi fogli
file_path = 'cc_cg_events.xlsx'  # Assicurati che il nome coincida con quello generato
excel_data = pd.ExcelFile(file_path)
# ... (parte iniziale identica)

for sheet_name in excel_data.sheet_names:
    print(f"Elaborazione foglio: {sheet_name}...")
    
    # Leggi specificando il motore openpyxl
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    
    # Sostituisce NaN con stringa vuota, ma mantiene i tipi numerici dove possibile
    df = df.fillna('') 
    
    # Converte i dati in una lista di liste (formato richiesto da gspread)
    # .astype(str) può essere utile se hai problemi di formattazione, 
    # ma rimuovilo se vuoi che i numeri rimangano numeri su Google Sheets
    header = df.columns.tolist()
    values = df.values.tolist()
    data_to_upload = [header] + values
    
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        print(f"Creata nuova tab: {sheet_name}")

    worksheet.clear()
    
    # Utilizzo di value_input_option='USER_ENTERED' 
    # Fondamentale: permette a Google Sheets di interpretare formule e date
    spreadsheet.values_update(
        f"'{sheet_name}'!A1",
        params={'valueInputOption': 'USER_ENTERED'},
        body={'values': data_to_upload}
    )
    print(f"Tab '{sheet_name}' aggiornata con successo!")