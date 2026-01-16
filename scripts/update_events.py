import gspread
import openpyxl
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

# 1. Configurazione Credenziali
json_creds = json.loads(os.environ['GCP_SERVICE_ACCOUNT_KEY'])
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(json_creds, scope)
client = gspread.authorize(creds)

# 2. Apri il foglio Google
spreadsheet_id = os.environ['SPREADSHEET_ID_EVENTS']
spreadsheet = client.open_by_key(spreadsheet_id)

# 3. Impostazioni file Excel
file_path = 'cc_cg_events.xlsx'

print(f"Apertura file Excel: {file_path}")

# Carichiamo il workbook con data_only=False. 
# Questo è il TRUCCO: dice alla libreria di leggere la formula (es. =SUM(A1:A5)) 
# invece del risultato (che sarebbe None o 0).
wb = openpyxl.load_workbook(file_path, data_only=False)

for sheet_name in wb.sheetnames:
    print(f"Elaborazione foglio: {sheet_name}...")
    
    ws = wb[sheet_name]
    
    # Estraiamo i dati riga per riga trasformandoli in una lista di liste
    data_to_upload = []
    
    for row in ws.iter_rows():
        row_data = []
        for cell in row:
            # cell.value qui restituisce la stringa della formula (es. "=A1+B1")
            # oppure il valore grezzo se non c'è formula.
            val = cell.value
            
            # Gestione dei valori nulli e conversione date/numeri complessi
            if val is None:
                val = ""
            # Convertiamo tutto in stringa per sicurezza, tranne le formule
            # Google Sheets interpreterà i numeri e le date automaticamente
            # se la stringa non inizia con '='
            
            row_data.append(val)
        data_to_upload.append(row_data)

    # 4. Caricamento su Google Sheets
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        print(f"Creata nuova tab: {sheet_name}")

    worksheet.clear()
    
    # USIAMO 'USER_ENTERED':
    # Questo dice a Google: "Fai finta che un utente stia digitando questi dati".
    # Se arriva "=SUM(A1:B1)", Google lo calcolerà.
    spreadsheet.values_update(
        f"'{sheet_name}'!A1",
        params={'valueInputOption': 'USER_ENTERED'},
        body={'values': data_to_upload}
    )
    
    print(f"Tab '{sheet_name}' aggiornata e formule attivate!")

print("Aggiornamento completo riuscito.")