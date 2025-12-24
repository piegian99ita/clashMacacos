import asyncio
import os
import coc

from ruamel.yaml import YAML
from datetime import datetime, timedelta

# --- CONFIGURAZIONE ---
COC_EMAIL = os.environ.get("COC_EMAIL", "ilredeitrattori@gmail.com")
COC_PASSWORD = os.environ.get("COC_PASSWORD", "3499Pg00")
CLAN_TAG = os.environ.get("CLAN_TAG", "2JPPL0922")

# !!! IMPORTANTE: Questo deve essere il nome ESATTO del file nel tuo repo !!!
WORKFLOW_FILENAME = ".github/workflows/war_script.yaml" 

def aggiorna_cron_workflow(day, hour, minute):
    nuovo_cron = f"{minute} {hour} {day} * *"
    
    yaml = YAML()
    yaml.preserve_quotes = True
    yaml.indent(mapping=2, sequence=4, offset=2)

    try:
        with open(WORKFLOW_FILENAME, 'r') as f:
            data = yaml.load(f)
    except FileNotFoundError:
        print(f"ERRORE: Non trovo il file {WORKFLOW_FILENAME}. Controlla il nome!")
        return

    # Inizializza le chiavi se mancano (evita l'errore NoneType)
    if not data: data = {}
    if 'on' not in data: data['on'] = {}
    
    # Gestione sicura della lista schedule
    if 'schedule' not in data['on'] or not isinstance(data['on']['schedule'], list):
        data['on']['schedule'] = [{'cron': nuovo_cron}]
    else:
        data['on']['schedule'][0]['cron'] = nuovo_cron

    with open(WORKFLOW_FILENAME, 'w') as f:
        yaml.dump(data, f)
    
    print(f"✅ Cron aggiornato con successo a: {nuovo_cron}")

async def check_data():
    async with coc.Client() as client:
        try:
            print("Login in corso...")
            await client.login(COC_EMAIL, COC_PASSWORD)
            
            print(f"Recupero dati war per: {CLAN_TAG}")
            clan_war = await client.get_current_war(CLAN_TAG)
            
            # Se siamo in CWL o non c'è war, non facciamo nulla al cron (o gestiscilo diversamente)
            if clan_war.is_cwl:
                print("Siamo in CWL. Nessuna modifica al cron standard.")
                return
            
            if clan_war.state == "warEnded":
                print("La war è finita. Attendo la prossima.")
                return

            if clan_war.end_time:
                # clan_war.end_time.time è un oggetto datetime UTC
                # Sottraiamo 15 minuti per eseguire lo script poco prima della fine
                final_time = clan_war.end_time.time - timedelta(minutes=15)
                
                print(f"La war finisce il {clan_war.end_time.time}")
                print(f"Imposto il prossimo cron per: Giorno {final_time.day} alle {final_time.hour}:{final_time.minute} UTC")
                
                aggiorna_cron_workflow(final_time.day, final_time.hour, final_time.minute)
            else:
                print("Nessun orario di fine disponibile.")

        except coc.errors.InvalidCredentials:
            print("❌ Credenziali non valide!")
        except Exception as e:
            print(f"❌ Errore generale: {e}")

if __name__ == "__main__":
    asyncio.run(check_data())