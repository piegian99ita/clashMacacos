import asyncio
import os
import json
import coc

# --- CONFIGURAZIONE ---

COC_EMAIL = os.getenv("COC_EMAIL")
COC_PASSWORD = os.getenv("COC_PASSWORD")
CLAN_TAG = os.getenv("CLAN_TAG")

async def recupera_punti_player(client, member_tag):
    """
    Recupera i punti del 'Games Champion' (Giochi del Clan) per un singolo giocatore.
    """
    try:
        # Recupera il profilo completo del giocatore
        player = await client.get_player(member_tag)
        
        # In coc.py gli achievements sono oggetti. Cerchiamo 'Games Champion'
        games_champion = next((a for a in player.achievements if a.name == "Games Champion"), None)
        
        points = games_champion.value if games_champion else 0
        return {"name": player.name, "points": points}
    
    except Exception as e:
        print(f"Errore per il tag {member_tag}: {e}")
        return {"name": "Sconosciuto", "points": "Errore"}

async def esporta_dati():
    # 1. Inizializzazione Client
    async with coc.Client(key_names="PC_Locale_Key") as client:
        try:
            await client.login(COC_EMAIL, COC_PASSWORD)
            
            # 2. Estrazione Membri del Clan
            print(f"Recupero membri per il clan: {CLAN_TAG}...")
            clan = await client.get_clan(CLAN_TAG)
            
            # 3. Recupero parallelo dei dati (Sostituisce ThreadPoolExecutor)
            print(f"Inizio scaricamento dati per {len(clan.members)} membri...\n")
            
            # Creiamo una lista di 'tasks' (compiti da svolgere)
            tasks = [recupera_punti_player(client, m.tag) for m in clan.members]
            
            # asyncio.gather esegue tutte le chiamate simultaneamente
            risultati = await asyncio.gather(*tasks)
            
            print("--- Download Completato ---")
            return risultati

        except coc.errors.InvalidCredentials:
            print("Credenziali non valide!")
            return []
        except Exception as e:
            print(f"Errore generale: {e}")
            return []

