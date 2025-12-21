import json
import requests
import concurrent.futures

def esporta_dati():

    # --- CONFIGURAZIONE E CARICAMENTO DATI ---
    nome_file = "./secret.json"
    with open(nome_file, 'r', encoding='utf-8') as file:
        dati = json.load(file)

    token = dati['token']
    clan_tag = dati['clan_tag']
    HEADERS = {'Authorization': f'Bearer {token}'}

    # --- 1. ESTRAZIONE MEMBRI ---
    URL_API = f"https://api.clashofclans.com/v1/clans/%23{clan_tag}/members"
    risposta_coc = requests.get(URL_API, headers=HEADERS)

    lista_membri_pulita = []
    try:
        risposta_coc.raise_for_status()
        dati_clan = risposta_coc.json()
        items = dati_clan.get('items', [])
        
        for m in items:
            # Puliamo il tag rimuovendo l'hash per l'URL successivo
            lista_membri_pulita.append({
                "name": m.get("name"),
                "tag": m.get("tag").lstrip('#')
            })
    except Exception as e:
        print(f"Errore recupero clan: {e}")
        exit()

    # --- 2. FUNZIONE DA ESEGUIRE IN PARALLELO ---
    def recupera_punti_player(member):
        """Esegue la chiamata API per il singolo giocatore"""
        tag = member["tag"]
        name = member["name"]
        url = f"https://api.clashofclans.com/v1/players/%23{tag}"
        
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            r.raise_for_status()
            dati_player = r.json()
            
            achievements = dati_player.get("achievements", [])
            punti_giochi = 0
            for ach in achievements:
                if ach["name"] == "Games Champion":
                    punti_giochi = ach["value"]
                    break
            
            #print(f"Scaricato: {name}") # Feedback visivo
            return {"name": name, "points": punti_giochi}
        
        except Exception as e:
            print(f"Errore per {name}: {e}")
            return {"name": name, "points": "Errore"}

    # --- 3. APPLICAZIONE THREADING ---
    player_data = []

    print(f"\nInizio scaricamento dati per {len(lista_membri_pulita)} membri...\n")

    # Usiamo ThreadPoolExecutor per gestire le chiamate simultanee
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        # Mappa la funzione sulla lista dei membri
        risultati = list(executor.map(recupera_punti_player, lista_membri_pulita))

    # Filtriamo i risultati (rimuovendo eventuali errori se necessario)
    player_data = risultati

    print("\n--- Ritorno ai points ---")
    
    return player_data



