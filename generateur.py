import pandas as pd
import json
import os
import re
import requests
import time
import math
import hashlib  # Pour détecter les changements de lignes

# =====================================================================
# CONFIGURATION AUTOMATIQUE
# =====================================================================
nom_fichier_excel = "00_Mes vinyles.xlsx"
nom_fichier_wanted = "01_Liste achat.xlsx"
fichier_cache = "vinyles_cache.json"  # Fichier mémoire pour accélérer le script

CONSUMER_KEY = "pTrgAPVrGOUYZbQrFbbh"
CONSUMER_SECRET = "XnsdSqnEoZQHJLjtEhVwffLFSYNJTYmV"
# =====================================================================

if not os.path.exists("pochettes"):
    os.makedirs("pochettes")

# Chargement du cache s'il existe
cache_donnees = {}
if os.path.exists(fichier_cache):
    try:
        with open(fichier_cache, 'r', encoding='utf-8') as f:
            cache_donnees = json.load(f)
        print(f"💾 Cache chargé : {len(cache_donnees)} vinyles en mémoire.")
    except Exception as e:
        print(f"⚠️ Impossible de lire le cache, il sera réinitialisé : {e}")

def generer_hash_ligne(row):
    """Crée une empreinte unique de la ligne pour détecter le moindre changement (ex: modification de prix, de titre...)"""
    chaine_complete = "".join(str(val) for val in row.values)
    return hashlib.md5(chaine_complete.encode('utf-8')).hexdigest()

def extraire_id_discogs(url):
    if pd.isna(url):
        return None
    match = re.search(r"release/(\d+)", str(url))
    return match.group(1) if match else None

def telecharger_pochette(release_id, artiste, titre):
    if not release_id:
        return "pochettes/placeholder.png"
    
    chemin_image = f"pochettes/{release_id}.jpg"
    
    if os.path.exists(chemin_image) and os.path.getsize(chemin_image) > 0:
        return chemin_image
        
    try:
        url_api = f"https://api.discogs.com/releases/{release_id}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Authorization": f"Discogs key={CONSUMER_KEY}, secret={CONSUMER_SECRET}"
        }
        reponse = requests.get(url_api, headers=headers)
        
        if reponse.status_code == 200:
            images = reponse.json().get('images', [])
            if images:
                url_image = images[0].get('uri') or images[0].get('resource_url')
                if url_image:
                    img_req = requests.get(url_image, headers=headers)
                    if img_req.status_code == 200:
                        with open(chemin_image, 'wb') as handler:
                            handler.write(img_req.content)
                        print(f"📸 Image récupérée : {artiste} - {titre}")
                        time.sleep(1.5)
                        return chemin_image
            
            print(f"⚠️ Pas d'image trouvée sur Discogs pour l'ID {release_id} ({artiste} - {titre})")
            
        elif reponse.status_code == 429:
            print("🛑 Discogs bloque temporairement (trop de requêtes). Pause de 10 secondes...")
            time.sleep(10)
        else:
            print(f"❌ Erreur API Discogs (Code {reponse.status_code}) pour l'ID {release_id} ({artiste} - {titre})")
            
    except Exception as e:
        print(f"⚠️ Impossible de vérifier l'ID {release_id} ({artiste}) : {e}")
        
    return "pochettes/placeholder.png"

def recuperer_prix_haut(release_id):
    if not release_id:
        return None
    try:
        url_api = f"https://api.discogs.com/marketplace/price_suggestions/{release_id}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Authorization": f"Discogs key={CONSUMER_KEY}, secret={CONSUMER_SECRET}"
        }
        reponse = requests.get(url_api, headers=headers)
        
        if reponse.status_code == 200:
            data = reponse.json()
            prix_brut = None
            if "Mint (M)" in data:
                prix_brut = data["Mint (M)"]["value"]
            elif "Near Mint (NM or M-)" in data:
                prix_brut = data["Near Mint (NM or M-)"]["value"]
            elif data.values():
                prix_brut = max(item["value"] for item in data.values() if isinstance(item, dict) and "value" in item)
            
            if prix_brut:
                return f"{math.ceil(prix_brut)}€"
        elif reponse.status_code == 429:
            time.sleep(3)
    except:
        pass
    return None

# =====================================================================
# TRAITEMENT DU FICHIER PRINCIPAL (MES VINYLES)
# =====================================================================
try:
    df_brut = pd.read_excel(nom_fichier_excel, header=None)
    valeur_a1 = df_brut.iloc[0, 0]
    total_vinyles = str(int(float(valeur_a1))) if not pd.isna(valeur_a1) else "Non spécifié"
    df = pd.read_excel(nom_fichier_excel)
except Exception as e:
    print(f"Impossible de lire le fichier principal Excel : {e}")
    exit()

colonne_type = next((c for c in df.columns if any(k in str(c).lower() for k in ['album', 'compil', 'sing'])), df.columns[2])
colonne_titre_a = next((c for c in df.columns if 'titre' in str(c).lower() and any(k in str(c).lower() for k in ['face a', 'album'])), None)
if not colonne_titre_a:
    colonne_titre_a = next((c for c in df.columns if 'titre' in str(c).lower() or 'album' in str(c).lower()), df.columns[1])
colonne_lien = next((c for c in df.columns if 'lien' in str(c).lower()), 'Lien')

colonne_genre = next((c for c in df.columns if 'genre' in str(c).lower()), None)
if not colonne_genre and len(df.columns) > 6:
    colonne_genre = df.columns[6]

collection = []
nouveau_cache = {}
liste_genres_uniques = set()
liste_types_uniques = set() 

print(f"\nAnalyse de la collection en cours... (Total global : {total_vinyles})")

compteur_api = 0

for index, row in df.iterrows():
    artiste = str(row.get('ARTISTE', '')).strip()
    if not artiste or artiste == 'nan':
        continue
    
    # Calcul de l'empreinte de la ligne
    hash_ligne = generer_hash_ligne(row)
    
    # Si la ligne existe à l'identique dans l'ancien cache, on reprend les données sans appeler l'API
    if hash_ligne in cache_donnees:
        viny_data = cache_donnees[hash_ligne]
        nouveau_cache[hash_ligne] = viny_data
        collection.append(viny_data)
        
        # On extrait quand même le genre et le type pour les filtres de la page
        if viny_data["type"]: liste_types_uniques.add(viny_data["type"])
        if viny_data["genre"] and viny_data["genre"] != "N/C": liste_genres_uniques.add(viny_data["genre"].upper())
        continue

    # SINON : C'est une nouvelle ligne ou elle a été modifiée -> Appel API
    compteur_api += 1
    titre_a = str(row.get(colonne_titre_a, '')).strip() if colonne_titre_a else ""
    if titre_a == 'nan': titre_a = ""
        
    lien_discogs = row.get(colonne_lien, '#')
    id_discogs = extraire_id_discogs(lien_discogs